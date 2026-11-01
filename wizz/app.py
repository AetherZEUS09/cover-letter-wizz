import streamlit as st
import pandas as pd
import zipfile
import os
import io
import subprocess
import random
import re
from docx import Document
from docx.shared import Pt
import fitz  # PyMuPDF

from bodies import BODIES

# ============================================================
# TECHY TEAL UI
# ============================================================
st.set_page_config(page_title="Cover Letter Wizz", page_icon="⚡", layout="wide")

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Space+Grotesk:wght@400;500;600;700&display=swap');
.stApp {
    background:
        radial-gradient(circle at 15% 10%, rgba(20, 184, 166, 0.15) 0%, transparent 45%),
        radial-gradient(circle at 85% 90%, rgba(34, 211, 238, 0.12) 0%, transparent 50%),
        radial-gradient(circle at 50% 50%, #0a2429 0%, #061418 60%, #030a0d 100%);
    background-attachment: fixed;
    font-family: 'Space Grotesk', -apple-system, sans-serif;
    color: #ccfbf1;
}
.stApp::before {
    content: "";
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background-image:
        linear-gradient(rgba(20, 184, 166, 0.05) 1px, transparent 1px),
        linear-gradient(90deg, rgba(20, 184, 166, 0.05) 1px, transparent 1px);
    background-size: 40px 40px;
    pointer-events: none;
    z-index: 0;
}
header[data-testid="stHeader"] {
    background: rgba(6, 20, 24, 0.7);
    backdrop-filter: blur(14px);
    border-bottom: 1px solid rgba(20, 184, 166, 0.2);
}
h1 {
    background: linear-gradient(90deg, #14b8a6 0%, #22d3ee 50%, #5eead4 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    letter-spacing: -1.2px;
    font-size: 3rem !important;
}
h2, h3 { color: #5eead4 !important; font-family: 'Space Grotesk', sans-serif; letter-spacing: -0.5px; }
.stMarkdown p, .stText, label, .stCaption { color: #99f6e4 !important; }
.stButton > button {
    background: linear-gradient(135deg, #0d9488 0%, #14b8a6 50%, #06b6d4 100%);
    color: #ffffff;
    border: 1px solid rgba(94, 234, 212, 0.4);
    border-radius: 10px;
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 600;
    letter-spacing: 0.5px;
    padding: 0.6rem 1.5rem;
    box-shadow: 0 0 12px rgba(20, 184, 166, 0.35), 0 0 24px rgba(34, 211, 238, 0.15);
    transition: all 0.25s ease;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 0 20px rgba(20, 184, 166, 0.6), 0 0 40px rgba(34, 211, 238, 0.3);
    border-color: #5eead4;
}
[data-testid="stFileUploader"] {
    background: rgba(13, 148, 136, 0.06);
    border: 1px dashed rgba(94, 234, 212, 0.35);
    border-radius: 12px;
    padding: 12px;
}
.stAlert { border-radius: 10px; border-left: 3px solid #14b8a6; background: rgba(13, 148, 136, 0.12); color: #ccfbf1; }
code, pre { background: rgba(6, 20, 24, 0.7) !important; color: #5eead4 !important; font-family: 'JetBrains Mono', monospace !important; border-radius: 6px; }
hr { border-color: rgba(20, 184, 166, 0.25); }
.stDownloadButton > button {
    background: linear-gradient(135deg, #134e4a 0%, #0d9488 100%);
    color: #ccfbf1;
    border: 1px solid #14b8a6;
    border-radius: 10px;
    font-weight: 600;
    padding: 0.6rem 1.5rem;
    box-shadow: 0 0 15px rgba(20, 184, 166, 0.4);
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ============================================================
# FOLDER NAME SANITIZER
# ============================================================
def sanitize_for_folder(name):
    if not name:
        return "unknown"
    name = name.replace("/", "_").replace("\\", "_")
    for ch in [':', '*', '?', '"', '<', '>', '|']:
        name = name.replace(ch, "")
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:100] if name else "unknown"

# ============================================================
# PDF HELPERS
# ============================================================
def detect_pdf_variables(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = set()
    pattern = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}')
    for page in doc:
        text = page.get_text()
        for match in pattern.finditer(text):
            found.add(match.group(0).strip())
    doc.close()
    return sorted(found)

def detect_docx_variables(doc):
    found = set()
    pattern = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}')
    def scan(text):
        for m in pattern.finditer(text):
            found.add(m.group(0).strip())
    for p in doc.paragraphs: scan(p.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: scan(p.text)
    return sorted(found)

def fill_pdf_template(pdf_bytes, job_role, company_name, body_text):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    resolved_body = body_text.replace('{{job role}}', job_role).replace('{{company name}}', company_name)
    for page in doc:
        job_areas = page.search_for("{{job role}}")
        comp_areas = page.search_for("{{company name}}")
        body_areas = page.search_for("{{body}}")
        combined_pairs = []
        used_job = set(); used_comp = set()
        for i, jr in enumerate(job_areas):
            for j, cr in enumerate(comp_areas):
                if j in used_comp: continue
                if abs(jr.y0 - cr.y0) < 5:
                    combined_pairs.append((jr, cr))
                    used_job.add(i); used_comp.add(j)
                    break
        header_fontsize = 13
        if combined_pairs:
            sample = combined_pairs[0][0]
            header_fontsize = max(10, min(14, sample.height - 2))
        for jr, cr in combined_pairs:
            whole_rect = fitz.Rect(min(jr.x0, cr.x0) - 1, min(jr.y0, cr.y0) - 1,
                                   max(jr.x1, cr.x1) + 1, max(jr.y1, cr.y1) + 1)
            page.add_redact_annot(whole_rect, fill=(1, 1, 1))
        for i, jr in enumerate(job_areas):
            if i in used_job: continue
            page.add_redact_annot(jr, fill=(1, 1, 1))
        for j, cr in enumerate(comp_areas):
            if j in used_comp: continue
            page.add_redact_annot(cr, fill=(1, 1, 1))
        for rect in body_areas:
            page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()
        for jr, cr in combined_pairs:
            left_x = min(jr.x0, cr.x0)
            baseline_y = max(jr.y1, cr.y1) - 3
            page.insert_text(fitz.Point(left_x, baseline_y), f"{job_role} at {company_name}",
                             fontsize=header_fontsize, fontname="helv", color=(0, 0, 0))
        for i, jr in enumerate(job_areas):
            if i in used_job: continue
            page.insert_text(fitz.Point(jr.x0, jr.y1 - 3), job_role,
                             fontsize=header_fontsize, fontname="helv", color=(0, 0, 0))
        for j, cr in enumerate(comp_areas):
            if j in used_comp: continue
            page.insert_text(fitz.Point(cr.x0, cr.y1 - 3), company_name,
                             fontsize=header_fontsize, fontname="helv", color=(0, 0, 0))
        if body_areas:
            body_rect = body_areas[0]
            body_box = fitz.Rect(body_rect.x0, body_rect.y0,
                                 page.rect.width - 40, page.rect.height - 60)
            leftover = page.insert_textbox(body_box, resolved_body, fontsize=10.5,
                                           fontname="helv", color=(0, 0, 0), align=0)
            if isinstance(leftover, str) and leftover.strip():
                extra_page = doc.new_page()
                extra_box = fitz.Rect(40, 40, extra_page.rect.width - 40, extra_page.rect.height - 60)
                extra_page.insert_textbox(extra_box, leftover, fontsize=10.5,
                                          fontname="helv", color=(0, 0, 0), align=0)
    result = doc.tobytes()
    doc.close()
    return result

def fill_docx_template(template_bytes, job_role, company_name, body_text):
    doc = Document(io.BytesIO(template_bytes))
    resolved_body = body_text.replace('{{job role}}', job_role).replace('{{company name}}', company_name)
    body_paragraphs = resolved_body.split('\n\n')
    body_para = None
    for p in doc.paragraphs:
        if '{{body}}' in p.text:
            body_para = p
            break
    if body_para is not None:
        orig_font = body_para.runs[0].font.name if body_para.runs else None
        orig_size = body_para.runs[0].font.size if body_para.runs else None
        current_elem = body_para._element
        for para_text in body_paragraphs:
            new_p = doc.add_paragraph()
            run = new_p.add_run(para_text)
            if orig_font: run.font.name = orig_font
            if orig_size: run.font.size = orig_size
            current_elem.addnext(new_p._element)
            current_elem = new_p._element
        body_para._element.getparent().remove(body_para._element)
    for p in doc.paragraphs:
        if '{{job role}}' in p.text:
            p.text = p.text.replace('{{job role}}', job_role)
        if '{{company name}}' in p.text:
            p.text = p.text.replace('{{company name}}', company_name)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if '{{job role}}' in p.text:
                        p.text = p.text.replace('{{job role}}', job_role)
                    if '{{company name}}' in p.text:
                        p.text = p.text.replace('{{company name}}', company_name)
    return doc

# ============================================================
# THE APP
# ============================================================
st.markdown("# ⚡ Cover Letter Wizz")
st.markdown(
    f"<p style='color:#5eead4; font-family: JetBrains Mono, monospace; "
    f"font-size: 0.95rem; letter-spacing: 0.5px; margin-top: -15px;'>"
    f"// {len(BODIES)} bodies · duplicate preview · full diagnostics"
    f"</p>",
    unsafe_allow_html=True
)
st.markdown("<hr style='border-color: rgba(20,184,166,0.25); margin: 1.5rem 0;'>", unsafe_allow_html=True)

st.markdown("### ⟩ Step 01 · Upload Template")
uploaded_template = st.file_uploader("Upload your template (PDF or DOCX)", type=["pdf", "docx"])

template_type = None
detected_vars = []
template_bytes = None

if uploaded_template is not None:
    template_bytes = uploaded_template.getbuffer()
    filename_lower = uploaded_template.name.lower()
    if filename_lower.endswith('.pdf'):
        template_type = 'pdf'
        try:
            detected_vars = detect_pdf_variables(template_bytes)
        except Exception as e:
            st.error(f"Could not read the PDF: {e}")
    else:
        template_type = 'docx'
        try:
            doc = Document(uploaded_template)
            detected_vars = detect_docx_variables(doc)
        except Exception as e:
            st.error(f"Could not read the DOCX: {e}")
    if detected_vars:
        st.success(f"✅ Found {len(detected_vars)} variable(s):")
        for v in detected_vars:
            st.write(f"  • `{v}`")

st.markdown("### ⟩ Step 02 · Upload Jobs File")
uploaded_zip = st.file_uploader("Upload ZIP (containing .csv or .xlsx — Col A = Company, Col B = Role)", type=["zip"])

st.markdown("### ⟩ Step 03 · Generate")
if st.button("⚡ Generate Batch"):
    if uploaded_zip is None:
        st.warning("Please upload your ZIP file.")
    elif uploaded_template is None:
        st.warning("Please upload your template.")
    else:
        with st.spinner("⚡ Generating unique cover letters..."):
            with zipfile.ZipFile(uploaded_zip, 'r') as z:
                data_files = [f for f in z.namelist() if f.endswith(('.csv', '.xlsx', '.xls'))]
                if not data_files:
                    st.error("No CSV or Excel file found!")
                    st.stop()
                filename = data_files[0]
                with z.open(filename) as f:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(f, header=None)
                    else:
                        df = pd.read_excel(f, header=None)
                    if len(df.columns) >= 2:
                        df = df.iloc[:, :2]
                        first_cell = str(df.iloc[0, 0]).lower()
                        second_cell = str(df.iloc[0, 1]).lower()
                        if 'company' in first_cell or 'role' in second_cell or 'job' in second_cell:
                            df = df.iloc[1:]

            total_rows = len(df)
            st.info(f"📊 CSV contains **{total_rows} rows**.")

            df['_company_clean'] = df.iloc[:, 0].astype(str).str.strip()
            df['_role_clean'] = df.iloc[:, 1].astype(str).str.strip()

            # ---- DUPLICATE DETECTION (before removal) ----
            duplicates_mask = df.duplicated(subset=['_company_clean', '_role_clean'], keep='first')
            duplicate_rows = df[duplicates_mask][['_company_clean', '_role_clean']].values.tolist()
            duplicates_removed = len(duplicate_rows)

            if duplicates_removed > 0:
                st.warning(f"⚠️ **{duplicates_removed} exact duplicates found and will be SKIPPED:**")
                for c, r in duplicate_rows:
                    st.write(f"  • ❌ `{c}` / `{r}` (already exists above)")
            else:
                st.success("✅ No exact duplicates found.")

            # ---- APPLY DEDUPLICATION ----
            before = len(df)
            df = df.drop_duplicates(subset=['_company_clean', '_role_clean'], keep='first')
            after = len(df)

            if duplicates_removed > 0:
                st.info(f"📋 Generating **{after}** unique letters from {before} rows.")
            else:
                st.info(f"📋 Generating **{after}** unique letters.")

            # ---- MULTI-ROLE COMPANIES ----
            multi_role = df.groupby('_company_clean')['_role_clean'].nunique()
            multi_role_companies = multi_role[multi_role > 1].to_dict()
            if multi_role_companies:
                st.info(f"🔄 **{len(multi_role_companies)} companies have multiple roles** (all kept):")
                for c, count in multi_role_companies.items():
                    st.write(f"  • **{c}** → {count} letters")

            body_order = list(range(len(BODIES)))
            random.shuffle(body_order)

            output_zip_buffer = io.BytesIO()
            generated_count = 0
            failed_rows = []
            written_folders = set()
            soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

            with zipfile.ZipFile(output_zip_buffer, 'w') as out_zip:
                for index, row in df.iterrows():
                    company_raw = str(row['_company_clean'])
                    role_raw = str(row['_role_clean'])
                    if not company_raw or not role_raw or company_raw == 'nan' or role_raw == 'nan':
                        failed_rows.append((company_raw, role_raw, "Empty field"))
                        continue

                    company = sanitize_for_folder(company_raw)
                    role = sanitize_for_folder(role_raw)

                    base_folder = f"{company} - {role}"
                    folder_path = base_folder
                    suffix = 2
                    while folder_path in written_folders:
                        folder_path = f"{base_folder} ({suffix})"
                        suffix += 1

                    chosen_body = BODIES[body_order[generated_count % len(body_order)]]
                    pdf_filename = f"{role}_Cover_Letter.pdf"

                    try:
                        if template_type == 'pdf':
                            pdf_bytes = fill_pdf_template(bytes(template_bytes), role_raw, company_raw, chosen_body)
                            out_zip.writestr(f"{folder_path}/{pdf_filename}", pdf_bytes)
                        else:
                            doc = fill_docx_template(bytes(template_bytes), role_raw, company_raw, chosen_body)
                            temp_docx_path = f"temp_{index}.docx"
                            doc.save(temp_docx_path)
                            subprocess.run([soffice_path, "--headless", "--convert-to", "pdf",
                                            "--outdir", ".", temp_docx_path], check=True, timeout=60)
                            generated_pdf_path = temp_docx_path.replace(".docx", ".pdf")
                            with open(generated_pdf_path, "rb") as pdf_file:
                                pdf_bytes = pdf_file.read()
                            out_zip.writestr(f"{folder_path}/{pdf_filename}", pdf_bytes)
                            os.remove(generated_pdf_path)
                            os.remove(temp_docx_path)

                        written_folders.add(folder_path)
                        generated_count += 1
                        st.write(f"✅ [{generated_count}/{after}] {role_raw} at {company_raw}")
                    except Exception as e:
                        failed_rows.append((company_raw, role_raw, str(e)))
                        st.error(f"❌ Failed: {company_raw} / {role_raw} → {e}")

            st.markdown("---")
            st.markdown(f"### ✅ Final Report")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("CSV Rows", total_rows)
            col2.metric("Duplicates Skipped", duplicates_removed)
            col3.metric("Generated", generated_count)
            col4.metric("Failed", len(failed_rows))

            if len(failed_rows) > 0:
                st.error(f"### ❌ {len(failed_rows)} failed rows")
                for c, r, err in failed_rows[:20]:
                    st.write(f"  • `{c}` / `{r}` → {err}")

            if generated_count > 0:
                st.download_button(
                    label=f"📥 Download {generated_count} Cover Letters (ZIP)",
                    data=output_zip_buffer.getvalue(),
                    file_name="Cover_Letters_Batch.zip",
                    mime="application/zip"
                )