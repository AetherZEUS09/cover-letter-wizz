# ⚡ Cover Letter Wizz

Automate your job applications. Generate hundreds of unique, tailored cover letter PDFs in seconds.

![Cover Letter Wizz Banner](https://img.shields.io/badge/status-active-brightgreen) ![Python](https://img.shields.io/badge/python-3.10%2B-blue) ![Streamlit](https://img.shields.io/badge/streamlit-1.30%2B-red)

---

## What It Does

Cover Letter Wizz takes a CSV list of companies and job roles and generates a unique cover letter for each one. It:

- Rotates through a library of resume-based body templates so no two companies get the same letter.
- Injects the company name and job role directly into PDF or DOCX templates.
- Uses **PyMuPDF** to read, redact, and write text directly onto PDF templates without needing Microsoft Word.
- Sanitizes folder names and detects duplicates to prevent overwrites.
- Outputs a clean ZIP file organized by company and role.

---

## Features

| Feature | Description |
|---|---|
| **Unique Bodies** | Rotating library of pre-written cover letter bodies, each 500+ words. |
| **Custom Templates** | Upload your own PDF (from Canva) or DOCX template. |
| **Direct PDF Editing** | Reads and stamps text directly onto Canva PDFs using PyMuPDF. No Word required. |
| **Duplicate Detection** | Skips exact duplicate (company + role) pairs so you don't generate the same letter twice. |
| **Multi-Role Support** | A company with multiple different roles gets one letter per role. |
| **Folder Sanitization** | Strips slashes and illegal characters so folder names never break the ZIP. |
| **Teal Vector UI** | Clean, techy Streamlit interface with a dark teal theme. |
| **Zero Config** | No database. No API. Runs entirely on your machine. |

---

## Requirements

- **Python 3.10 or higher**
- **Windows, macOS, or Linux**
- *(Optional)* **LibreOffice** — Only required if you upload **DOCX** templates. If you use PDF templates from Canva, you do not need LibreOffice.

---

## Installation

### Step 1: Install Python
Download and install Python from [python.org](https://www.python.org/downloads/). **Make sure to check "Add Python to PATH"** during installation.

### Step 2: (Optional) Install LibreOffice
If you plan to upload **Word (.docx)** templates, download LibreOffice from [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/). If you only plan to upload Canva PDFs, skip this step.

### Step 3: Clone the Repository

Open Command Prompt (Windows) or Terminal (macOS/Linux) and run:

```bash
git clone https://github.com/YOUR_USERNAME/cover-letter-wizz.git
cd cover-letter-wizz
```

### Step 4: Install Python Dependencies

```bash
pip install streamlit pandas python-docx PyMuPDF openpyxl
```

Here is what each package does:

| Package | Purpose |
|---|---|
| `streamlit` | The web UI framework |
| `pandas` | Reads the CSV/Excel job list |
| `python-docx` | Handles DOCX templates (if used) |
| `PyMuPDF` | **Reads, redacts, and edits PDF templates** |
| `openpyxl` | Reads `.xlsx` Excel files |

---

# 🚀 The Beginner Walkthrough

Before you run the app, you need three things:
1. A **Cover Letter Template** (created in Canva as a PDF).
2. A **Jobs List** (Excel or CSV).
3. A **ZIP File** containing the Jobs List.

Here is exactly how to make all three.

### Part 1: Creating the Template (Using Canva)

You will design your cover letter in Canva. Because the app uses **PyMuPDF** to edit PDFs directly, we do not need Word or a fancy converter.

1. Open **Canva** and create a new "Document" (US Letter or A4).
2. Design your cover letter exactly how you want it to look (name, contact info, lines, signature).
3. **DO NOT** type your actual letter body or "Dear Hiring Team". 
4. Instead, type these exact variable names where you want the software to insert the text:
   - Type `{{job role}}` at the top where the job title goes.
   - Type `{{company name}}` where the company name goes.
   - Type `{{body}}` where the main paragraphs of the letter go.
5. **IMPORTANT:** Do NOT include a greeting line like "Dear Hiring Team," in your Canva design. The software automatically adds the greeting for you.
6. Click **Share** → **Download** → choose **PDF Standard**.
7. Save this file to your computer (e.g., `MyTemplate.pdf`).

### Part 2: Creating the Jobs List (Excel or CSV)

The software needs a list of companies and the roles you are applying to. 

1. Open **Microsoft Excel**, **Google Sheets**, or **WPS Office**.
2. You only need two columns. **Do not use headers**. Just start typing on the very first row.
3. **Column A** must be the **Company Name**.
4. **Column B** must be the **Job Role**.
5. Example:
   | A | B |
   |---|---|
   | Google | Data Engineer |
   | Nike | Machine Learning Engineer |
   | Fidelity | Data Architect |
   | Google | Software Engineer |
6. Save the file as a **CSV (Comma Delimited)** or **Excel Workbook (.xlsx)**. Name it `jobs.csv`.

*(Note: If a company has two different roles, put them on two separate rows. The software will generate two different letters. If you put the exact same company and exact same role twice, the software will skip the duplicate).*

### Part 3: Zipping the Jobs List (How to Upload)

The app requires the jobs list to be inside a **ZIP file** before you upload it.

1. Locate your `jobs.csv` or `jobs.xlsx` file on your computer.
2. **Right-click** the file.
3. If you are on Windows 11, click **"Compress to ZIP file"**.
4. If you are on Windows 10, click **"Send to"** → **"Compressed (zipped) folder"**.
5. If you are on Mac, right-click and choose **"Compress jobs.csv"**.
6. You will now have a file called `jobs.zip` (or `jobs.csv.zip`).

---

# 💻 Running the App & Generating Letters

### Step 1: Launch the App

Open your Command Prompt (Windows) or Terminal (macOS/Linux), navigate to your project folder, and run:

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

### Step 2: Upload and Generate

1. **Upload your template:** Click the first upload button and select your Canva PDF (e.g., `MyTemplate.pdf`). The app will show a green ✅ if it found your `{{body}}` and `{{company name}}` placeholders.
2. **Upload your jobs list:** Click the second upload button and select your `jobs.zip` file.
3. **Click "⚡ Generate Batch".**
4. The app will display a preview of any duplicates it found (which it will skip) and any companies that have multiple roles (which it will keep).
5. Wait for the green progress bars to finish.
6. Click the **"📥 Download All Cover Letters (ZIP)"** button.

### Step 3: What You Get

The downloaded ZIP file contains neatly organized folders. 

```
Cover_Letters_Batch.zip
├── Google - Data Engineer/
│   └── Data Engineer_Cover_Letter.pdf
├── Google - Software Engineer/
│   └── Software Engineer_Cover_Letter.pdf
├── Nike - Machine Learning Engineer/
│   └── Machine Learning Engineer_Cover_Letter.pdf
└── Fidelity - Data Architect/
    └── Data Architect_Cover_Letter.pdf
```

---

## Customizing the Bodies

To add new cover letter bodies, edit `bodies.py` and add a new triple-quoted string to the `BODIES` list:

```python
BODIES = [
    """Dear Hiring Team at {{company name}},\n\nYour opening paragraph...""",
    """Dear {{company name}} Recruiters,\n\nAnother unique body...""",
]
```

Each body should:
- Start with a greeting line using `{{company name}}`
- Use `\n\n` for paragraph breaks
- Use `{{job role}}` and `{{company name}}` throughout
- Be at least 500 words

---

## Troubleshooting

**LibreOffice not found (Only if using DOCX)**  
If you upload a DOCX template, the app needs LibreOffice to convert it to PDF. Look for the line in `app.py` that says `soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"`. Update it if your installation path is different. **If you use PDF templates from Canva, you do not need this.**

**Letters generated but body is blank**  
Your PDF template is missing the `{{body}}` variable, or the text is combined with other elements in Canva (like a group). Make sure `{{body}}` is its own separate text box in Canva.

**Fonts in the body look different**  
Because the app uses **PyMuPDF** to stamp the body text onto the PDF, it uses a standard clean font (Helvetica). The header (`{{job role}} at {{company name}}`) uses the exact same font matching logic.

**Folder names look weird**  
Slash characters like `/` in job titles (e.g., "AI / ML Engineer") are replaced with `_` in folder names to prevent the ZIP file from breaking. The letter content is unchanged.

**Fewer folders than CSV rows**  
This means duplicate (company + role) pairs were found and skipped. Check the duplicate report shown in the app before generation.

---

## License

MIT License. Free to use, modify, and share.

---

## Author

Built by **S. ZEUS Macaulay**.

This tool is used to remotely generate tailored cover letters for clients.

Want to contribute or report a bug? [Open an issue](https://github.com/AetherZEUS/cover-letter-wizz/issues).
