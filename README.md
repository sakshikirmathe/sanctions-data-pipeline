# EU Sanctions Data Pipeline
A fully automated pipeline to extract, normalize, and analyze European Union travel-ban sanctions data by combining official XML feeds and PDF documents into a single structured Excel output.

This project is designed for AML, compliance, and risk-intelligence workflows, where accurate, structured sanctions data is required for screening, investigations, and reporting.

---

## What this project does
The pipeline performs the following steps end-to-end:

1. Downloads EU sanctions data  
   - XML feed containing sanctioned entities  
   - Official PDF document containing detailed travel-ban information  

2. Parses and normalizes data  
   - Splits large XML files into structured entity records  
   - Extracts and chunks text from the PDF  
   - Aligns entities across XML and PDF sources  

3. Enriches and maps data  
   - Matches XML entities to corresponding PDF entries  
   - Extracts personal and entity-level details  
   - Performs gender inference using name-based dictionaries  

4. Generates analyst-ready output  
   - Produces a clean Excel file containing all matched and enriched sanctions records  

The result is a single Excel file that can be used directly for:
- Sanctions screening  
- AML investigations  
- Compliance reporting  
- Risk analysis  

---

## Output
After a successful run, the pipeline creates a `data/` folder containing:
```
data/
├── xml_files/ # Raw EU XML feed
├── xml_chunks/ # Parsed XML entity records
├── pdf/ # Official EU sanctions PDF
├── pdf_text_chunks/ # Extracted & chunked PDF text
└── sanctions_output.xlsx # Final structured output
```
The main deliverable is: *data/sanctions_output.xlsx*

This file contains all matched and enriched sanctions entities.

---

## How to run
### Option 1 — Run with Python

1. Clone the repository  
2. Create and activate a virtual environment  
3. Install dependencies - pip install -r requirements.txt
4. Run the pipeline - python main.py

 
---

### Option 2 — Run as a Windows Executable

This project supports being packaged as a standalone `.exe` using PyInstaller.

The executable:
- Includes Chromium for Playwright  
- Does not require Python on the target machine  
- Writes output to a `data/` folder next to the EXE  

To build the executable: pyinstaller main.spec
The EXE will be created in: dist/main.exe

Run it by double-clicking or from the command line.

---

## Technologies Used

- Python  
- Playwright (Chromium)  
- Pandas  
- PDFPlumber  
- XML parsing  
- gender-guesser  
- PyInstaller  

---

## Use cases

This tool is suitable for:

- AML & KYC teams  
- Financial crime & compliance analysts  
- Sanctions screening pipelines  
- Risk and regulatory reporting  
- Data engineering workflows in fintech or consulting  

---

## Disclaimer

This project is for data processing and analysis purposes only.  
It does not provide legal or regulatory advice.  
Users are responsible for ensuring compliance with applicable laws and regulations when using sanctions data.

Author - Sakshi Kirmathe
