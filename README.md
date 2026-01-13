# Sanctions Data Pipeline (EU Travel Ban)

A full-stack data extraction & transformation pipeline that:

• Scrapes official EU Sanctions data  
• Downloads XML + PDF directly from SanctionsMap  
• Splits entities into structured files  
• Extracts identity, nationality, DOB, aliases, and violations  
• Matches XML entities with PDF references  
• Produces a clean Excel file ready for compliance teams  

This simulates how Financial Crime, AML, and Sanctions teams process raw regulatory data.

---

## 🚀 What this project does

1. Connects to EU SanctionsMap using Playwright  
2. Downloads:
   - Official XML export  
   - Official PDF sanction list  
3. Splits XML into one file per entity  
4. Extracts PDF text into entity blocks  
5. Matches XML names to PDF references  
6. Builds a clean Excel workbook with:
   - Name
   - Gender
   - DOB
   - Nationality
   - Address
   - Aliases
   - Violation numbers
   - Programme info

This is the same workflow used in:
• AML teams  
• Sanctions screening engines  
• Watchlist data vendors  

---

## 🧠 Why this is valuable

This is not a toy scraper.  
It demonstrates:

• Web automation (Playwright)  
• XML parsing  
• PDF text extraction  
• Data normalization  
• Entity resolution  
• Excel automation  
• Real-world regulatory data engineering  

This is exactly what FinTech & Compliance data teams do.

---

## 🛠 Tech Stack

- Python  
- Playwright  
- Requests  
- Pandas  
- PDFPlumber  
- OpenPyXL  
- Regex  

---

## ▶ How to Run

```bash
pip install -r requirements.txt
playwright install
python main.py
