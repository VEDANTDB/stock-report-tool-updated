# Stock Report Tool

Generates a professional equity research report from screener.in data.
No API key required.

## Input Files
| File | Required? | Where to get it |
|------|-----------|-----------------|
| Company.xlsx | ✅ Mandatory | screener.in → search company → Export to Excel |
| AnnualReport.pdf | Optional | Company website → Investor Relations |
| Transcript.pdf | Optional | BSE/NSE filings or company website |

## One-Time Setup
Install Python: https://python.org
Install Node.js: https://nodejs.org (pick LTS version)

Then run:
    pip install -r requirements.txt
    npm install

## How to Run
    python3 run_report.py --xlsx YourCompany.xlsx

With PDFs:
    python3 run_report.py --xlsx Company.xlsx --pdf1 Annual.pdf --pdf2 Transcript.pdf

## Output
A fully formatted Word document (.docx) with:
- Financial tables (annual, quarterly, balance sheet, cash flow)
- 8-dimension financial health score
- BUY/SELL/HOLD ratings for 6M, 1Y, 3Y, 5Y horizons
- Valuation via P/E, EV/EBITDA, and DCF
- Key positives, risks, analyst views, monitorables
```

4. Click **"Commit changes"**

---

## Part 6: Your Final Repo Should Look Like This
```
stock-report-tool/
├── run_report.py
├── extract_data.py
├── ai_analysis.py
├── generate_report.js
├── requirements.txt
├── package.json
├── .gitignore
├── .env.example
└── README.md
```

Go to your repo homepage — if you see all 9 files listed, you're done. ✅

---

## Part 7: Using It on Any Computer

**Step 1** — Download your repo:
Click the green **"Code"** button → **"Download ZIP"** → unzip on your Desktop

**Step 2** — Install dependencies (only once per computer):
```
pip install -r requirements.txt
npm install
```

**Step 3** — Put your files in the folder, then run:
```
python3 run_report.py --xlsx TataMotors.xlsx
