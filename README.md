# Asset Register Automation App

This is a Streamlit web app built for the uploaded TeamLease asset register structure.

## Features
- Upload the existing asset register workbook
- Detect and read the correct header row
- Select the source sheet
- Roll balances forward into FY 2026-27
- Compute month-wise depreciation for Apr-26 to Mar-27
- Capture optional disposals for FY 2026-27
- Download the generated FY 2026-27 register and summary workbook

## Files
- `app.py` - Streamlit application
- `requirements.txt` - Python dependencies
- `runtime.txt` - Python version pin for Streamlit Cloud

## Local run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud
- Repository: your GitHub repo
- Branch: main
- Main file path: `app.py`

The app shows its version on the page so you can confirm the latest deployment is live.
