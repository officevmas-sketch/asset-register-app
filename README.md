
# Asset Register Automation Portal

A simple online-ready web app for automating the asset register and computing depreciation for FY 2026–27.

## Features
- Upload your asset register workbook
- Automatically detects the latest FY sheet
- Rolls balances into FY 2026–27
- Calculates depreciation asset-wise
- Handles part-year depreciation for disposals
- Shows summary by asset class
- Exports an output workbook

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy online
This can be deployed easily on:
- Streamlit Community Cloud
- Render
- Azure App Service
- AWS / any Linux VM

## Notes
- The app is designed around the structure of your uploaded asset register.
- It reads the header from row 3 in the FY sheets.
- For WDV assets, it uses the rate if available; otherwise it infers a rate from cost, salvage value, and useful life.
- For SLM assets, depreciation = (cost - salvage) / life.


## Streamlit Cloud note
Add `runtime.txt` with `python-3.12` to use a stable Python version on Streamlit Cloud.
