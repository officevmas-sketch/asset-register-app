# Asset Register Automation App - FY 2026-27

## What it does
This Streamlit app automates the asset register for FY 2026-27.

It:
- reads the FY 2025-26 closing workbook as the opening base
- computes month-wise depreciation from April 2026 to March 2027
- supports additions during the year
- supports disposals during the year
- prevents net block from going below scrap/salvage value for selected assets
- exports a new workbook with opening data, movement inputs, detailed register, and summary

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Input files
1. Base register: the FY 2025-26 workbook
2. Additions file: optional xlsx using the app's template
3. Disposals file: optional xlsx using the app's template

## Important note
The app assumes full disposal at asset-ID level unless a disposed amount is explicitly provided.


## Template notes
- Additions template uses the corrected column name **Assets Description**.
- Depreciation Rate remains in the template as a manual-entry field (no dropdown).
- Disposal template can auto-fill disposed amount from selected Asset ID gross block.
