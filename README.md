# Cornerstones Payroll IIF Generator

A Streamlit web app that converts a Paycom payroll export (.xlsx) into four
QuickBooks IIF journal entry files.

## Files generated

| Output | QB Entry | Filename |
|--------|----------|----------|
| 1 | Payroll JE | `Output 1 Payroll JE DD Mon.iif` |
| 2 | Taxes JE | `Output 2 Taxes JE DD Mon.iif` |
| 3 | 401K JE | `Output 3 Payroll 401k JE DD Mon.iif` |
| 4 | Payroll Fee JE | `Output 4 Payroll Fee JE DD Mon.iif` |

## How to use

1. Upload the source Paycom xlsx (e.g. `FY26 - PD 20260313 Payroll.xlsx`)
2. Enter the pay date, pay period number, and 2-digit fiscal year
3. Optionally enter the prior pay date (used in the cash line memo)
4. Click **Generate IIF Files**
5. Download each file and import into QuickBooks

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select this repo → set main file to `app.py`
4. Click **Deploy**

The reference class tables (`reference/`) are bundled in the repo and loaded
automatically — no need to upload them each run.

## Class override

If a labor code is missing or mapped incorrectly, use the **Advanced** panel
in the app to enter overrides in the format:

```
CODE=Full QB Class Path
```

Example:
```
22000=20000 Resource Development:22000 Volunteers
```
