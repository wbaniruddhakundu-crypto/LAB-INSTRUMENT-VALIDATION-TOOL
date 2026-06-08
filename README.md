# Precision Validation Tool — Streamlit

Upload Siemens/Atellica analyzer Assay Report PDFs (SID `CH PRECISION DAYn Ln`)
and auto-generate one precision Excel per analyte — either from a built-in
template or populated straight into your own uploaded template, with **live
Excel formulas** (STDEV / AVERAGE / CV).

- 5-replicate PDFs → inter-day (one row per day)
- 10/15/20-replicate PDFs → intra-day (single within-run row)
- Optional **template upload**: data fills into your exact file, respecting
  however many days / replicates / control levels it has (2 or 3 levels both
  work; extra measurements are clipped to fit).

## Run locally

```bash
# 1. (optional) create a virtual environment
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 2. install dependencies
pip install -r requirements.txt

# 3. start the app
streamlit run app.py
```

Then open the URL Streamlit prints (default http://localhost:8501).

## Deploy to Streamlit Community Cloud

1. Push this folder to a GitHub repo.
2. On https://share.streamlit.io, create a new app pointing at `app.py`.
3. Streamlit Cloud installs `requirements.txt` automatically.

## Files

| File              | Purpose                                                        |
| ----------------- | ------------------------------------------------------------- |
| `app.py`          | Streamlit UI (uploads, detection, downloads)                  |
| `pdf_parser.py`   | Extracts analyzer text and parses measurements per analyte    |
| `excel_export.py` | Builds / populates the precision Excel, plus the ZIP of all   |
| `requirements.txt`| Python dependencies                                           |

## How it works

1. **PDF parsing** (`pdf_parser.py`): reconstructs text lines by each word's
   vertical position, finds each analyte heading, then counts sequential
   `CH PRECISION DAYn` measurement blocks to assign replicate numbers
   (the reportable Result value — the number before the unit — is used).
2. **Excel generation** (`excel_export.py`): either builds a sheet that matches
   the lab template cell-for-cell, or loads your uploaded template and writes
   values into its exact grid while preserving its formulas and formatting.
