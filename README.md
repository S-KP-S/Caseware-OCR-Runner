# Caseware OCR Runner

Simple local UI to batch OCR PDFs, scans, and bank CSVs into a Caseware-ready CSV.

## What it does
- One CSV per client folder in `Extracted-Data\`
- One summary workbook with monthly totals + missing field flags
- Supports PDFs, images, and CSV bank exports
- No API keys are stored on disk

## Setup
1) Install Python 3.11+ (3.13 works).
2) Install dependencies:

```bash
py -m pip install -r requirements.txt
```

## Run
```bash
py app.py
```

Paste your OpenRouter API key, set the model, pick the input folder, then click **Run OCR**.

## Output
For a folder input:
- `Extracted-Data\{FolderName}_transactions.csv`
- `Extracted-Data\{FolderName}_summary.xlsx`

## Notes
- Free models are rate-limited. Use a lower RPM for stability.
- Use "Refine amounts" for higher accuracy (slower).
