# Caseware OCR Runner

Caseware OCR Runner is a local OCR + review tool for accounting workflows.
It processes PDFs/images/CSVs, normalizes vendor names, assigns accounts, scores confidence, flags duplicates, and exports formats compatible with Caseware, QBO, and Xero.

## Features

- Profile-based settings persisted in `~/.caseware_ocr/config.json`
- Vendor normalization map in `~/.caseware_ocr/vendor_map.json`
- Chart of accounts map in `~/.caseware_ocr/accounts_map.json`
- File-hash cache in `Extracted-Data/.ocr_cache/`
- Smart 4-page PDF batching for OpenRouter calls
- Confidence scoring (0-100) and row-level flags
- Duplicate detection per file and across merged output
- Review tab with filtering, sorting, inline edits, and row deletion
- Styled summary workbook with:
  - `Monthly Totals`
  - `Missing Fields`
  - `Potential Duplicates`
  - `Low Confidence`

## Requirements

Install dependencies:

```bash
py -m pip install -r requirements.txt
```

`requirements.txt` is unchanged:

- `pymupdf`
- `requests`
- `openpyxl`
- `pillow`

## Run UI

```bash
py app.py
```

The UI has three tabs:

- `OCR`: run extraction with profiles and live progress
- `Review`: edit/filter/export extracted rows
- `Mappings`: manage vendor and account mappings

## CLI Usage

```bash
py ocr_tool.py --input "C:\path\to\folder" --model "nvidia/nemotron-nano-12b-v2-vl:free"
```

Set API key in your shell:

```powershell
$env:OPENROUTER_API_KEY="YOUR_KEY"
```

### New/important flags

- `--no-cache` disables `.ocr_cache` usage
- `--export-format {caseware,full,qbo,xero}`
- Existing flags still work:
  - `--max-pages`
  - `--rpm`
  - `--currency`
  - `--no-recursive`
  - `--no-refine-amounts`
  - `--max-retries`
  - `--retry-backoff`
  - `--retry-max-sleep`
  - `--output`

## Output

Directory input writes to:

- `Extracted-Data/{FolderName}_transactions.csv`
- `Extracted-Data/{FolderName}_summary.xlsx`

Single-file input writes to:

- `Extracted-Data/{FileName}_transactions.csv`
- `Extracted-Data/{FileName}_summary.xlsx`

