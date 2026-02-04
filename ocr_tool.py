import argparse
import base64
import csv
import json
import os
import re
import sys
import time
from datetime import datetime

import fitz  # PyMuPDF
import requests
from openpyxl import Workbook

PROMPT = (
    "You are an OCR extraction assistant for accounting imports. Extract transactions from the provided "
    "document images. Return a single JSON object only with key: transactions. "
    "transactions must be a list of objects with keys: date, description, amount, direction, "
    "transaction_type, currency, page_number. "
    "date must be YYYY-MM-DD. amount should be the numeric value as shown (commas/decimals ok); do not make negative. "
    "direction must be 'Debit' or 'Credit'. "
    "transaction_type must be one of Deposit, Withdrawal, Fee, Interest, Refund, Transfer, Payment, Purchase, Tax, Other. "
    "currency must be ISO 4217 (e.g., CAD, USD). page_number is 1-based. "
    "For receipts/invoices without line items, create a single transaction using the total amount. "
    "For bank/credit card statements, return one transaction per line item. "
    "Do not include balances or subtotals. "
    "If you cannot find transactions, return an empty list. Do not include any other keys or wrap in markdown."
)

AMOUNT_PROMPT = (
    "You are correcting transaction amounts for accounting import. "
    "Use the provided document images to verify and fix ONLY the Amount values. "
    "You are given a JSON list of transactions with an index. "
    "Return a single JSON object with key: transactions, a list of objects with keys: index, amount. "
    "Do not change the number of items or their index. "
    "amount must be the numeric value as shown (commas/decimals ok) and must not be negative. "
    "If the amount is not visible, return an empty string for amount. "
    "Do not include any other keys or wrap in markdown."
)

CSV_FIELDNAMES = [
    "Date",
    "Description",
    "Amount",
    "Direction",
    "Transaction Type",
    "Currency",
    "Source-File Name",
    "Page Number",
]

SUMMARY_SHEET_TOTALS = "Monthly Totals"
SUMMARY_SHEET_MISSING = "Missing Fields"
OUTPUT_DIRNAME = "Extracted-Data"
DEFAULT_CURRENCY = "CAD"
KEY_FIELDS = [
    "Date",
    "Description",
    "Amount",
    "Direction",
    "Transaction Type",
    "Currency",
    "Source-File Name",
]

SKIP_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
CSV_EXTENSIONS = {".csv"}

TRANSACTION_TYPE_SET = {
    "Deposit",
    "Withdrawal",
    "Fee",
    "Interest",
    "Refund",
    "Transfer",
    "Payment",
    "Purchase",
    "Tax",
    "Other",
}

CSV_DATE_KEYS = {
    "date",
    "transactiondate",
    "postingdate",
    "postdate",
    "transdate",
    "valuedate",
}
CSV_DESC_KEYS = {
    "description",
    "details",
    "memo",
    "payee",
    "merchant",
    "name",
    "narration",
    "reference",
}
CSV_AMOUNT_KEYS = {
    "amount",
    "transactionamount",
    "amt",
    "value",
    "total",
    "amountcad",
    "amountcredit",
    "amountdebit",
}
CSV_DEBIT_KEYS = {
    "debit",
    "withdrawal",
    "withdrawals",
    "moneyout",
    "debitamount",
    "charges",
    "charge",
    "paidout",
}
CSV_CREDIT_KEYS = {
    "credit",
    "deposit",
    "deposits",
    "moneyin",
    "creditamount",
    "paymentreceived",
    "receipt",
}
CSV_DIRECTION_KEYS = {
    "direction",
    "drcr",
    "debitcredit",
    "debitcreditindicator",
}
CSV_CURRENCY_KEYS = {
    "currency",
    "curr",
    "ccy",
    "currencycode",
    "iso",
}
CSV_TYPE_KEYS = {
    "transactiontype",
    "type",
    "category",
    "classification",
    "transtype",
    "transaction_type",
}

RATE_LIMITER = None
DEFAULT_RPM = 12
DEFAULT_MAX_RETRIES = 5
DEFAULT_RETRY_BACKOFF = 5
DEFAULT_RETRY_MAX_SLEEP = 60


def parse_args():
    parser = argparse.ArgumentParser(
        description="OCR PDFs or images via OpenRouter and write a Caseware-ready CSV plus summary."
    )
    parser.add_argument("--input", required=True, help="Path to a PDF/image or a directory")
    parser.add_argument("--output", help="Path to output CSV (single or batch)")
    parser.add_argument("--model", default="nvidia/nemotron-nano-12b-v2-vl:free", help="OpenRouter model")
    parser.add_argument("--max-pages", type=int, default=12, help="Max pages to render")
    parser.add_argument("--zoom", type=float, default=2.0, help="Render zoom factor")
    parser.add_argument("--currency", default=DEFAULT_CURRENCY, help="Default currency code")
    parser.add_argument("--rpm", type=int, default=DEFAULT_RPM, help="Max OpenRouter requests per minute")
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES, help="Max OpenRouter retries")
    parser.add_argument("--retry-backoff", type=int, default=DEFAULT_RETRY_BACKOFF, help="Base retry backoff seconds")
    parser.add_argument("--retry-max-sleep", type=int, default=DEFAULT_RETRY_MAX_SLEEP, help="Max retry sleep seconds")
    parser.add_argument("--recursive", action="store_true", help="Process subfolders (default)")
    parser.add_argument("--no-recursive", dest="recursive", action="store_false", help="Disable subfolder processing")
    parser.add_argument(
        "--no-refine-amounts",
        action="store_false",
        dest="refine_amounts",
        help="Disable second-pass amount correction",
    )
    parser.set_defaults(refine_amounts=True, recursive=True)
    return parser.parse_args()


def render_pdf_to_png_bytes(pdf_path, max_pages, zoom):
    doc = fitz.open(pdf_path)
    images = []
    for i, page in enumerate(doc):
        if max_pages and i >= max_pages:
            break
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        images.append(pix.tobytes("png"))
    page_count = doc.page_count
    doc.close()
    return images, page_count


def strip_code_fences(text):
    text = text.strip()
    if not text.startswith("```"):
        return text
    lines = text.splitlines()
    if lines and lines[0].startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].startswith("```"):
        lines = lines[:-1]
    return "\n".join(lines).strip()


def parse_json(text):
    text = strip_code_fences(text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise
        return json.loads(text[start:end + 1])


def safe_str(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    return str(value)


def normalize_header(value):
    text = safe_str(value).lower().strip()
    text = re.sub(r"[^a-z0-9]", "", text)
    return text


class RateLimiter:
    def __init__(self, rpm):
        self.min_interval = 60.0 / rpm if rpm and rpm > 0 else 0.0
        self.last_time = 0.0

    def wait(self):
        if self.min_interval <= 0:
            return
        now = time.time()
        elapsed = now - self.last_time
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self.last_time = time.time()


def is_skipped_extension(path):
    ext = os.path.splitext(path)[1].lower()
    return ext in SKIP_EXTENSIONS


def is_csv_extension(path):
    ext = os.path.splitext(path)[1].lower()
    return ext in CSV_EXTENSIONS


def is_image_extension(path):
    ext = os.path.splitext(path)[1].lower()
    return ext in IMAGE_EXTENSIONS


def load_images(path, max_pages, zoom):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return render_pdf_to_png_bytes(path, max_pages, zoom)
    if ext in IMAGE_EXTENSIONS:
        with open(path, "rb") as f:
            return [f.read()], 1
    return None, 0


def build_content(images, prompt_text):
    content = [{"type": "text", "text": prompt_text}]
    for img in images:
        b64 = base64.b64encode(img).decode("ascii")
        content.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + b64}})
    return content


def compute_retry_sleep(response, attempt, backoff_base, backoff_max):
    reset_header = response.headers.get("X-RateLimit-Reset")
    if reset_header:
        try:
            reset_val = float(reset_header)
            if reset_val > 1e12:
                reset_val /= 1000.0
            delay = max(0.0, reset_val - time.time()) + 1.0
            return min(backoff_max, delay)
        except ValueError:
            pass
    delay = backoff_base * (2 ** (attempt - 1))
    return min(backoff_max, delay)


def call_openrouter(content, api_key, model, max_retries, backoff_base, backoff_max):
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {"role": "user", "content": content}
        ],
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "X-Title": "Caseware OCR Runner",
    }

    for attempt in range(1, max_retries + 1):
        if RATE_LIMITER:
            RATE_LIMITER.wait()
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=120,
        )

        if response.status_code == 200:
            try:
                data = response.json()
                content_text = data["choices"][0]["message"]["content"]
                return parse_json(content_text)
            except Exception as exc:
                if attempt < max_retries:
                    sleep_time = min(backoff_max, backoff_base * (2 ** (attempt - 1)))
                    print(f"INFO: OpenRouter parse error, retrying in {sleep_time:.1f}s")
                    time.sleep(sleep_time)
                    continue
                raise RuntimeError(f"OpenRouter response parse failed: {exc}") from exc

        text_lower = response.text.lower()
        retryable = response.status_code in (408, 409, 429, 500, 502, 503, 504)
        if "degraded" in text_lower:
            retryable = True

        if retryable and attempt < max_retries:
            sleep_time = compute_retry_sleep(response, attempt, backoff_base, backoff_max)
            print(f"INFO: OpenRouter retry {attempt}/{max_retries} in {sleep_time:.1f}s (status {response.status_code})")
            time.sleep(sleep_time)
            continue

        raise RuntimeError(f"OpenRouter request failed: {response.status_code} {response.text}")

    raise RuntimeError("OpenRouter request failed after retries")


def refine_amounts(rows, images, api_key, model, max_retries, backoff_base, backoff_max):
    payload_rows = []
    for idx, row in enumerate(rows):
        payload_rows.append(
            {
                "index": idx,
                "date": row.get("Date", ""),
                "description": row.get("Description", ""),
                "direction": row.get("Direction", ""),
                "transaction_type": row.get("Transaction Type", ""),
                "currency": row.get("Currency", ""),
                "page_number": row.get("Page Number", ""),
                "amount": row.get("Amount", ""),
            }
        )

    prompt_text = AMOUNT_PROMPT + "\nTransactions JSON:\n" + json.dumps(
        payload_rows, ensure_ascii=True
    )
    content = build_content(images, prompt_text)
    extracted = call_openrouter(content, api_key, model, max_retries, backoff_base, backoff_max)

    if isinstance(extracted, list):
        corrections = extracted
    elif isinstance(extracted, dict):
        corrections = extracted.get("transactions") or extracted.get("amounts") or []
    else:
        corrections = []

    updated = 0
    for item in corrections:
        if not isinstance(item, dict):
            continue
        idx = item.get("index")
        if idx is None:
            continue
        try:
            idx = int(idx)
        except (TypeError, ValueError):
            continue
        if idx < 0 or idx >= len(rows):
            continue
        amount = safe_str(item.get("amount"))
        if amount != "":
            rows[idx]["Amount"] = amount
            updated += 1

    return updated


def normalize_transactions(extracted):
    if isinstance(extracted, list):
        return extracted
    if not isinstance(extracted, dict):
        return []
    transactions = extracted.get("transactions")
    if isinstance(transactions, list):
        return transactions
    return []


def normalize_transaction_type(value):
    if not value:
        return ""
    cleaned = str(value).strip()
    lowered = cleaned.lower()
    alias_map = {
        "credit": "Deposit",
        "cr": "Deposit",
        "deposit": "Deposit",
        "debit": "Withdrawal",
        "dr": "Withdrawal",
        "withdrawal": "Withdrawal",
        "withdraw": "Withdrawal",
        "fee": "Fee",
        "charge": "Fee",
        "interest": "Interest",
        "refund": "Refund",
        "reversal": "Refund",
        "transfer": "Transfer",
        "payment": "Payment",
        "purchase": "Purchase",
        "pos": "Purchase",
        "tax": "Tax",
    }
    if lowered in alias_map:
        return alias_map[lowered]
    cleaned_title = cleaned.title()
    if cleaned_title in TRANSACTION_TYPE_SET:
        return cleaned_title
    return cleaned_title if cleaned_title else ""


def infer_direction(direction, transaction_type):
    if direction:
        cleaned = str(direction).strip().lower()
        if cleaned in ("debit", "dr"):
            return "Debit"
        if cleaned in ("credit", "cr"):
            return "Credit"
        if "debit" in cleaned:
            return "Debit"
        if "credit" in cleaned:
            return "Credit"
        if "withdraw" in cleaned:
            return "Debit"
        if "deposit" in cleaned:
            return "Credit"
    if transaction_type:
        ttype = str(transaction_type).strip().lower()
        if ttype in ("deposit", "interest", "refund"):
            return "Credit"
        if ttype in ("withdrawal", "fee", "payment", "purchase", "tax"):
            return "Debit"
    return ""


def infer_transaction_type_from_description(description):
    if not description:
        return ""
    text = str(description).lower()
    if "refund" in text or "reversal" in text:
        return "Refund"
    if "interest" in text:
        return "Interest"
    if "fee" in text or "charge" in text:
        return "Fee"
    if "tax" in text or "hst" in text or "cra" in text:
        return "Tax"
    if "transfer" in text or "e-transfer" in text or "etransfer" in text:
        return "Transfer"
    if "payroll" in text or "payment" in text:
        return "Payment"
    if "deposit" in text or "payment received" in text:
        return "Deposit"
    if "withdrawal" in text or "atm" in text:
        return "Withdrawal"
    if "purchase" in text or "pos" in text or "card" in text:
        return "Purchase"
    return ""


def build_rows(input_path, extracted, default_currency, page_count, source_name=None):
    source_name = source_name or os.path.basename(input_path)
    transactions = normalize_transactions(extracted)
    rows = []
    for item in transactions:
        if not isinstance(item, dict):
            continue
        date = safe_str(item.get("date") or item.get("transaction_date"))
        description = safe_str(
            item.get("description")
            or item.get("memo")
            or item.get("payee")
            or item.get("vendor")
            or item.get("details")
        )
        amount = safe_str(item.get("amount") or item.get("total") or item.get("value"))
        transaction_type = normalize_transaction_type(item.get("transaction_type") or item.get("type"))
        if not transaction_type:
            transaction_type = infer_transaction_type_from_description(description)
        direction = infer_direction(item.get("direction") or item.get("sign"), transaction_type)
        currency = safe_str(item.get("currency") or default_currency)
        page_number = safe_str(item.get("page_number") or item.get("page") or "")
        if not page_number and page_count == 1:
            page_number = "1"

        rows.append(
            {
                "Date": date,
                "Description": description,
                "Amount": amount,
                "Direction": direction,
                "Transaction Type": transaction_type,
                "Currency": currency,
                "Source-File Name": source_name,
                "Page Number": page_number,
            }
        )
    return rows


def write_csv(output_path, rows):
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def parse_date_info(value):
    if not value:
        return None, "missing"
    text = str(value).strip()
    match = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", text)
    if match:
        year, month, day = match.groups()
        try:
            return datetime(int(year), int(month), int(day)).date(), "ok"
        except ValueError:
            return None, "invalid"
    match = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", text)
    if match:
        month, day, year = match.groups()
        try:
            return datetime(int(year), int(month), int(day)).date(), "ok"
        except ValueError:
            return None, "invalid"
    return None, "unparseable"


def parse_amount(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    cleaned = re.sub(r"[^\d,.\-]", "", text)
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(",", "")
    elif "," in cleaned and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def get_first_value(row, keys):
    for key in keys:
        value = row.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text != "":
            return text
    return ""


def parse_csv_file(path, default_currency, source_name):
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel

        reader = csv.reader(f, dialect)
        raw_rows = list(reader)

    if not raw_rows:
        return rows

    header = raw_rows[0]
    norm_headers = [normalize_header(h) for h in header]
    data_rows = raw_rows[1:]

    for raw in data_rows:
        if not any(cell.strip() for cell in raw if isinstance(cell, str)):
            continue
        row_dict = {}
        for idx, key in enumerate(norm_headers):
            value = raw[idx] if idx < len(raw) else ""
            row_dict[key] = value.strip() if isinstance(value, str) else str(value)

        date = get_first_value(row_dict, CSV_DATE_KEYS)
        description = get_first_value(row_dict, CSV_DESC_KEYS)
        amount = get_first_value(row_dict, CSV_AMOUNT_KEYS)
        debit = get_first_value(row_dict, CSV_DEBIT_KEYS)
        credit = get_first_value(row_dict, CSV_CREDIT_KEYS)
        direction_raw = get_first_value(row_dict, CSV_DIRECTION_KEYS)
        type_raw = get_first_value(row_dict, CSV_TYPE_KEYS)
        currency = get_first_value(row_dict, CSV_CURRENCY_KEYS) or default_currency

        if not any([date, description, amount, debit, credit]):
            continue

        transaction_type = normalize_transaction_type(type_raw)
        direction = infer_direction(direction_raw, transaction_type)

        if not direction and type_raw:
            maybe_direction = infer_direction(type_raw, transaction_type)
            if maybe_direction:
                direction = maybe_direction
            elif not transaction_type:
                transaction_type = normalize_transaction_type(type_raw)

        if debit and not amount:
            amount = debit
            direction = direction or "Debit"
        elif credit and not amount:
            amount = credit
            direction = direction or "Credit"

        if amount:
            amount_value = parse_amount(amount)
            if amount_value is not None:
                if amount_value < 0:
                    if not direction:
                        direction = "Debit"
                elif amount_value > 0:
                    if not direction and not (debit or credit):
                        direction = "Credit"

        if not transaction_type:
            transaction_type = infer_transaction_type_from_description(description)

        rows.append(
            {
                "Date": date,
                "Description": description,
                "Amount": amount,
                "Direction": direction,
                "Transaction Type": transaction_type,
                "Currency": currency,
                "Source-File Name": source_name,
                "Page Number": "",
            }
        )

    return rows


def write_summary(output_path, rows):
    wb = Workbook()
    default_ws = wb.active
    totals_ws = default_ws
    totals_ws.title = SUMMARY_SHEET_TOTALS

    monthly = {}
    missing_rows = []

    for row in rows:
        missing = []
        for field in KEY_FIELDS:
            value = row.get(field)
            if value is None or str(value).strip() == "":
                missing.append(field)

        parsed_date, date_status = parse_date_info(row.get("Date"))
        if row.get("Date") and date_status != "ok":
            if date_status == "invalid":
                missing.append("Date (invalid)")
            else:
                missing.append("Date (unparseable)")

        amount_value = parse_amount(row.get("Amount"))
        if row.get("Amount") and amount_value is None:
            missing.append("Amount (unparseable)")

        if missing:
            missing_rows.append((row, ", ".join(sorted(set(missing)))))

        if not parsed_date:
            continue
        if amount_value is None:
            continue

        month_key = f"{parsed_date.year:04d}-{parsed_date.month:02d}"
        bucket = monthly.setdefault(
            month_key,
            {"count": 0, "debit": 0.0, "credit": 0.0, "net": 0.0},
        )
        bucket["count"] += 1
        direction = str(row.get("Direction", "")).strip().title()
        if direction == "Credit":
            bucket["credit"] += amount_value
            bucket["net"] += amount_value
        elif direction == "Debit":
            bucket["debit"] += amount_value
            bucket["net"] -= amount_value
        else:
            bucket["net"] += amount_value

    totals_ws.append(["Month", "Row Count", "Debit Total", "Credit Total", "Net Total"])
    for month_key in sorted(monthly.keys()):
        bucket = monthly[month_key]
        totals_ws.append(
            [
                month_key,
                bucket["count"],
                round(bucket["debit"], 2),
                round(bucket["credit"], 2),
                round(bucket["net"], 2),
            ]
        )

    missing_ws = wb.create_sheet(title=SUMMARY_SHEET_MISSING)
    missing_ws.append(CSV_FIELDNAMES + ["Missing Fields"])
    for row, missing in missing_rows:
        missing_ws.append([row.get(field, "") for field in CSV_FIELDNAMES] + [missing])

    wb.save(output_path)


def get_source_name(root_dir, path):
    if root_dir:
        try:
            return os.path.relpath(path, root_dir)
        except ValueError:
            return os.path.basename(path)
    return os.path.basename(path)


def iter_input_files(input_path, recursive):
    if recursive:
        for root, dirs, files in os.walk(input_path):
            dirs[:] = [
                d for d in dirs
                if d != OUTPUT_DIRNAME and not d.startswith(".")
            ]
            for name in files:
                yield os.path.join(root, name)
    else:
        for name in sorted(os.listdir(input_path)):
            yield os.path.join(input_path, name)


def main():
    args = parse_args()
    global RATE_LIMITER
    RATE_LIMITER = RateLimiter(args.rpm)
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        print("ERROR: OPENROUTER_API_KEY is not set.")
        print("Set it in PowerShell: $env:OPENROUTER_API_KEY=\"YOUR_KEY\"")
        sys.exit(1)

    input_path = args.input
    if os.path.isdir(input_path):
        output_dir = os.path.join(input_path, OUTPUT_DIRNAME)
        os.makedirs(output_dir, exist_ok=True)
        folder_name = os.path.basename(os.path.abspath(input_path))
        output_path = args.output or os.path.join(output_dir, f"{folder_name}_transactions.csv")
        summary_path = os.path.join(output_dir, f"{folder_name}_summary.xlsx")

        files = []
        for full_path in iter_input_files(input_path, args.recursive):
            if not os.path.isfile(full_path):
                continue
            if is_skipped_extension(full_path):
                continue
            if output_path and os.path.abspath(full_path) == os.path.abspath(output_path):
                continue
            if summary_path and os.path.abspath(full_path) == os.path.abspath(summary_path):
                continue
            if OUTPUT_DIRNAME in full_path.split(os.sep):
                continue
            files.append(full_path)

        if not files:
            print("ERROR: No eligible files found in the directory.")
            sys.exit(1)

        all_rows = []
        for path in files:
            source_name = get_source_name(input_path, path)
            if is_csv_extension(path):
                rows = parse_csv_file(path, args.currency, source_name)
                if not rows:
                    print(f"SKIP: No transactions found in CSV: {source_name}")
                    continue
                all_rows.extend(rows)
                continue

            if not (path.lower().endswith(".pdf") or is_image_extension(path)):
                print(f"SKIP: Unsupported file type: {source_name}")
                continue
            images, page_count = load_images(path, args.max_pages, args.zoom)
            if not images:
                print(f"SKIP: No pages rendered: {source_name}")
                continue
            content = build_content(images, PROMPT)
            try:
                extracted = call_openrouter(
                    content,
                    api_key,
                    args.model,
                    args.max_retries,
                    args.retry_backoff,
                    args.retry_max_sleep,
                )
            except Exception as exc:
                print(f"ERROR: {source_name} -> {exc}")
                continue
            rows = build_rows(path, extracted, args.currency, page_count, source_name=source_name)
            if not rows:
                print(f"SKIP: No transactions found: {source_name}")
                continue
            if args.refine_amounts:
                try:
                    updated = refine_amounts(
                        rows,
                        images,
                        api_key,
                        args.model,
                        args.max_retries,
                        args.retry_backoff,
                        args.retry_max_sleep,
                    )
                    if updated:
                        print(f"INFO: {source_name} amounts refined: {updated}")
                except Exception as exc:
                    print(f"ERROR: {source_name} amount refine failed -> {exc}")
            all_rows.extend(rows)

        if not all_rows:
            print("ERROR: No transactions were extracted.")
            sys.exit(1)

        write_csv(output_path, all_rows)
        write_summary(summary_path, all_rows)
        print(f"Wrote {output_path}")
        print(f"Wrote {summary_path}")
        return

    output_path = args.output
    if not output_path:
        base_dir = os.path.dirname(os.path.abspath(input_path))
        output_dir = os.path.join(base_dir, OUTPUT_DIRNAME)
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}_transactions.csv")
        summary_path = os.path.join(output_dir, f"{base_name}_summary.xlsx")
    else:
        summary_path = os.path.splitext(output_path)[0] + "_summary.xlsx"

    if is_csv_extension(input_path):
        source_name = os.path.basename(input_path)
        rows = parse_csv_file(input_path, args.currency, source_name)
        if not rows:
            print("ERROR: No transactions found in CSV.")
            sys.exit(1)
        write_csv(output_path, rows)
        write_summary(summary_path, rows)
        print(f"Wrote {output_path}")
        print(f"Wrote {summary_path}")
        return

    images, page_count = load_images(input_path, args.max_pages, args.zoom)
    if not images:
        print("ERROR: No pages rendered from input.")
        sys.exit(1)

    content = build_content(images, PROMPT)
    try:
        extracted = call_openrouter(
            content,
            api_key,
            args.model,
            args.max_retries,
            args.retry_backoff,
            args.retry_max_sleep,
        )
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    rows = build_rows(input_path, extracted, args.currency, page_count)
    if not rows:
        print("ERROR: No transactions were extracted.")
        sys.exit(1)
    if args.refine_amounts:
        try:
            updated = refine_amounts(
                rows,
                images,
                api_key,
                args.model,
                args.max_retries,
                args.retry_backoff,
                args.retry_max_sleep,
            )
            if updated:
                print(f"INFO: {os.path.basename(input_path)} amounts refined: {updated}")
        except Exception as exc:
            print(f"ERROR: amount refine failed -> {exc}")
    write_csv(output_path, rows)
    write_summary(summary_path, rows)
    print(f"Wrote {output_path}")
    print(f"Wrote {summary_path}")


if __name__ == "__main__":
    main()
