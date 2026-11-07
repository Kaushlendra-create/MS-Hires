"""Validate sourced company CSV rows and export a research workbook."""

import argparse
import csv
import re
from pathlib import Path
from urllib.parse import urlsplit

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "data" / "company_records.csv"
OUTPUT_FILE = BASE_DIR / "output" / "Company_Research.xlsx"
REJECTED_FILE = BASE_DIR / "output" / "rejected_company_rows.csv"
FIELDS = (
    "Company name",
    "Website",
    "Industry",
    "Country",
    "Business email",
    "Contact page",
    "LinkedIn company URL",
    "Email source URL",
    "Verification status",
    "Source URL",
    "Last verified",
)
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _safe_text(value):
    value = (value or "").strip()
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _valid_url(value):
    try:
        parsed = urlsplit(value.strip())
        return (
            parsed.scheme in {"http", "https"}
            and bool(parsed.hostname)
            and (parsed.port is None or 0 < parsed.port < 65536)
        )
    except ValueError:
        return False


def _website_key(value):
    host = urlsplit(value.strip()).hostname or ""
    host = host.lower()
    return host[4:] if host.startswith("www.") else host


def load_records(input_file):
    accepted = []
    rejected = []
    seen_websites = set()

    with input_file.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        if reader.fieldnames is None:
            raise ValueError(f"CSV has no header row: {input_file}")
        missing = set(FIELDS) - set(reader.fieldnames)
        if missing:
            raise ValueError(f"CSV is missing required columns: {', '.join(sorted(missing))}")

        for row_number, raw in enumerate(reader, start=2):
            row = {field: _safe_text(raw.get(field)) for field in FIELDS}
            reasons = []
            if not row["Company name"]:
                reasons.append("Company name is required")
            if not _valid_url(row["Website"]):
                reasons.append("Website must be an http(s) URL")
            if not _valid_url(row["Source URL"]):
                reasons.append("Source URL must be an http(s) URL")
            if row["Contact page"] and not _valid_url(row["Contact page"]):
                reasons.append("Contact page must be an http(s) URL")
            if row["LinkedIn company URL"] and not _valid_url(row["LinkedIn company URL"]):
                reasons.append("LinkedIn company URL must be an http(s) URL")
            email = row["Business email"]
            if email and not EMAIL_PATTERN.fullmatch(email):
                reasons.append("Business email is not a valid email address")
            if email and not _valid_url(row["Email source URL"]):
                reasons.append("An email source URL is required for every supplied email")

            website_key = _website_key(row["Website"]) if _valid_url(row["Website"]) else ""
            if website_key and website_key in seen_websites:
                reasons.append("Duplicate website domain")
            if reasons:
                rejected.append(
                    {
                        "CSV row": row_number,
                        "Company name": row["Company name"],
                        "Reason": "; ".join(reasons),
                    }
                )
                continue
            seen_websites.add(website_key)
            accepted.append(row)

    return accepted, rejected


def write_workbook(records, rejected, output_file):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Company Database"
    sheet.append(list(FIELDS))

    for record in records:
        sheet.append([record[field] for field in FIELDS])
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1565C0")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if (
                cell.column in (2, 6, 7, 8, 10)
                and isinstance(cell.value, str)
                and _valid_url(cell.value)
            ):
                cell.hyperlink = cell.value
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in enumerate((32, 42, 28, 22, 38, 42, 48, 48, 28, 48, 18), start=1):
        sheet.column_dimensions[sheet.cell(row=1, column=column).column_letter].width = width
    if records:
        table = Table(displayName="SourcedCompanyRecords", ref=f"A1:K{sheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)

    notes = workbook.create_sheet("Read me")
    notes.append(["Measure", "Count / details"])
    notes.append(["Accepted company records", len(records)])
    notes.append(["Rejected rows", len(rejected)])
    notes.append(
        [
            "Email handling",
            "Emails are imported only when supplied with a source URL; none are inferred.",
        ]
    )
    notes.append(
        [
            "Coverage",
            "This workbook represents supplied CSV rows only; it is not a 50,000-record dataset "
            "unless the input contains that many valid unique companies.",
        ]
    )
    notes.column_dimensions["A"].width = 32
    notes.column_dimensions["B"].width = 110
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)


def write_rejections(rejected, rejected_file):
    rejected_file.parent.mkdir(parents=True, exist_ok=True)
    with rejected_file.open("w", encoding="utf-8-sig", newline="") as destination:
        writer = csv.DictWriter(destination, fieldnames=("CSV row", "Company name", "Reason"))
        writer.writeheader()
        writer.writerows(rejected)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=INPUT_FILE, help="Sourced company CSV")
    parser.add_argument("--output", type=Path, default=OUTPUT_FILE, help="Output workbook")
    parser.add_argument("--rejected", type=Path, default=REJECTED_FILE, help="Rejected-row report")
    args = parser.parse_args()

    records, rejected = load_records(args.input)
    write_workbook(records, rejected, args.output)
    write_rejections(rejected, args.rejected)
    print(
        f"Exported {len(records)} unique companies; rejected {len(rejected)} rows. "
        f"Workbook: {args.output}; rejected rows: {args.rejected}"
    )


if __name__ == "__main__":
    main()
