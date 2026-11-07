"""Export company-level details from the saved Microsoft partner response."""

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent.parent
JSON_FILE = BASE_DIR / "data" / "partner_response.json"
EXCEL_FILE = BASE_DIR / "output" / "Microsoft_Partner_Profiles.xlsx"
HEADERS = (
    "Company name",
    "Partner ID",
    "Country",
    "City",
    "Description",
    "Industry",
    "Products",
    "Services",
    "LinkedIn organization profile",
    "Gold competencies",
    "Silver competencies",
    "Partner logo",
)


def _excel_text(value):
    """Keep imported strings as text if they begin with an Excel formula marker."""
    if not isinstance(value, str):
        return ""
    if value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _list_text(value):
    if not isinstance(value, list):
        return ""
    return ", ".join(str(item) for item in value if item is not None)


def _records_from_response(data):
    matching = data.get("matchingPartners") if isinstance(data, dict) else None
    partners = matching.get("items") if isinstance(matching, dict) else None
    if not isinstance(partners, list):
        raise ValueError("The response must contain matchingPartners.items.")

    records = []
    for partner in partners:
        if not isinstance(partner, dict):
            raise ValueError("The response contains a partner record that is not an object.")
        location = partner.get("location")
        location = location if isinstance(location, dict) else {}
        address = location.get("address")
        address = address if isinstance(address, dict) else {}
        records.append(
            (
                _excel_text(partner.get("name")),
                _excel_text(partner.get("partnerId")),
                _excel_text(address.get("country")),
                _excel_text(address.get("city")),
                _excel_text(partner.get("description")),
                _excel_text(_list_text(partner.get("industryFocus"))),
                _excel_text(_list_text(partner.get("product"))),
                _excel_text(_list_text(partner.get("serviceType"))),
                _excel_text(partner.get("linkedInOrganizationProfile")),
                _excel_text(_list_text(partner.get("competenciesGold"))),
                _excel_text(_list_text(partner.get("competenciesSilver"))),
                _excel_text(partner.get("logo")),
            )
        )
    return records


def write_workbook(records, response):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Partner companies"
    sheet.append(HEADERS)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0078D4")
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for record in records:
        sheet.append(record)
        row = sheet.max_row
        for column in range(1, len(HEADERS) + 1):
            sheet.cell(row=row, column=column).alignment = Alignment(
                vertical="top",
                wrap_text=column in (5, 6, 7, 8, 10, 11),
            )
        for column in (9, 12):
            cell = sheet.cell(row=row, column=column)
            if isinstance(cell.value, str) and cell.value.startswith(("https://", "http://")):
                cell.hyperlink = cell.value

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 32
    widths = (34, 40, 22, 24, 70, 38, 38, 38, 52, 48, 48, 52)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    notes = workbook.create_sheet("Read me")
    notes.append(["Source file", str(JSON_FILE.relative_to(BASE_DIR))])
    notes.append(["Matching records in this response", len(records)])
    notes.append(["Estimated total matching partners", response.get("estimatedTotalMatchingPartners", "")])
    notes.append(["Email addresses", "Not present in the supplied response; none were generated or inferred."])
    notes.append(["LinkedIn data", "Organization profile links only; no individual profiles were collected."])
    notes.column_dimensions["A"].width = 38
    notes.column_dimensions["B"].width = 90
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    workbook.save(EXCEL_FILE)


def main():
    if not JSON_FILE.is_file():
        raise FileNotFoundError(f"Partner response file not found: {JSON_FILE}")

    with JSON_FILE.open("r", encoding="utf-8") as file:
        response = json.load(file)
    records = _records_from_response(response)
    write_workbook(records, response)
    print(f"Exported {len(records)} company profiles to {EXCEL_FILE}")


if __name__ == "__main__":
    main()
