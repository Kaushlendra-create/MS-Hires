"""Create a workbook for managing permission-based company email contacts."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_FILE = Path(__file__).resolve().parent / "Permission_Based_Company_Contacts.xlsx"
HEADERS = (
    "Company email",
    "Company name",
    "Source URL",
    "Permission / opt-in evidence",
    "Date collected",
    "Do not contact",
)


def main():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Company contacts"
    sheet.append(HEADERS)
    sheet.append([""] * len(HEADERS))

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0078D4")

    table = Table(displayName="PermissionBasedContacts", ref="A1:F10001")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.freeze_panes = "A2"

    for column, width in {
        "A": 36,
        "B": 32,
        "C": 52,
        "D": 48,
        "E": 20,
        "F": 20,
    }.items():
        sheet.column_dimensions[column].width = width

    notes = workbook.create_sheet("Read me")
    notes.append(["Use this workbook for company contacts collected with permission."])
    notes.append(["Only import addresses you are authorized to use for outreach."])
    notes.append(["Keep the source and permission / opt-in evidence for each address."])
    notes.append(["Rows are prepared for up to 10,000 records; the workbook contains no harvested addresses."])

    workbook.save(OUTPUT_FILE)
    print(f"Created {OUTPUT_FILE} with capacity for 10,000 permission-based contacts.")


if __name__ == "__main__":
    main()
