"""Build a manager-ready Microsoft partner research workbook."""

import json
from datetime import date
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


BASE_DIR = Path(__file__).resolve().parent.parent
PROFILE_FILE = BASE_DIR / "data" / "partner_response.json"
OUTPUT_FILE = BASE_DIR / "output" / "Manager_Microsoft_Partner_Research.xlsx"
DIRECTORY_URL = "https://microsoftpartners.microsoft.com/directory-copilot.json"
DIRECTORY_PAGE = (
    "https://microsoftpartners.microsoft.com/abs/"
    "Microsoft-365-Copilot-Partner-Directory/"
)
TABLE_STYLE = "TableStyleMedium2"


def _text(value):
    if not isinstance(value, str):
        return ""
    value = value.strip()
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _list_text(value):
    if not isinstance(value, list):
        return ""
    return ", ".join(_text(item) for item in value if isinstance(item, str))


def fetch_directory():
    request = Request(
        DIRECTORY_URL,
        headers={"User-Agent": "MicrosoftPartnerManagerReport/1.0"},
    )
    try:
        with urlopen(request, timeout=30) as response:
            data = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"Could not read the public partner directory: {error}") from error

    partners = data.get("partnersList") if isinstance(data, dict) else None
    if not isinstance(partners, list) or any(not isinstance(item, dict) for item in partners):
        raise RuntimeError("The partner directory response has an invalid partnersList.")
    return partners


def read_supplementary_profiles():
    if not PROFILE_FILE.is_file():
        return [], None
    try:
        with PROFILE_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"Could not read supplementary profiles: {error}") from error

    matching = data.get("matchingPartners") if isinstance(data, dict) else None
    profiles = matching.get("items") if isinstance(matching, dict) else None
    if not isinstance(profiles, list) or any(not isinstance(item, dict) for item in profiles):
        raise RuntimeError("The supplementary file has an invalid matchingPartners.items list.")
    return profiles, data


def _format_sheet(sheet, widths, wrap_columns=()):
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 34
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1565C0")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=cell.column in wrap_columns,
            )
            if isinstance(cell.value, str) and cell.value.startswith(("https://", "http://")):
                cell.hyperlink = cell.value


def _add_table(sheet, name, last_column, last_row=None):
    ref = f"A1:{last_column}{last_row or sheet.max_row}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name=TABLE_STYLE,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def create_workbook(partners, profiles, profile_response):
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Manager Summary"
    summary.append(["Microsoft Partner Research", ""])
    summary.append(["Generated", date.today().isoformat()])
    summary.append(["Directory scope", "Microsoft 365 Copilot Partner Directory"])
    summary.append(["Partner companies listed", len(partners)])
    summary.append(["Supplementary company profiles", len(profiles)])
    summary.append(
        [
            "Task 1 - company links",
            "Completed from Microsoft's public directory; see Partner Directory and Marketplace Offers.",
        ]
    )
    summary.append(
        [
            "Task 1 - senior leaders",
            "The supplied directory feeds do not provide verified senior-leader names or titles. "
            "Use the Leadership Tracker to research and record official public sources.",
        ]
    )
    summary.append(
        [
            "Task 2 - 50,000 LinkedIn email addresses",
            "Not collected. LinkedIn/email harvesting is not included; Outreach Contacts is a "
            "blank template for contacts with documented permission.",
        ]
    )
    summary.append(["Primary source page", DIRECTORY_PAGE])
    summary.append(["Primary data feed", DIRECTORY_URL])
    summary.append(["Supplementary source file", str(PROFILE_FILE.relative_to(BASE_DIR)) if profiles else "Not available"])
    summary.append(
        [
            "Important",
            "Partner listing indicates Copilot directory presence; verify partner status and any "
            "leadership details with the linked company/Microsoft sources before external use.",
        ]
    )
    summary.column_dimensions["A"].width = 38
    summary.column_dimensions["B"].width = 110
    for cell in summary[1]:
        cell.font = Font(bold=True, size=16, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1565C0")
    for row in summary.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = f"A1:B{summary.max_row}"
    for row_num in range(2, summary.max_row + 1):
        summary.row_dimensions[row_num].height = 34
    summary["B9"].hyperlink = DIRECTORY_PAGE
    summary["B10"].hyperlink = DIRECTORY_URL

    directory = workbook.create_sheet("Partner Directory")
    directory.append(
        [
            "Partner company",
            "Company website",
            "Microsoft marketplace offer categories",
            "Offers listed",
            "Countries listed",
            "Date added",
        ]
    )
    for partner in partners:
        offers = partner.get("partnerOffers")
        offers = offers if isinstance(offers, list) else []
        categories = sorted(
            {
                _text(offer.get("offerCategory"))
                for offer in offers
                if isinstance(offer, dict) and _text(offer.get("offerCategory"))
            }
        )
        countries = sorted(
            {
                _text(country)
                for offer in offers
                if isinstance(offer, dict) and isinstance(offer.get("offerCountries"), list)
                for country in offer["offerCountries"]
                if isinstance(country, str) and _text(country)
            }
        )
        directory.append(
            [
                _text(partner.get("partnerName")),
                _text(partner.get("partnerLink")),
                ", ".join(categories),
                len(offers),
                ", ".join(countries),
                _text(partner.get("dateAdded")),
            ]
        )
    _format_sheet(directory, (34, 48, 48, 14, 76, 16), wrap_columns=(3, 5))
    _add_table(directory, "MicrosoftCopilotPartners", "F")

    offers_sheet = workbook.create_sheet("Marketplace Offers")
    offers_sheet.append(
        ["Partner company", "Company website", "Offer category", "Marketplace listing", "Countries"]
    )
    for partner in partners:
        offers = partner.get("partnerOffers")
        if not isinstance(offers, list):
            continue
        for offer in offers:
            if not isinstance(offer, dict):
                continue
            countries = offer.get("offerCountries")
            countries = countries if isinstance(countries, list) else []
            offers_sheet.append(
                [
                    _text(partner.get("partnerName")),
                    _text(partner.get("partnerLink")),
                    _text(offer.get("offerCategory")),
                    _text(offer.get("offerLink")),
                    ", ".join(_text(country) for country in countries if isinstance(country, str)),
                ]
            )
    _format_sheet(offers_sheet, (34, 48, 42, 84, 76), wrap_columns=(3, 4, 5))
    _add_table(offers_sheet, "PartnerMarketplaceOffers", "E")

    leadership = workbook.create_sheet("Leadership Tracker")
    leadership.append(
        [
            "Partner company",
            "Official company website",
            "Senior leader name",
            "Title / senior role",
            "Public professional profile URL",
            "Official source URL",
            "Verification status",
            "Verified date",
            "Notes",
        ]
    )
    for partner in partners:
        website = _text(partner.get("partnerLink"))
        leadership.append(
            [
                _text(partner.get("partnerName")),
                website,
                "",
                "",
                "",
                website,
                "Needs verification from official public source",
                "",
                "Directory does not supply a verified senior-leader name or title.",
            ]
        )
    _format_sheet(leadership, (34, 48, 30, 30, 54, 54, 42, 18, 68), wrap_columns=(9,))
    _add_table(leadership, "LeadershipResearchTracker", "I")

    profile_sheet = workbook.create_sheet("Supplementary Profiles")
    profile_sheet.append(
        [
            "Company",
            "Partner ID",
            "Country",
            "City",
            "Industry",
            "Products",
            "Services",
            "LinkedIn organization profile",
            "Description",
            "Gold competencies",
            "Silver competencies",
        ]
    )
    for profile in profiles:
        location = profile.get("location")
        location = location if isinstance(location, dict) else {}
        address = location.get("address")
        address = address if isinstance(address, dict) else {}
        profile_sheet.append(
            [
                _text(profile.get("name")),
                _text(profile.get("partnerId")),
                _text(address.get("country")),
                _text(address.get("city")),
                _list_text(profile.get("industryFocus")),
                _list_text(profile.get("product")),
                _list_text(profile.get("serviceType")),
                _text(profile.get("linkedInOrganizationProfile")),
                _text(profile.get("description")),
                _list_text(profile.get("competenciesGold")),
                _list_text(profile.get("competenciesSilver")),
            ]
        )
    _format_sheet(
        profile_sheet,
        (36, 40, 24, 24, 40, 40, 40, 54, 78, 48, 48),
        wrap_columns=(5, 6, 7, 9, 10, 11),
    )
    if profiles:
        _add_table(profile_sheet, "SupplementaryPartnerProfiles", "K")
    profile_notes = workbook.create_sheet("Profile Source Notes")
    profile_notes.append(["Item", "Details"])
    profile_notes.append(["Source file", str(PROFILE_FILE.relative_to(BASE_DIR)) if profiles else "Not available"])
    profile_notes.append(["Records in supplied response", len(profiles)])
    profile_notes.append(
        [
            "Estimated total matching partners",
            profile_response.get("estimatedTotalMatchingPartners", "") if profile_response else "",
        ]
    )
    profile_notes.append(
        [
            "Scope",
            "These are the records present in the supplied response only; they are not represented as the full estimated result set.",
        ]
    )
    _format_sheet(profile_notes, (38, 110), wrap_columns=(2,))

    outreach = workbook.create_sheet("Outreach Contacts")
    outreach.append(
        [
            "Company",
            "Company / role-based email",
            "Source URL",
            "Permission / opt-in evidence",
            "Date collected",
            "Do not contact",
        ]
    )
    _format_sheet(outreach, (34, 40, 56, 50, 20, 20))
    _add_table(outreach, "PermissionBasedOutreachContacts", "F", 50001)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_FILE)


def main():
    partners = fetch_directory()
    profiles, profile_response = read_supplementary_profiles()
    create_workbook(partners, profiles, profile_response)
    print(
        f"Created manager workbook: {OUTPUT_FILE} "
        f"({len(partners)} directory partners, {len(profiles)} supplementary profiles)"
    )


if __name__ == "__main__":
    main()
