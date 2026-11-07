# Task 1 — Microsoft partner research

This folder contains research sourced from Microsoft's public Microsoft 365 Copilot Partner Directory.

## Run

From the project root:

```powershell
python Task_1_Microsoft_Partner_Research\src\create_manager_workbook.py
python Task_1_Microsoft_Partner_Research\src\extract_partners.py
```

The first command fetches the current public directory feed and writes
`output\Manager_Microsoft_Partner_Research.xlsx`. The second exports the
supplementary records supplied in `data\partner_response.json` to
`output\Microsoft_Partner_Profiles.xlsx`.

## Contents

- `data\partner_response.json` — supplied Microsoft partner response (18 records).
- `output\Microsoft_365_Copilot_Partners.xlsx` — existing company and marketplace links.
- `output\Microsoft_365_Copilot_Partners_Company_Links.xlsx` — existing website, career-page,
  and company LinkedIn links.
- `output\Manager_Microsoft_Partner_Research.xlsx` — directory, offers, and leadership research
  tracker.

The public directory provides company/offer records, not verified executive names, individual
professional profiles, or business email addresses. The leadership tracker is a research worksheet
and must be completed from cited, public company sources. A directory listing is evidence of
inclusion in this directory, not a blanket claim about every Microsoft partner designation.
