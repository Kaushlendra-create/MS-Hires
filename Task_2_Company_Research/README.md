# Task 2 — Company research and contact-data workflow

This folder provides a transparent CSV import and validation workflow for company-level records.
It does not scrape LinkedIn, infer email addresses, or invent records. The supplied input file is
an empty schema: no verified, cross-industry company dataset was present in the project.

## Run

From the project root:

```powershell
python Task_2_Company_Research\src\import_companies.py
```

Add records to `data\company_records.csv` from sources you are allowed to use, then run the
command again. The workbook is written to `output\Company_Research.xlsx`. Rows missing a company
name, valid website, or source URL—or containing an invalid email/source pair—are written to
`output\rejected_company_rows.csv` with a reason. Duplicate website domains are reported there too.

Required source fields are retained with each record. Supply a separate email source URL whenever
an email address is included. Only use company-level contact details from sources whose terms
permit collection and use; keep consent/opt-in evidence before using addresses for outreach.
LinkedIn company URLs are optional source data, not a scraping target.

## Dataset size and scope

The importer supports large CSV inputs (including 50,000+ rows), but this project does not contain
50,000 verified company records. A large target is not a substitute for sourced, accurate data.
Build the dataset by importing appropriately licensed business registries, official company
directories, or records collected with permission; record the source and verification date for
every row. The importer validates supplied rows but does not independently verify company claims.
