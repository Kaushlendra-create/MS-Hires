import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from import_companies import FIELDS, load_records


class LoadRecordsTests(unittest.TestCase):
    def _write_csv(self, directory, rows):
        source = Path(directory) / "companies.csv"
        with source.open("w", encoding="utf-8", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(rows)
        return source

    def _row(self, **updates):
        row = {field: "" for field in FIELDS}
        row.update(
            {
                "Company name": "Example Company",
                "Website": "https://example.com",
                "Source URL": "https://directory.example.com/company",
            }
        )
        row.update(updates)
        return row

    def test_accepts_sourced_company_and_rejects_duplicate_domain(self):
        with tempfile.TemporaryDirectory() as directory:
            source = self._write_csv(
                directory,
                [
                    self._row(**{"Business email": "info@example.com", "Email source URL": "https://example.com/contact"}),
                    self._row(**{"Company name": "Duplicate", "Website": "https://www.example.com/about"}),
                ],
            )

            accepted, rejected = load_records(source)

        self.assertEqual(len(accepted), 1)
        self.assertEqual(accepted[0]["Business email"], "info@example.com")
        self.assertEqual(len(rejected), 1)
        self.assertIn("Duplicate website domain", rejected[0]["Reason"])

    def test_rejects_email_without_valid_source(self):
        with tempfile.TemporaryDirectory() as directory:
            source = self._write_csv(
                directory,
                [self._row(**{"Business email": "not-an-email"})],
            )

            accepted, rejected = load_records(source)

        self.assertEqual(accepted, [])
        self.assertIn("Business email is not a valid email address", rejected[0]["Reason"])
        self.assertIn("email source URL is required", rejected[0]["Reason"])

    def test_rejects_csv_missing_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "companies.csv"
            source.write_text("Company name,Website\nExample,https://example.com\n", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "missing required columns"):
                load_records(source)


if __name__ == "__main__":
    unittest.main()
