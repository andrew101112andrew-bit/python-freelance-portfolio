import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from cleaner import process, write_csv, FIELDS  # noqa: E402


class CleanerTests(unittest.TestCase):
    def test_merge_normalizes_and_reports_rejected_rows(self):
        root = Path(__file__).resolve().parents[1]
        cleaned, rejected, summary = process([
            root / "examples/contacts_a.csv", root / "examples/contacts_b.csv"
        ])
        self.assertEqual(summary, {"input_files": 2, "rows_read": 6, "rows_saved": 3, "rows_rejected": 3})
        self.assertEqual([row["email"] for row in cleaned],
                         ["anna@example.com", "ivan@example.com", "olga@example.com"])
        self.assertEqual(cleaned[0]["name"], "Анна Сергеева")
        self.assertEqual({row["reason"] for row in rejected},
                         {"дубликат email", "некорректный или пустой email"})

    def test_missing_email_column_is_clear_error(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "bad.csv"
            path.write_text("name,phone\nАнна,123\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "email/почта"):
                process([path])

    def test_output_is_readable_as_utf8_with_bom(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "cleaned.csv"
            write_csv(path, FIELDS, [{"name": "Анна", "email": "anna@example.com",
                                      "phone": "", "company": "", "source": "a.csv"}])
            with path.open(encoding="utf-8-sig", newline="") as stream:
                self.assertEqual(next(csv.DictReader(stream))["name"], "Анна")


if __name__ == "__main__":
    unittest.main()
