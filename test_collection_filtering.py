import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from collection_filtering import build_nsano_collection_filter_workbook


class NsanoCollectionFilteringTests(unittest.TestCase):
    def test_extra_csv_value_does_not_trigger_none_casefold_error(self):
        with TemporaryDirectory() as tmpdir:
            source = Path(tmpdir) / "nsano.csv"
            output = Path(tmpdir) / "filtered.xlsx"
            source.write_text(
                "Result,Type,Amount\nSuccessful,W2A,10,unexpected\n",
                encoding="utf-8",
            )

            result = build_nsano_collection_filter_workbook(output, [source])

            self.assertTrue(output.exists())
            self.assertEqual(result.filter_counts["Successful W2A"], 1)

    def test_missing_required_header_has_clear_error(self):
        with TemporaryDirectory() as tmpdir:
            source = Path(tmpdir) / "nsano.csv"
            output = Path(tmpdir) / "filtered.xlsx"
            source.write_text(
                "Result,Amount\nSuccessful,10,unexpected\n",
                encoding="utf-8",
            )

            with self.assertRaisesRegex(ValueError, r"requires.*Type"):
                build_nsano_collection_filter_workbook(output, [source])


if __name__ == "__main__":
    unittest.main()
