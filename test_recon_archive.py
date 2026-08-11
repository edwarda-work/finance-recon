import unittest

from recon_archive import _record_from_drive_file


SPECS = [
    ("Filtering", "Mambu Collection Filtering", "Mambu_Collection_Filtered.xlsx"),
]


class ArchiveIndexRecoveryTests(unittest.TestCase):
    def test_recovers_archive_metadata_from_drive_file(self):
        record = _record_from_drive_file(
            {
                "id": "drive-file-123",
                "name": (
                    "2026-06_Mambu Collection Filtering_deadbeef_"
                    "Mambu_Collection_Filtered.xlsx"
                ),
                "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "size": "4096",
                "createdTime": "2026-07-01T10:15:00.000Z",
                "webViewLink": "https://drive.google.com/example",
            },
            "2026-06",
            "Filtering",
            SPECS,
        )

        self.assertIsNotNone(record)
        self.assertEqual(record["workflow"], "Mambu Collection Filtering")
        self.assertEqual(record["file_name"], "Mambu_Collection_Filtered.xlsx")
        self.assertEqual(record["drive_file_id"], "drive-file-123")
        self.assertEqual(record["size_bytes"], 4096)
        self.assertTrue(record["drive_saved"])

    def test_recovers_file_after_workflow_label_was_renamed(self):
        record = _record_from_drive_file(
            {
                "id": "drive-file-456",
                "name": (
                    "2026-06_Old Mambu Filter Name_cafebabe_"
                    "Mambu_Collection_Filtered.xlsx"
                ),
            },
            "2026-06",
            "Filtering",
            SPECS,
        )

        self.assertIsNotNone(record)
        self.assertEqual(record["workflow"], "Mambu Collection Filtering")

    def test_skips_unrecognized_files(self):
        record = _record_from_drive_file(
            {"id": "other-file", "name": "notes.txt"},
            "2026-06",
            "Filtering",
            SPECS,
        )
        self.assertIsNone(record)


if __name__ == "__main__":
    unittest.main()
