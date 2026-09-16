import tempfile
import unittest
import zipfile
from pathlib import Path

from collection_filtering import write_vodafone_wallet_ledger_workbook
from disbursement_reconciliation import (
    ITC_WALLET_NOT_FOUND_STATUS,
    NSANO_WALLET_OTHER_STATUS,
    write_itc_wallet_ledger_workbook,
    write_nsano_collection_ledger_workbook,
    write_nsano_wallet_ledger_workbook,
)
from zenith_collection_reconciliation import write_zenith_wallet_ledger_workbook


def workbook_sheet_names(path: Path) -> str:
    with zipfile.ZipFile(path) as workbook:
        return workbook.read("xl/workbook.xml").decode("utf-8")


def worksheet_xml(path: Path, sheet_number: int) -> str:
    with zipfile.ZipFile(path) as workbook:
        return workbook.read(f"xl/worksheets/sheet{sheet_number}.xml").decode("utf-8")


class WalletLedgerNotFoundSheetTests(unittest.TestCase):
    def test_nsano_collection_and_disbursement_ledgers_have_not_found_sheet(self) -> None:
        not_found = {
            "Purpose": "Unknown wallet movement",
            "Transfer Category": NSANO_WALLET_OTHER_STATUS,
        }
        with tempfile.TemporaryDirectory() as directory:
            disb_path = Path(directory) / "nsano-disb.xlsx"
            collection_path = Path(directory) / "nsano-collection.xlsx"
            write_nsano_wallet_ledger_workbook(
                disb_path, [], ["Amount"], [], ["Amount"], [], ["Purpose"], [], [not_found]
            )
            write_nsano_collection_ledger_workbook(
                collection_path,
                [],
                ["Amount"],
                [],
                ["Amount"],
                [],
                ["Charge"],
                [],
                ["Purpose"],
                [],
                [not_found],
            )

            self.assertIn('name="Nsano Not Found"', workbook_sheet_names(disb_path))
            self.assertIn('name="Nsano Not Found"', workbook_sheet_names(collection_path))
            self.assertIn("Unknown wallet movement", worksheet_xml(disb_path, 4))
            self.assertIn("Unknown wallet movement", worksheet_xml(collection_path, 4))

    def test_itc_disbursement_ledger_has_not_found_sheet(self) -> None:
        not_found = {"Debit Narration": "Unknown debit", "Match Status": ITC_WALLET_NOT_FOUND_STATUS}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "itc.xlsx"
            write_itc_wallet_ledger_workbook(
                path,
                [],
                ["narration"],
                [],
                ["narration"],
                [],
                ["narration"],
                [],
                [],
                [not_found],
            )

            self.assertIn('name="ITC Not Found"', workbook_sheet_names(path))
            self.assertIn("Unknown debit", worksheet_xml(path, 5))

    def test_collection_wallet_ledgers_have_respective_not_found_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            voda_path = Path(directory) / "voda.xlsx"
            zenith_path = Path(directory) / "zenith.xlsx"
            voda_headers = ["Paid In", "Withdrawn", "Details"]
            write_vodafone_wallet_ledger_workbook(
                voda_path,
                [],
                voda_headers,
                [{"Paid In": "", "Withdrawn": "25", "Details": "Unknown debit"}],
            )
            write_zenith_wallet_ledger_workbook(
                zenith_path,
                [],
                ["Description", "Debit", "Credit"],
                [],
            )

            self.assertIn('name="Voda Not Found"', workbook_sheet_names(voda_path))
            self.assertIn('name="Zenith Not Found"', workbook_sheet_names(zenith_path))
            self.assertIn("Unknown debit", worksheet_xml(voda_path, 2))


if __name__ == "__main__":
    unittest.main()
