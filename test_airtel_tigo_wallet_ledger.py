import unittest
import zipfile
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory

from collection_filtering import (
    AIRTEL_TIGO_WALLETS,
    build_airtel_tigo_wallet_ledger_workbook,
    calculate_airtel_tigo_wallet_ledger,
)


class AirtelTigoWalletLedgerTests(unittest.TestCase):
    def test_positive_delayed_transaction_is_available_funds_credit(self):
        metrics = calculate_airtel_tigo_wallet_ledger({
            "opening_balance": "100",
            "total_repayment": "40",
            "interest_received": "10",
            "charges": "5",
            "total_disbursement": "60",
            "delayed_transaction": "15",
        })

        self.assertEqual(metrics["available_funds"], Decimal("165"))
        self.assertEqual(metrics["total_debit"], Decimal("65"))
        self.assertEqual(metrics["wallet_balance"], Decimal("100"))

    def test_negative_delayed_transaction_is_debit(self):
        metrics = calculate_airtel_tigo_wallet_ledger({
            "opening_balance": "100",
            "total_repayment": "40",
            "interest_received": "10",
            "charges": "5",
            "total_disbursement": "60",
            "delayed_transaction": "-15",
        })

        self.assertEqual(metrics["available_funds"], Decimal("150"))
        self.assertEqual(metrics["total_debit"], Decimal("80"))
        self.assertEqual(metrics["wallet_balance"], Decimal("70"))

    def test_workbook_contains_summary_and_all_five_wallet_sheets(self):
        wallet_inputs = {
            name: {
                "opening_balance": "100",
                "total_repayment": "20",
                "interest_received": "5",
                "charges": "2",
                "total_disbursement": "50",
                "delayed_transaction": "0",
            }
            for name, _number in AIRTEL_TIGO_WALLETS
        }
        with TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "Airtel_Tigo_Wallet_vs_Ledger.xlsx"
            metrics = build_airtel_tigo_wallet_ledger_workbook(output, wallet_inputs)

            self.assertEqual(metrics["wallet_count"], 5)
            self.assertEqual(metrics["total_wallet_balance"], Decimal("365"))
            with zipfile.ZipFile(output) as workbook:
                workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
                first_wallet_xml = workbook.read("xl/worksheets/sheet2.xml").decode("utf-8")
            self.assertIn("Summary", workbook_xml)
            for wallet_name, _number in AIRTEL_TIGO_WALLETS:
                self.assertIn(wallet_name, workbook_xml)
            self.assertNotIn("Type", first_wallet_xml)
            self.assertNotIn("Notes", first_wallet_xml)
            self.assertIn("Prepared By:", first_wallet_xml)
            self.assertIn("Reviewed By:", first_wallet_xml)
            self.assertIn("Signature:", first_wallet_xml)


if __name__ == "__main__":
    unittest.main()
