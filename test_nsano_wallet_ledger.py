import unittest
from collections import OrderedDict
from decimal import Decimal

from build_reconciliation_template import (
    STYLE_WALLET_BALANCE_LABEL,
    STYLE_WALLET_LINE_ITEM,
    STYLE_WALLET_SIGNATURE,
)
from disbursement_reconciliation import (
    NSANO_WALLET_OTHER_STATUS,
    NSANO_WALLET_TOPUP_COLLECTION_STATUS,
    NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
    build_nsano_transfer_classification_rows,
    build_nsano_wallet_ledger_summary,
    classify_nsano_transfer_purpose,
)


class NsanoWalletLedgerTests(unittest.TestCase):
    def test_wallet_summary_uses_regular_line_items_and_bold_key_rows(self) -> None:
        rows, _metrics = build_nsano_wallet_ledger_summary(
            Decimal("100"),
            Decimal("0"),
            Decimal("0"),
            Decimal("0"),
            "JANUARY 2026",
            "Jan 2026",
            [],
            [],
        )
        styles_by_label = {
            str(row[0].value): row[0].style
            for row in rows
            if hasattr(row[0], "value")
        }

        self.assertEqual(styles_by_label["Top-up through collections"], STYLE_WALLET_LINE_ITEM)
        self.assertEqual(styles_by_label["Transfer to Bank"], STYLE_WALLET_LINE_ITEM)
        self.assertEqual(styles_by_label["TOTAL DEBIT"], STYLE_WALLET_BALANCE_LABEL)
        self.assertEqual(styles_by_label["Prepared By:"], STYLE_WALLET_SIGNATURE)

    def test_approved_collection_purpose_variants_are_included(self) -> None:
        purposes = (
            "Top-up through collections JAN 07A",
            "Top-up collection JAN 07A",
            "Top-up through collectrion JAN 07A",
            "Top up through colletion JAN 07A",
            "top-up through coll. JAN 07A",
        )
        for purpose in purposes:
            with self.subTest(purpose=purpose):
                self.assertEqual(
                    classify_nsano_transfer_purpose(purpose),
                    NSANO_WALLET_TOPUP_COLLECTION_STATUS,
                )

    def test_unrelated_topup_is_not_included(self) -> None:
        self.assertEqual(
            classify_nsano_transfer_purpose("Top-up through merchant JAN 07A"),
            NSANO_WALLET_OTHER_STATUS,
        )

    def test_nonstandard_collection_purpose_is_flagged_as_exception(self) -> None:
        source = OrderedDict(
            [
                ("Date", "2026-01-07"),
                ("Amount", "125.50"),
                ("Purpose", "Top-up through collectrion JAN 07A"),
            ]
        )
        row = build_nsano_transfer_classification_rows([source])[0]

        self.assertEqual(row["Transfer Category"], NSANO_WALLET_TOPUP_COLLECTION_STATUS)
        self.assertEqual(row["TOP-UP Through Collection with Date"], 125.5)
        self.assertEqual(row["Narration Exception"], "Yes")
        self.assertIn("collectrion interpreted as collection", row["Exception Reason"])

    def test_nsano_transfer_to_bank_purpose_variants(self) -> None:
        self.assertEqual(
            classify_nsano_transfer_purpose("Transfer to Bank"),
            NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
        )
        self.assertEqual(
            classify_nsano_transfer_purpose("Transfer Bank"),
            NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
        )
