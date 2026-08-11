import unittest

from disbursement_reconciliation import (
    NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
    classify_nsano_transfer_purpose,
)


class NsanoWalletLedgerTests(unittest.TestCase):
    def test_nsano_transfer_to_bank_purpose_variants(self) -> None:
        self.assertEqual(
            classify_nsano_transfer_purpose("Transfer to Bank"),
            NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
        )
        self.assertEqual(
            classify_nsano_transfer_purpose("Transfer Bank"),
            NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
        )
