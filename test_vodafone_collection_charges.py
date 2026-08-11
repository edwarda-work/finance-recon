from collections import OrderedDict
from decimal import Decimal
import unittest

from build_reconciliation_template import (
    Cell,
    Record,
    VODAFONE_CHARGE_DETAIL,
    build_itc_summary_rows,
    summarize_vodafone_charges,
)


def _record(details: str, withdrawn: str) -> Record:
    return Record(
        sheet_values=OrderedDict([
            ("Details", details),
            ("Withdrawn", withdrawn),
        ]),
        raw={},
        source_file="vodafone.csv",
        source_row=1,
        source_dataset="Vodafone Collections",
        direction="",
        key="",
        amount_cents=None,
        datetime_value="",
        phone="",
        channel_type="",
        account="",
    )


def _empty_compare_stats() -> dict:
    return {
        "status_counts": {},
        "status_amounts": {},
        "total": 0,
        "source_amount": Decimal("0"),
    }


class VodafoneCollectionChargeTests(unittest.TestCase):
    def test_charge_summary_uses_pay_bill_charge_withdrawals(self) -> None:
        records = [
            _record("Pay Bill Charge", "-1.25"),
            _record("PAY BILL CHARGE - utility", "2.50"),
            _record("Funds Transfer", "10.00"),
            _record("Pay Bill Charge", "0.00"),
        ]

        summary = summarize_vodafone_charges(records)

        self.assertEqual(summary["count"], 2)
        self.assertEqual(summary["amount"], Decimal("3.75"))

    def test_charge_summary_is_added_to_reconciliation_summary_rows(self) -> None:
        charge_summary = {"count": 2, "amount": Decimal("3.75")}
        rows = build_itc_summary_rows(
            _empty_compare_stats(),
            _empty_compare_stats(),
            _empty_compare_stats(),
            {},
            {},
            charge_summary,
        )

        charge_row = next(row for row in rows if row[0] == VODAFONE_CHARGE_DETAIL)
        self.assertEqual(charge_row[1], 2)
        self.assertIsInstance(charge_row[2], Cell)
        self.assertEqual(charge_row[2].value, Decimal("3.75"))


if __name__ == "__main__":
    unittest.main()
