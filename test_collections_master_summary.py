import unittest
from collections import OrderedDict
from decimal import Decimal

from build_reconciliation_template import (
    ITC_NOT_FOUND_STATUS,
    Record,
    WRITE_OFF_ZENITH_STATUS,
    not_found_detail_rows,
)
from collections_master_summary import (
    _daily_summary_overview_rows,
    _daily_wallet_sheet_rows,
    _write_off_analysis_rows,
    collection_mambu_source_summary_rows,
    master_summary_dashboard_rows,
    master_daily_sheet_specs,
    master_not_found_sheet_specs,
    write_off_summary_rows,
)
from zenith_collection_reconciliation import (
    ZENITH_NOT_FOUND_STATUS,
    ZENITH_WRITE_OFF_STATUS,
)


class CollectionMambuSummaryTests(unittest.TestCase):
    def test_disbursement_master_summary_breaks_down_not_found_by_side(self):
        state = {
            "nsano_disb_result": {
                "nvm": {
                    "status_counts": {"not found": 2},
                    "status_amounts": {"not found": Decimal("30")},
                },
                "mvn": {
                    "status_counts": {"not found": 3},
                    "status_amounts": {"not found": Decimal("40")},
                },
            },
        }

        rows, _merges, _row_heights = master_summary_dashboard_rows(state)

        # The two directional tables sit side by side in the combined view.
        self.assertEqual(rows[37][1].value, "NOT FOUND · WALLET SIDE")
        self.assertEqual(rows[37][11].value, "NOT FOUND · MAMBU SIDE")
        self.assertEqual(rows[39][1].value, "●  Nsano")
        self.assertEqual(rows[39][5].value, 2)
        self.assertEqual(rows[39][7].value, Decimal("30"))
        self.assertEqual(rows[39][11].value, "●  Nsano")
        self.assertEqual(rows[39][15].value, 3)
        self.assertEqual(rows[39][17].value, Decimal("40"))

    def test_master_write_off_analysis_shows_both_amounts_and_variance(self):
        snapshot = {
            "write_off_recon_result": {
                "stats": {
                    "total": 1,
                    "source_amount": Decimal("100"),
                    "wallet_amount": Decimal("105"),
                    "amount_variance": Decimal("-5"),
                    "status_counts": {"ITC": 1},
                    "status_amounts": {"ITC": Decimal("100")},
                    "status_wallet_amounts": {"ITC": Decimal("105")},
                    "status_variances": {"ITC": Decimal("-5")},
                    "investigation_counts": {"Amount Mismatch": 1},
                },
            },
        }

        analysis = _write_off_analysis_rows(snapshot)

        self.assertIsNotNone(analysis)
        rows, _merges = analysis
        itc_row = rows[5]
        self.assertEqual(itc_row[0].value, "Itc")
        self.assertEqual(itc_row[2].value, Decimal("100"))
        self.assertEqual(itc_row[3].value, Decimal("105"))
        self.assertEqual(itc_row[4].value, Decimal("-5"))

    def test_write_off_by_wallet_uses_allocated_write_off_control_amount(self):
        state = {
            "zenith_result": {
                "zvs": {
                    "status_counts": {ZENITH_WRITE_OFF_STATUS: 4},
                    "status_amounts": {ZENITH_WRITE_OFF_STATUS: Decimal("4100")},
                },
            },
            "write_off_recon_result": {
                "stats": {
                    "status_counts": {WRITE_OFF_ZENITH_STATUS: 4},
                    "status_amounts": {WRITE_OFF_ZENITH_STATUS: Decimal("4048.79")},
                },
            },
        }

        zenith = next(row for row in write_off_summary_rows(state) if row["Wallet"] == "Zenith")

        self.assertEqual(zenith["Count"], 4)
        self.assertEqual(zenith["Amount"], Decimal("4048.79"))

    def test_not_found_is_broken_down_by_reconciliation_source(self):
        state = {
            "nsano_result": {
                "mvn": {
                    "not_found": 12,
                    "not_found_amount": Decimal("900"),
                },
            },
            "itc_result": {
                "mvi": {
                    "status_counts": {ITC_NOT_FOUND_STATUS: 9},
                    "status_amounts": {ITC_NOT_FOUND_STATUS: Decimal("700")},
                },
            },
            "zenith_result": {
                "mvz": {
                    "status_counts": {ZENITH_NOT_FOUND_STATUS: 7},
                    "status_amounts": {ZENITH_NOT_FOUND_STATUS: Decimal("500")},
                },
            },
        }

        rows = collection_mambu_source_summary_rows(state)
        not_found_rows = rows[4:]

        self.assertEqual(
            [row["Wallet"] for row in not_found_rows],
            [
                "Not-Found · Nsano",
                "Not-Found · ITC / Vodafone",
                "Not-Found · Zenith",
            ],
        )
        self.assertEqual(sum(row["Count"] for row in not_found_rows), 28)
        self.assertEqual(sum((row["Amount"] for row in not_found_rows), Decimal("0")), Decimal("2100"))

    def test_not_found_details_are_grouped_in_one_collection_sheet(self):
        nsano_mambu = {"Transaction ID": "M-1", "Wallet": "Mambu", "Recon Side": "Mambu → Wallet"}
        nsano_wallet = {"Transaction ID": "N-1", "Wallet": "Nsano", "Recon Side": "Wallet → Mambu"}
        itc_wallet = {"Transaction ID": "I-1", "Wallet": "ITC", "Recon Side": "Wallet → Mambu"}
        snapshot = {
            "nsano_result": {
                "mvn": {"not_found_details": [nsano_mambu]},
                "nvm": {"not_found_details": [nsano_wallet]},
            },
            "itc_result": {
                "mvi": {"not_found_details": []},
                "ivm": {"not_found_details": [itc_wallet]},
                "vvm": {"not_found_details": []},
            },
        }

        specs = master_not_found_sheet_specs(snapshot)

        self.assertEqual([name for name, _groups in specs], ["Collections Not Found"])
        groups = dict(specs[0][1])
        self.assertEqual(groups["NSANO · MAMBU AS SOURCE"], [nsano_mambu])
        self.assertEqual(groups["NSANO · WALLET AS SOURCE"], [nsano_wallet])
        self.assertEqual(groups["ITC / VODAFONE · MAMBU AS SOURCE"], [])
        self.assertEqual(groups["ITC · WALLET AS SOURCE"], [itc_wallet])
        self.assertEqual(groups["VODAFONE · WALLET AS SOURCE"], [])

    def test_not_found_detail_keeps_raw_and_normalized_ids(self):
        record = Record(
            sheet_values=OrderedDict(
                Transaction_ID="TX-001",
                External_Debit_Reference="00001234/EXTRA",
            ),
            raw={},
            source_file="nsano.xlsx",
            source_row=15,
            source_dataset="Nsano",
            direction="COLLECTION",
            key="1234",
            amount_cents=2500,
            datetime_value="2026-06-30 10:00:00",
            phone="233200000001",
            channel_type="W2A",
            account="ACC-1",
        )
        comparison = OrderedDict(
            Match_Status="NOT FOUND",
            Match_Reason="reference not found in Mambu",
        )

        details = not_found_detail_rows(
            [record],
            [comparison],
            "NOT FOUND",
            "Nsano",
            "Wallet → Mambu",
        )

        self.assertEqual(details[0]["Transaction ID"], "TX-001")
        self.assertEqual(details[0]["Raw Reconciliation ID"], "00001234/EXTRA")
        self.assertEqual(details[0]["Reconciliation Key"], "1234")

    def test_daily_wallet_sheets_are_separated_and_totalled(self):
        snapshot = {
            "nsano_result": {
                "daily_summary": {
                    "Nsano": [
                        {"date": "2026-07-01", "amount": Decimal("100"), "charges": Decimal("2")},
                        {"date": "2026-07-02", "amount": Decimal("150"), "charges": Decimal("3")},
                    ],
                },
            },
            "itc_result": {
                "daily_summary": {
                    "ITC": [
                        {"date": "2026-07-01", "amount": Decimal("75"), "charges": Decimal("1.50")},
                    ],
                    "Vodafone": [],
                },
            },
        }

        specs = master_daily_sheet_specs(snapshot)

        self.assertEqual(
            [name for name, _rows in specs],
            ["Daily - Nsano", "Daily - ITC", "Daily - Vodafone"],
        )
        nsano_rows = _daily_wallet_sheet_rows(*specs[0])
        self.assertEqual(nsano_rows[-1][0].value, "TOTAL")
        self.assertEqual(nsano_rows[-1][1].value, Decimal("250"))
        self.assertEqual(nsano_rows[-1][2].value, Decimal("5"))

        overview_rows, overview_merges = _daily_summary_overview_rows(specs)
        self.assertEqual(overview_rows[0][0].value, "DAILY TRANSACTION SUMMARY")
        self.assertEqual(overview_rows[2][0].value, "Nsano")
        self.assertEqual(overview_rows[2][2].value, Decimal("100"))
        self.assertEqual(overview_merges, ["A1:D1"])


if __name__ == "__main__":
    unittest.main()
