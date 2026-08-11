import unittest
from collections import OrderedDict
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory

from build_reconciliation_template import (
    Record,
    WRITE_OFF_AMBIGUOUS_STATUS,
    WRITE_OFF_ITC_STATUS,
    WRITE_OFF_ZENITH_STATUS,
    _load_write_off_source,
    _load_write_off_vodafone_target,
    _load_write_off_zenith_target,
    _write_off_mismatch_analysis_rows,
    compare_write_off_to_collection_sources,
    summarize_write_off_recon,
)
from collection_filtering import build_vodafone_collection_cleanup_workbook


def record(
    *,
    dataset: str,
    key: str,
    amount: str,
    channel: str = "",
    date: str = "2026-01-15",
) -> Record:
    return Record(
        sheet_values=OrderedDict(),
        raw={},
        source_file=f"{dataset}.xlsx",
        source_row=2,
        source_dataset=dataset,
        direction="",
        key=key,
        amount_cents=int(Decimal(amount) * 100),
        datetime_value=date,
        phone="",
        channel_type=channel,
        account="",
    )


class WriteOffAllocationTests(unittest.TestCase):
    def test_summary_variance_includes_missing_wallet_amount(self):
        write_off = record(
            dataset="Write Off",
            key="MISSING",
            amount="480.34",
            channel="ITC",
        )

        row = compare_write_off_to_collection_sources(
            [write_off], [], [], [], []
        )[0]
        stats = summarize_write_off_recon([row])

        self.assertEqual(stats["status_amounts"]["Not Found"], Decimal("480.34"))
        self.assertEqual(stats["status_wallet_amounts"]["Not Found"], Decimal("0"))
        self.assertEqual(stats["status_variances"]["Not Found"], Decimal("480.34"))

    def test_amount_mismatch_analysis_groups_transactions_by_wallet(self):
        itc_write_off = record(
            dataset="Write Off", key="ITC-M", amount="100", channel="ITC"
        )
        itc = record(dataset="ITC", key="ITC-M", amount="105")
        vodafone_write_off = record(
            dataset="Write Off", key="VF-M", amount="200", channel="Vodafone"
        )
        vodafone = record(dataset="Vodafone", key="VF-M", amount="190")
        comparisons = compare_write_off_to_collection_sources(
            [itc_write_off, vodafone_write_off],
            [],
            [itc],
            [],
            [vodafone],
        )

        rows, _merges = _write_off_mismatch_analysis_rows(comparisons)

        self.assertEqual(rows[4][0].value, "ITC")
        self.assertEqual(rows[4][1].value, 1)
        self.assertEqual(rows[4][4].value, Decimal("-5"))
        self.assertEqual(rows[5][0].value, "Vodafone Collections")
        self.assertEqual(rows[5][4].value, Decimal("10"))

    def test_amount_mismatch_analysis_separates_not_found_breakdown(self):
        amount_mismatch = record(
            dataset="Write Off", key="ITC-M", amount="100", channel="ITC"
        )
        itc = record(dataset="ITC", key="ITC-M", amount="105")
        missing = record(
            dataset="Write Off", key="ZEN-MISSING", amount="75", channel="Zenith"
        )
        comparisons = compare_write_off_to_collection_sources(
            [amount_mismatch, missing], [], [itc], [], []
        )

        rows, _merges = _write_off_mismatch_analysis_rows(comparisons)

        mismatch_total = next(
            row for row in rows
            if getattr(row[0], "value", None) == "TOTAL"
            and getattr(row[1], "value", None) == 1
            and getattr(row[2], "value", None) == Decimal("100")
        )
        self.assertEqual(mismatch_total[4].value, Decimal("-5"))

        not_found_title_index = next(
            index
            for index, row in enumerate(rows)
            if getattr(row[0], "value", None) == "WRITE-OFF NOT FOUND ANALYSIS"
        )
        zenith_breakdown = next(
            row
            for row in rows[not_found_title_index:]
            if getattr(row[0], "value", None) == WRITE_OFF_ZENITH_STATUS
        )
        self.assertEqual(zenith_breakdown[1].value, 1)
        self.assertEqual(zenith_breakdown[2].value, Decimal("75"))

    def test_vodafone_same_receipt_uses_paid_in_row_not_charge_row(self):
        write_off = record(
            dataset="Write Off",
            key="11927711719",
            amount="899",
            channel="Vodafone Collection",
        )
        charge = record(
            dataset="Vodafone Collections",
            key="11927711719",
            amount="0",
        )
        charge.amount_cents = None
        paid_in = record(
            dataset="Vodafone Collections",
            key="11927711719",
            amount="899",
        )

        row = compare_write_off_to_collection_sources(
            [write_off], [], [], [], [charge, paid_in]
        )[0]

        self.assertEqual(row["Match Status"], "Vodafone Collections")
        self.assertEqual(row["Wallet Amount (GHC)"], Decimal("899"))
        self.assertEqual(row["Amount Variance (GHC)"], Decimal("0"))

    def test_vodafone_investigation_uses_paid_in_not_withdrawn(self):
        with TemporaryDirectory() as tmpdir:
            raw_path = Path(tmpdir) / "vodafone.csv"
            cleaned_path = Path(tmpdir) / "Voda_Collection_Cleaned.xlsx"
            raw_path.write_text(
                "Receipt No.,Completion Time,Initiation Time,Details,Transaction Status,"
                "Currency,Paid In,Withdrawn,Balance,Reason Type,Opposite Party\n"
                "VF-1,2026-01-15,2026-01-15,Collection,Completed,GHS,"
                "125.50,999.00,1000,Payment,Customer\n",
                encoding="utf-8",
            )
            build_vodafone_collection_cleanup_workbook(cleaned_path, [raw_path])

            records, _headers = _load_write_off_vodafone_target(cleaned_path)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].key, "VF-1")
        self.assertEqual(records[0].amount_decimal, Decimal("125.5"))

    def test_non_zenith_route_never_includes_zenith_as_candidate(self):
        write_off = record(
            dataset="Write Off",
            key="ITC-ID",
            amount="100",
            channel="Vodafone Collection",
            date="2026-01-15",
        )
        itc = record(dataset="ITC", key="ITC-ID", amount="100")
        zenith = record(
            dataset="Zenith",
            key="2026-01-15|10000",
            amount="100",
            date="2026-01-15",
        )

        row = compare_write_off_to_collection_sources(
            [write_off], [], [itc], [zenith], []
        )[0]

        self.assertEqual(row["Candidate Wallets"], "ITC")
        self.assertEqual(row["Match Status"], WRITE_OFF_ITC_STATUS)
        self.assertEqual(row["Investigation Status"], "Channel Mismatch")
        self.assertNotIn("Zenith", row["Match Reason"])
        self.assertEqual(row["Zenith Date Check"], "Not Applicable")

    def test_zenith_route_ignores_identifier_matches_in_other_wallets(self):
        write_off = record(
            dataset="Write Off",
            key="DUPLICATE-ID",
            amount="50",
            channel="Zenith Collection",
            date="2026-01-27",
        )
        nsano = record(dataset="Nsano", key="DUPLICATE-ID", amount="50")
        itc = record(dataset="ITC", key="DUPLICATE-ID", amount="50")

        row = compare_write_off_to_collection_sources(
            [write_off], [nsano], [itc], [], []
        )[0]

        self.assertEqual(row["Routed Wallet"], WRITE_OFF_ZENITH_STATUS)
        self.assertEqual(row["Match Status"], "Not Found")
        self.assertEqual(row["Candidate Wallets"], "")
        self.assertNotIn("Ambiguous", row["Investigation Status"])

    def test_zenith_amount_mismatch_is_shown_explicitly(self):
        write_off = record(
            dataset="Write Off", key="WO-A", amount="100", channel="Zenith", date="2026-01-15"
        )
        zenith = record(
            dataset="Zenith", key="2026-01-15|10500", amount="105", date="2026-01-15"
        )

        row = compare_write_off_to_collection_sources(
            [write_off], [], [], [zenith], []
        )[0]

        self.assertEqual(row["Zenith Date Check"], "Matched")
        self.assertEqual(row["Zenith Amount Check"], "Mismatch")
        self.assertEqual(row["Zenith Candidate Amount (GHC)"], Decimal("105"))
        self.assertEqual(row["Wallet Amount (GHC)"], Decimal("105"))
        self.assertEqual(row["Amount Variance (GHC)"], Decimal("-5"))
        self.assertEqual(row["Match Status"], WRITE_OFF_ZENITH_STATUS)
        self.assertEqual(row["Investigation Status"], "Amount Mismatch")
        self.assertIn("Zenith Amount Mismatch", row["Zenith Mismatch Detail"])

        stats = summarize_write_off_recon([row])
        self.assertEqual(stats["status_counts"][WRITE_OFF_ZENITH_STATUS], 1)
        self.assertEqual(stats["not_found"], 0)

    def test_zenith_date_mismatch_is_shown_explicitly(self):
        write_off = record(
            dataset="Write Off", key="WO-D", amount="100", channel="Zenith", date="2026-01-15"
        )
        zenith = record(
            dataset="Zenith", key="2026-01-16|10000", amount="100", date="2026-01-16"
        )

        row = compare_write_off_to_collection_sources(
            [write_off], [], [], [zenith], []
        )[0]

        self.assertEqual(row["Zenith Date Check"], "Mismatch")
        self.assertEqual(row["Zenith Amount Check"], "Matched")
        self.assertEqual(row["Zenith Candidate Date"], "2026-01-16")
        self.assertIn("Zenith Date Mismatch", row["Zenith Mismatch Detail"])

    def test_write_off_loader_uses_zenith_key_date_and_amount_fields(self):
        with TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "write_off.csv"
            path.write_text(
                "Date,Date (KEY),Amount,Amount (KEY),Channel,Identifier (Key)\n"
                "2026-01-20,2026-01-15,999.00,4048.79,Zenith,WO-1\n",
                encoding="utf-8",
            )

            records, _headers = _load_write_off_source(path)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].datetime_value, "2026-01-15")
        self.assertEqual(records[0].amount_decimal, Decimal("4048.79"))

    def test_zenith_investigation_loader_uses_create_date_and_credit(self):
        with TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "zenith.csv"
            path.write_text(
                "ZENITH BANK STATEMENT\n"
                "Opening Balance,12000.00\n"
                "Description,Value Date (Entry Date),Create Date,Amount,Credit\n"
                "Collection,2026-01-20,2026-01-15,999.00,4048.79\n",
                encoding="utf-8",
            )

            records, _headers = _load_write_off_zenith_target(path)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].key, "2026-01-15|404879")
        self.assertEqual(records[0].amount_decimal, Decimal("4048.79"))

    def test_channel_routes_duplicate_identifier_to_itc_and_reports_variance(self):
        write_off = record(dataset="Write Off", key="TX-1", amount="100", channel="ITC")
        nsano = record(dataset="Nsano", key="TX-1", amount="100")
        itc = record(dataset="ITC", key="TX-1", amount="105")

        row = compare_write_off_to_collection_sources(
            [write_off], [nsano], [itc], [], []
        )[0]

        self.assertEqual(row["Match Status"], WRITE_OFF_ITC_STATUS)
        self.assertEqual(row["Candidate Wallets"], "Nsano, ITC")
        self.assertEqual(row["Wallet Amount (GHC)"], Decimal("105"))
        self.assertEqual(row["Amount Variance (GHC)"], Decimal("-5"))
        self.assertEqual(row["Investigation Status"], "Amount Mismatch")

    def test_unrouted_duplicate_identifier_is_ambiguous(self):
        write_off = record(dataset="Write Off", key="TX-2", amount="50")
        nsano = record(dataset="Nsano", key="TX-2", amount="50")
        itc = record(dataset="ITC", key="TX-2", amount="50")

        row = compare_write_off_to_collection_sources(
            [write_off], [nsano], [itc], [], []
        )[0]

        self.assertEqual(row["Match Status"], WRITE_OFF_AMBIGUOUS_STATUS)
        self.assertEqual(row["Wallet Amount (GHC)"], "")
        stats = summarize_write_off_recon([row])
        self.assertEqual(
            stats["ambiguous_breakdown"],
            {"Channel: Blank / Unrecognized · Candidates: Nsano, ITC": 1},
        )

    def test_zenith_channel_allocates_by_date_and_amount(self):
        write_off = record(dataset="Write Off", key="WO-3", amount="4048.79", channel="Zenith")
        zenith = record(dataset="Zenith", key="2026-01-15|404879", amount="4048.79")

        row = compare_write_off_to_collection_sources(
            [write_off], [], [], [zenith], []
        )[0]
        stats = summarize_write_off_recon([row])

        self.assertEqual(row["Match Status"], WRITE_OFF_ZENITH_STATUS)
        self.assertEqual(stats["status_counts"][WRITE_OFF_ZENITH_STATUS], 1)
        self.assertEqual(stats["status_amounts"][WRITE_OFF_ZENITH_STATUS], Decimal("4048.79"))
        self.assertEqual(stats["status_variances"][WRITE_OFF_ZENITH_STATUS], Decimal("0"))


if __name__ == "__main__":
    unittest.main()
