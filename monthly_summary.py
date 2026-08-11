#!/usr/bin/env python3
"""Monthly master summary workbook for completed reconciliation runs."""

from __future__ import annotations

import os
import zipfile
import datetime as dt
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Mapping

from build_reconciliation_template import (
    FAST_XLSX_COMPRESSLEVEL,
    ITC_MAMBU_STATUS,
    ITC_NOT_FOUND_STATUS,
    ITC_STATUS,
    ITC_UNIDENTIFIED_STATUS,
    ITC_UPSALE_STATUS,
    ITC_WRITE_OFF_STATUS,
    STYLE_GRAN_COUNT,
    STYLE_GRAN_DATA,
    STYLE_GRAN_SECTION,
    STYLE_GRAN_SECTION_COL,
    STYLE_GRAN_TOTAL,
    STYLE_GRAN_TOTAL_MONEY,
    STYLE_GRAN_TOTAL_NUM,
    STYLE_GRAN_WORKFLOW,
    STYLE_HEADER,
    STYLE_MONEY,
    STYLE_PERCENT,
    STYLE_TABLE_SECTION,
    STYLE_TITLE,
    STYLE_DASH_CARD_DETAIL_LARGE,
    STYLE_DASH_CARD_LABEL_LARGE,
    STYLE_DASH_CARD_TOP_DARK,
    STYLE_DASH_CARD_TOP_PRIMARY,
    STYLE_DASH_CARD_VALUE_ALERT_FULL,
    STYLE_DASH_CARD_VALUE_FULL,
    STYLE_DASH_EYEBROW,
    STYLE_DASH_FOOTER,
    STYLE_DASH_LOGO,
    STYLE_DASH_META_LABEL,
    STYLE_DASH_META_VALUE,
    STYLE_DASH_STATUS,
    STYLE_DASH_TITLE,
    VODA_COLL_STATUS,
    VODAFONE_WRITE_OFF_STATUS,
    WRITE_OFF_ITC_STATUS,
    WRITE_OFF_NOT_FOUND_STATUS,
    WRITE_OFF_NSANO_STATUS,
    WRITE_OFF_VODAFONE_STATUS,
    WRITE_OFF_ZENITH_STATUS,
    Cell,
    dashboard_grid,
    dashboard_merge,
    app_xml,
    column_letter,
    content_types_xml,
    core_xml,
    root_rels_xml,
    styles_xml,
    workbook_rels_xml,
    workbook_xml,
    write_worksheet,
)
from disbursement_reconciliation import (
    DISB_MAMBU_STATUS,
    DISB_NOT_FOUND_STATUS,
    DISB_NSANO_STATUS,
    ITC_DISB_NOT_FOUND_STATUS,
    ITC_DISB_REFERRAL_BONUS_STATUS,
    ITC_DISB_SAVINGS_REWARD_STATUS,
    ITC_DISB_STATUS,
    ITC_DISB_UPSALES_REFOUND_STATUS,
    ITC_WALLET_CREDIT_TRANSFER_STATUS,
    ITC_WALLET_DISBURSEMENT_STATUS,
    ITC_WALLET_NOT_FOUND_STATUS,
    ITC_WALLET_PREPAID_REVERSAL_STATUS,
    ITC_WALLET_REFERRAL_AWARD_STATUS,
    ITC_WALLET_REVERSAL_STATUS,
    ITC_WALLET_SAVINGS_STATUS,
    ITC_WALLET_SETTLEMENT_STATUS,
    ITC_WALLET_TRANSFER_TO_WALLET_STATUS,
    ITC_WALLET_TRANSFERS_FROM_BANK_STATUS,
    ITC_WALLET_UPSALES_REFUND_STATUS,
    MTN_MANUAL_MAMBU_STATUS,
    MTN_MANUAL_MATCHED_STATUS,
    MTN_MANUAL_NOT_FOUND_STATUS,
    MTN_MANUAL_REFUND_STATUS,
    NSANO_WALLET_BANK_TO_WALLET_STATUS,
    NSANO_WALLET_OTHER_STATUS,
    NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS,
    NSANO_WALLET_TOPUP_COLLECTION_STATUS,
    NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
    VODAFONE_MANUAL_MAMBU_STATUS,
    VODAFONE_MANUAL_MATCHED_STATUS,
    VODAFONE_MANUAL_NOT_FOUND_STATUS,
    VODAFONE_MANUAL_REFUND_STATUS,
)
from zenith_collection_reconciliation import (
    ZENITH_MAMBU_STATUS,
    ZENITH_MATCHED_STATUS,
    ZENITH_NOT_FOUND_STATUS,
    ZENITH_UNIDENTIFIED_STATUS,
    ZENITH_WRITE_OFF_STATUS,
)
from collections_master_summary import (
    _collection_exceptions_panel,
    _ledger_panel,
    _row_summary,
    _sum_summaries,
    charges_summary_rows,
    collection_mambu_source_summary_rows,
    collection_source_summary_rows,
    disbursement_mambu_summary_rows,
    disbursement_not_found_summary_rows,
    disbursement_wallet_summary_rows,
    not_found_summary_rows,
    refund_summary_rows,
    unidentified_summary_rows,
    write_off_summary_rows,
)


MONTHLY_RESULT_KEYS = (
    "nsano_result",
    "itc_result",
    "zenith_result",
    "write_off_recon_result",
    "nsano_disb_result",
    "nsano_collection_ledger_result",
    "nsano_wallet_ledger_result",
    "itc_disb_result",
    "itc_wallet_ledger_result",
    "mtn_manual_disb_result",
    "vodafone_manual_disb_result",
)

MONTHLY_SUMMARY_LAYOUT_VERSION = "dashboard-v5-nsano-collection-ledger"

RUN_STATUS_SPECS = (
    ("nsano_result", "Collection", "Nsano Coll Recon"),
    ("itc_result", "Collection", "ITC/Voda Coll Recon"),
    ("zenith_result", "Collection", "Zenith Coll Recon"),
    ("write_off_recon_result", "Collection", "Write-off Recon"),
    ("nsano_disb_result", "Disbursement", "Nsano Wallet vs Mambu"),
    ("nsano_collection_ledger_result", "Ledger", "Nsano Collections vs Ledger"),
    ("nsano_wallet_ledger_result", "Ledger", "Nsano Disb Wallet vs Ledger"),
    ("itc_disb_result", "Disbursement", "ITC Wallet vs Mambu"),
    ("itc_wallet_ledger_result", "Disbursement", "ITC Wallet vs Ledger"),
    ("mtn_manual_disb_result", "Disbursement", "MTN Manual Disbursement"),
    ("vodafone_manual_disb_result", "Disbursement", "Vodafone Manual Disbursement"),
)

RUN_STATUS_HEADERS = ["Area", "Workflow", "Status", "Rows", "Amount (GHC)", "Notes"]
RECON_HEADERS = [
    "Workflow",
    "Comparison",
    "Total Count",
    "Matched Count",
    "Not Found Count",
    "Match Rate",
    "Source Amount (GHC)",
    "Matched Amount (GHC)",
    "Not Found Amount (GHC)",
    "Notes",
]
WALLET_HEADERS = ["Wallet", "Group", "Line Item", "Count", "Amount (GHC)", "Notes"]


def monthly_summary_snapshot(state: Mapping[str, Any]) -> dict[str, Any]:
    """Return only the completed recon results needed for the monthly summary."""
    snapshot: dict[str, Any] = {}
    for key in MONTHLY_RESULT_KEYS:
        if key not in state:
            continue
        value = state[key]
        if isinstance(value, Mapping):
            result: dict[str, Any] = {}
            for field, field_value in value.items():
                if field == "output_bytes":
                    continue
                if isinstance(field_value, Mapping) and "not_found_details" in field_value:
                    result[field] = {
                        nested_key: nested_value
                        for nested_key, nested_value in field_value.items()
                        if nested_key != "not_found_details"
                    }
                else:
                    result[field] = field_value
            snapshot[key] = result
        else:
            snapshot[key] = value
    return snapshot


def monthly_summary_status_rows(snapshot: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [
        _run_status_dict(key, area, workflow, snapshot.get(key))
        for key, area, workflow in RUN_STATUS_SPECS
    ]


def build_monthly_summary_workbook(output_path: Path, snapshot: Mapping[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Monthly Dashboard", "Run Status", "Collection Summary", "Disbursement Summary", "Wallet Summary", "Granular Detail"]

    dashboard_rows, dashboard_merges, dashboard_heights = monthly_summary_dashboard_rows(snapshot)

    run_status_rows = _sheet_rows(
        "MONTHLY RECONCILIATION SUMMARY",
        RUN_STATUS_HEADERS,
        _run_status_rows(snapshot),
    )
    collection_rows = _sheet_rows(
        "COLLECTION SUMMARY",
        RECON_HEADERS,
        _collection_summary_rows(snapshot),
    )
    disbursement_rows = _sheet_rows(
        "DISBURSEMENT SUMMARY",
        RECON_HEADERS,
        _disbursement_summary_rows(snapshot),
    )
    wallet_rows = _sheet_rows(
        "WALLET LEDGER SUMMARY",
        WALLET_HEADERS,
        _wallet_summary_rows(snapshot),
    )
    gran_rows, gran_merges = _granular_sheet_rows(snapshot)

    summary_sheets = [
        (run_status_rows, [18, 32, 16, 14, 18, 74]),
        (collection_rows, [28, 30, 14, 16, 16, 14, 20, 20, 20, 60]),
        (disbursement_rows, [30, 32, 14, 16, 16, 14, 20, 20, 20, 60]),
        (wallet_rows, [18, 22, 34, 14, 18, 64]),
    ]

    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=FAST_XLSX_COMPRESSLEVEL, allowZip64=True) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheet_names)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml(sheet_names))
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheet_names)))
        zf.writestr("xl/styles.xml", styles_xml())

        write_worksheet(
            zf,
            "xl/worksheets/sheet1.xml",
            dashboard_rows,
            len(dashboard_rows),
            21,
            [2.2, 3.0, 11.5, 11.5, 11.5, 2.2, 11.5, 11.5, 11.5, 11.5, 2.2, 3.0, 11.5, 11.5, 11.5, 2.2, 11.5, 11.5, 11.5, 11.5, 2.2],
            freeze_top_row=False,
            autofilter=False,
            merges=dashboard_merges,
            row_heights=dashboard_heights,
            show_gridlines=False,
            zoom_scale=85,
            page_orientation="landscape",
            fit_to_width=1,
        )

        for idx, (rows, widths) in enumerate(summary_sheets, start=2):
            col_count = len(widths)
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{idx}.xml",
                rows,
                len(rows),
                col_count,
                widths,
                freeze_top_row=False,
                autofilter=False,
                merges=[f"A1:{column_letter(col_count)}1"],
                style_func=_monthly_style,
            )

        gran_rows_list = list(gran_rows)
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            gran_rows_list,
            len(gran_rows_list),
            3,
            [46, 14, 20],
            freeze_top_row=False,
            autofilter=False,
            merges=gran_merges or None,
            style_func=None,
        )

    os.replace(temp_path, output_path)


def build_monthly_summary_bytes(snapshot: Mapping[str, Any], output_path: Path) -> bytes:
    build_monthly_summary_workbook(output_path, snapshot)
    return output_path.read_bytes()


def monthly_summary_dashboard_rows(
    snapshot: Mapping[str, Any],
) -> tuple[list[list[Any]], list[str], dict[int, float]]:
    """Build the executive dashboard that opens first in every monthly export.

    Large-format layout: twelve top-line KPIs, a full wallet-vs-Mambu analysis
    panel for each of Collection and Disbursement, and a per-wallet breakdown
    grid for the exception-style categories (not-found, refund, unidentified,
    charges, write-off).
    """
    row_count = 112
    rows = dashboard_grid(row_count, 21)
    merges: list[str] = []
    row_heights: dict[int, float] = {
        1: 8, 2: 18, 3: 24, 4: 24, 5: 18, 6: 22, 7: 10,
        14: 10, 21: 10, 28: 10, 35: 10,
    }

    collection_wallet = collection_source_summary_rows(snapshot)
    collection_mambu = collection_mambu_source_summary_rows(snapshot)
    disbursement_wallet = disbursement_wallet_summary_rows(snapshot)
    disbursement_mambu = disbursement_mambu_summary_rows(snapshot)
    refunds = refund_summary_rows(snapshot)
    unidentified = unidentified_summary_rows(snapshot)
    not_found_collections = not_found_summary_rows(snapshot)
    not_found_disbursement = disbursement_not_found_summary_rows(snapshot)
    charges = charges_summary_rows(snapshot)
    write_offs = write_off_summary_rows(snapshot)

    collection_wallet_total = _sum_summaries(*(_row_summary(row) for row in collection_wallet))
    collection_mambu_total = _sum_summaries(*(_row_summary(row) for row in collection_mambu))
    disb_wallet_total = _sum_summaries(*(_row_summary(row) for row in disbursement_wallet))
    disb_mambu_total = _sum_summaries(*(_row_summary(row) for row in disbursement_mambu))
    refund_total = _sum_summaries(*(_row_summary(row) for row in refunds))
    unidentified_total = _sum_summaries(*(_row_summary(row) for row in unidentified))
    not_found_collections_total = _sum_summaries(*(_row_summary(row) for row in not_found_collections))
    not_found_disbursement_total = _sum_summaries(*(_row_summary(row) for row in not_found_disbursement))
    charges_total = _sum_summaries(*(_row_summary(row) for row in charges))
    write_off_total = _sum_summaries(*(_row_summary(row) for row in write_offs))

    collection_variance_count = int(collection_wallet_total["count"]) - int(collection_mambu_total["count"])
    collection_variance_amount = collection_wallet_total["amount"] - collection_mambu_total["amount"]
    disb_variance_count = int(disb_wallet_total["count"]) - int(disb_mambu_total["count"])
    disb_variance_amount = disb_wallet_total["amount"] - disb_mambu_total["amount"]
    collection_variance = {"count": collection_variance_count, "amount": collection_variance_amount}
    disb_variance = {"count": disb_variance_count, "amount": disb_variance_amount}
    variance_flagged = bool(collection_variance_count or collection_variance_amount or disb_variance_count or disb_variance_amount)

    # Header
    dashboard_merge(rows, merges, 2, 2, 5, 3, "₵", STYLE_DASH_LOGO)
    dashboard_merge(rows, merges, 2, 4, 2, 11, "FINANCE OPERATIONS", STYLE_DASH_EYEBROW)
    dashboard_merge(rows, merges, 3, 4, 5, 14, "Monthly Reconciliation Summary", STYLE_DASH_TITLE)
    dashboard_merge(rows, merges, 2, 15, 2, 17, "GENERATED", STYLE_DASH_META_LABEL)
    dashboard_merge(rows, merges, 3, 15, 5, 17, dt.date.today().strftime("%d %b %Y"), STYLE_DASH_META_VALUE)
    dashboard_merge(
        rows, merges, 2, 18, 5, 20,
        "●  Variance flagged" if variance_flagged else "●  Reconciled",
        STYLE_DASH_STATUS,
    )
    dashboard_merge(
        rows, merges, 6, 2, 6, 20,
        "Collections and disbursements reconciled against Mambu core banking  •  All amounts in GH¢",
        STYLE_DASH_META_LABEL,
    )

    def add_kpi(top_row: int, start_col: int, end_col: int, label: str, value: Any, value_style: int, detail: str, *, dark_top: bool) -> None:
        dashboard_merge(
            rows, merges, top_row, start_col, top_row, end_col, "",
            STYLE_DASH_CARD_TOP_DARK if dark_top else STYLE_DASH_CARD_TOP_PRIMARY,
        )
        dashboard_merge(rows, merges, top_row + 1, start_col, top_row + 1, end_col, label.upper(), STYLE_DASH_CARD_LABEL_LARGE)
        dashboard_merge(rows, merges, top_row + 2, start_col, top_row + 3, end_col, value, value_style)
        dashboard_merge(rows, merges, top_row + 4, start_col, top_row + 4, end_col, detail, STYLE_DASH_CARD_DETAIL_LARGE)
        dashboard_merge(rows, merges, top_row + 5, start_col, top_row + 5, end_col, "", STYLE_DASH_CARD_DETAIL_LARGE)

    kpi_specs = [
        ("Total Collections Mambu", collection_mambu_total, STYLE_DASH_CARD_VALUE_FULL, False),
        ("Total Collections Wallet", collection_wallet_total, STYLE_DASH_CARD_VALUE_FULL, False),
        ("Total Disbursement Mambu", disb_mambu_total, STYLE_DASH_CARD_VALUE_FULL, False),
        ("Total Disbursement Wallet", disb_wallet_total, STYLE_DASH_CARD_VALUE_FULL, False),
        ("Total Refund", refund_total, STYLE_DASH_CARD_VALUE_FULL, False),
        ("Total Collections Variance", collection_variance, STYLE_DASH_CARD_VALUE_ALERT_FULL, True),
        ("Total Disbursement Variance", disb_variance, STYLE_DASH_CARD_VALUE_ALERT_FULL, True),
        ("Total Unidentified", unidentified_total, STYLE_DASH_CARD_VALUE_FULL, True),
        ("Total Not-Found Collections", not_found_collections_total, STYLE_DASH_CARD_VALUE_FULL, True),
        ("Total Not-Found Disbursement", not_found_disbursement_total, STYLE_DASH_CARD_VALUE_FULL, True),
        ("Total Write-Off", write_off_total, STYLE_DASH_CARD_VALUE_FULL, True),
        ("Total Charges", charges_total, STYLE_DASH_CARD_VALUE_FULL, True),
    ]
    kpi_positions = [
        (8, 2, 7), (8, 8, 14), (8, 15, 20),
        (15, 2, 7), (15, 8, 14), (15, 15, 20),
        (22, 2, 7), (22, 8, 14), (22, 15, 20),
        (29, 2, 7), (29, 8, 14), (29, 15, 20),
    ]
    for (top_row, start_col, end_col), (label, summary, style, dark_top) in zip(kpi_positions, kpi_specs):
        count = int(summary["count"])
        detail = f"{count:+,} transaction gap" if "Variance" in label else f"{count:,} txns"
        add_kpi(top_row, start_col, end_col, label, summary["amount"], style, detail, dark_top=dark_top)

    for top_row in (8, 15, 22, 29):
        row_heights[top_row] = 6
        row_heights[top_row + 1] = 24
        row_heights[top_row + 2] = 30
        row_heights[top_row + 3] = 30
        row_heights[top_row + 4] = 24
        row_heights[top_row + 5] = 10

    # Analysis for Collection and Disbursement — full wallet-vs-Mambu ledger panels.
    _ledger_panel(
        rows, merges, row_heights, 39, "COLLECTIONS", "COLLECTIONS ANALYSIS",
        collection_wallet, collection_mambu, "COLLECTIONS PER WALLET", "MAMBU COLLECTIONS PER WALLET",
        show_variance=True, large_text=True,
    )
    _ledger_panel(
        rows, merges, row_heights, 59, "DISBURSEMENTS", "DISBURSEMENTS ANALYSIS",
        disbursement_wallet, disbursement_mambu, "DISBURSEMENT PER WALLET", "MAMBU PER WALLETS",
        show_variance=True, large_text=True,
    )

    # Breakdown of each exception category per wallet.
    _collection_exceptions_panel(
        rows, merges, row_heights, 79,
        [
            ("Not-Found Collections by Wallet", not_found_collections),
            ("Not-Found Disbursement by Wallet", not_found_disbursement),
            ("Refund by Wallet", refunds),
            ("Unidentified by Wallet", unidentified),
            ("Charges by Wallet", charges),
            ("Write-Off by Wallet", write_offs),
        ],
        subtitle="Not Found · Refund · Unidentified · Charges · Write-Off",
    )

    row_heights[row_count] = 18
    dashboard_merge(rows, merges, row_count, 2, row_count, 10, "Monthly Reconciliation Summary · Finance Operations", STYLE_DASH_FOOTER)
    dashboard_merge(rows, merges, row_count, 12, row_count, 20, "Source: Collection Summary · Disbursement Summary · Wallet Summary", STYLE_DASH_FOOTER)
    return rows, merges, row_heights


def _sheet_rows(title: str, headers: list[str], rows: list[list[Any]]) -> list[list[Any]]:
    col_count = len(headers)
    return [
        [Cell(title, STYLE_TITLE), *("" for _ in range(col_count - 1))],
        [*("" for _ in range(col_count))],
        [Cell(header, STYLE_TABLE_SECTION) for header in headers],
        *rows,
    ]


def _monthly_style(row_number: int, col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    if row_number == 3:
        return STYLE_HEADER
    if isinstance(value, Decimal):
        return STYLE_MONEY
    return None


def _run_status_rows(snapshot: Mapping[str, Any]) -> list[list[Any]]:
    return [
        list(_run_status_dict(key, area, workflow, snapshot.get(key)).values())
        for key, area, workflow in RUN_STATUS_SPECS
    ]


def _run_status_dict(key: str, area: str, workflow: str, result: Any) -> dict[str, Any]:
    if not result:
        return {
            "Area": area,
            "Workflow": workflow,
            "Status": "Not run",
            "Rows": None,
            "Amount (GHC)": "",
            "Notes": "Run this workflow before generating a full monthly summary.",
        }

    rows, amount, notes = _run_status_metrics(key, result)
    return {
        "Area": area,
        "Workflow": workflow,
        "Status": "Ready",
        "Rows": rows,
        "Amount (GHC)": _money(amount),
        "Notes": notes,
    }


def _run_status_metrics(key: str, result: Mapping[str, Any]) -> tuple[int | str, Any, str]:
    if key == "nsano_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(
                ("Nsano rows", result.get("nsano_count", 0)),
                ("Write-off rows", result.get("write_off_count", 0)),
                ("Unidentified rows", result.get("unidentified_count", 0)),
            ),
        )
    if key == "itc_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(
                ("ITC rows", result.get("itc_count", 0)),
                ("Vodafone rows", result.get("vodafone_count", 0)),
                ("Write-off rows", result.get("write_off_count", 0)),
                ("Unidentified rows", result.get("unidentified_count", 0)),
            ),
        )
    if key == "zenith_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(
                ("Zenith rows", result.get("zenith_count", 0)),
                ("Write-off rows", result.get("write_off_count", 0)),
                ("Unidentified rows", result.get("unidentified_count", 0)),
            ),
        )
    if key == "write_off_recon_result":
        stats = _stats(result, "stats")
        return (
            int(result.get("write_off_count", stats.get("total", 0))),
            stats.get("source_amount", ""),
            _notes(
                ("Nsano rows", result.get("nsano_count", 0)),
                ("ITC rows", result.get("itc_count", 0)),
                ("Zenith rows", result.get("zenith_count", 0)),
                ("Vodafone rows", result.get("vodafone_count", 0)),
            ),
        )
    if key == "nsano_disb_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(("Nsano disb rows", result.get("nsano_count", 0))),
        )
    if key == "itc_disb_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(("ITC wallet rows", result.get("itc_count", 0))),
        )
    if key == "mtn_manual_disb_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(
                ("MTN rows", result.get("mtn_count", 0)),
                ("Refund rows", result.get("refund_count", 0)),
            ),
        )
    if key == "vodafone_manual_disb_result":
        return (
            int(result.get("mambu_count", 0)),
            result.get("mambu_amount", ""),
            _notes(
                ("Vodafone rows", result.get("vodafone_count", 0)),
                ("Refund rows", result.get("refund_count", 0)),
            ),
        )

    metrics = result.get("metrics", {})
    if key == "nsano_collection_ledger_result":
        rows = (
            int(metrics.get("mambu_collection_total", 0))
            + int(metrics.get("mambu_disb_total", 0))
            + int(metrics.get("transfer_total", 0))
            + int(metrics.get("nsano_charge_total", 0))
        )
        return (
            rows,
            metrics.get("wallet_statement_balance", ""),
            _notes(
                ("Mambu collection rows", metrics.get("mambu_collection_total", 0)),
                ("Mambu disb rows", metrics.get("mambu_disb_total", 0)),
                ("Transfer rows", metrics.get("transfer_total", 0)),
                ("Nsano charge rows", metrics.get("nsano_charge_total", 0)),
            ),
        )
    if key == "nsano_wallet_ledger_result":
        rows = int(metrics.get("mambu_total", 0)) + int(metrics.get("transfer_total", 0))
        return (
            rows,
            metrics.get("wallet_statement_balance", ""),
            _notes(
                ("Mambu rows", metrics.get("mambu_total", 0)),
                ("Nsano disb rows", metrics.get("nsano_disb_total", 0)),
                ("Transfer rows", metrics.get("transfer_total", 0)),
            ),
        )
    if key == "itc_wallet_ledger_result":
        rows = int(metrics.get("statement_total", 0)) + int(metrics.get("debit_total", 0)) + int(metrics.get("credit_total", 0))
        return (
            rows,
            metrics.get("wallet_statement_balance", ""),
            _notes(
                ("Statement rows", metrics.get("statement_total", 0)),
                ("Debit rows", metrics.get("debit_total", 0)),
                ("Credit rows", metrics.get("credit_total", 0)),
            ),
        )
    return "", "", ""


def _granular_sheet_rows(snapshot: Mapping[str, Any]) -> tuple[list[list[Any]], list[str]]:
    rows: list[list[Any]] = []
    merges: list[str] = []

    def add_workflow_title(text: str) -> None:
        r = len(rows) + 1
        merges.append(f"A{r}:C{r}")
        rows.append([Cell(text, STYLE_GRAN_WORKFLOW), "", ""])

    def add_section(name: str) -> None:
        rows.append([
            Cell(name, STYLE_GRAN_SECTION),
            Cell("Count", STYLE_GRAN_SECTION_COL),
            Cell("Amount (GHC)", STYLE_GRAN_SECTION_COL),
        ])

    def add_row(label: str, count: int | str, amount: Any) -> None:
        amount_val = _money(amount)
        rows.append([
            Cell(label, STYLE_GRAN_DATA),
            Cell(count, STYLE_GRAN_COUNT) if isinstance(count, int) else Cell("", STYLE_GRAN_DATA),
            Cell(amount_val, STYLE_MONEY) if isinstance(amount_val, Decimal) else Cell("", STYLE_GRAN_DATA),
        ])

    def add_total(count: int | str, amount: Any) -> None:
        amount_val = _money(amount)
        rows.append([
            Cell("TOTAL", STYLE_GRAN_TOTAL),
            Cell(count, STYLE_GRAN_TOTAL_NUM) if isinstance(count, int) else Cell("", STYLE_GRAN_TOTAL),
            Cell(amount_val, STYLE_GRAN_TOTAL_MONEY) if isinstance(amount_val, Decimal) else Cell("", STYLE_GRAN_TOTAL),
        ])

    def add_balance(label: str, amount: Any) -> None:
        amount_val = _money(amount)
        rows.append([
            Cell(label, STYLE_GRAN_TOTAL),
            Cell("", STYLE_GRAN_TOTAL),
            Cell(amount_val, STYLE_GRAN_TOTAL_MONEY) if isinstance(amount_val, Decimal) else Cell("", STYLE_GRAN_TOTAL),
        ])

    def add_blank() -> None:
        rows.append(["", "", ""])

    def sc(stats: Mapping[str, Any], key: str) -> int:
        return int(stats.get("status_counts", {}).get(key, 0))

    def sa(stats: Mapping[str, Any], key: str) -> Any:
        return stats.get("status_amounts", {}).get(key, "")

    # ── Main title ───────────────────────────────────────────────────────
    add_workflow_title("MONTHLY RECONCILIATION — GRANULAR DETAIL")
    add_blank()

    # ── COLLECTION ───────────────────────────────────────────────────────
    has_collection = any(k in snapshot for k in (
        "nsano_result", "itc_result", "zenith_result", "write_off_recon_result",
    ))
    if has_collection:
        add_workflow_title("COLLECTION WORKFLOWS")
        add_blank()

    if "nsano_result" in snapshot:
        r = snapshot["nsano_result"]
        mvn = r.get("mvn", {})
        nvm = r.get("nvm", {})
        add_workflow_title("NSANO COLLECTION RECONCILIATION")
        add_section("Mambu vs Nsano")
        add_row("Nsano (Matched)", int(mvn.get("matched", 0)), mvn.get("matched_amount", ""))
        add_row("Not Found", int(mvn.get("not_found", 0)), mvn.get("not_found_amount", ""))
        add_total(int(mvn.get("total", 0)), mvn.get("source_amount", ""))
        add_blank()
        add_section("Nsano vs Mambu")
        add_row("Mambu", int(nvm.get("mambu", 0)), nvm.get("mambu_amount", ""))
        add_row("Unidentified", int(nvm.get("unidentified", 0)), nvm.get("unidentified_amount", ""))
        add_row("Write Off", int(nvm.get("write_off", 0)), nvm.get("write_off_amount", ""))
        add_row("Not Found", int(nvm.get("not_found", 0)), nvm.get("not_found_amount", ""))
        add_total(int(nvm.get("total", 0)), nvm.get("source_amount", ""))
        add_blank()

    if "itc_result" in snapshot:
        r = snapshot["itc_result"]
        mvi = r.get("mvi", {})
        ivm = r.get("ivm", {})
        vvm = r.get("vvm", {})
        add_workflow_title("ITC / VODAFONE COLLECTION RECONCILIATION")
        add_section("Mambu vs ITC / Voda")
        add_row(ITC_STATUS, sc(mvi, ITC_STATUS), sa(mvi, ITC_STATUS))
        add_row(VODA_COLL_STATUS, sc(mvi, VODA_COLL_STATUS), sa(mvi, VODA_COLL_STATUS))
        add_row(ITC_NOT_FOUND_STATUS, sc(mvi, ITC_NOT_FOUND_STATUS), sa(mvi, ITC_NOT_FOUND_STATUS))
        add_total(int(mvi.get("total", 0)), mvi.get("source_amount", ""))
        add_blank()
        add_section("ITC vs Mambu")
        add_row(ITC_MAMBU_STATUS, sc(ivm, ITC_MAMBU_STATUS), sa(ivm, ITC_MAMBU_STATUS))
        add_row(VODA_COLL_STATUS, sc(ivm, VODA_COLL_STATUS), sa(ivm, VODA_COLL_STATUS))
        add_row(ITC_UNIDENTIFIED_STATUS, sc(ivm, ITC_UNIDENTIFIED_STATUS), sa(ivm, ITC_UNIDENTIFIED_STATUS))
        add_row(ITC_WRITE_OFF_STATUS, sc(ivm, ITC_WRITE_OFF_STATUS), sa(ivm, ITC_WRITE_OFF_STATUS))
        add_row(ITC_UPSALE_STATUS, sc(ivm, ITC_UPSALE_STATUS), sa(ivm, ITC_UPSALE_STATUS))
        add_row(ITC_NOT_FOUND_STATUS, sc(ivm, ITC_NOT_FOUND_STATUS), sa(ivm, ITC_NOT_FOUND_STATUS))
        add_total(int(ivm.get("total", 0)), ivm.get("source_amount", ""))
        add_blank()
        add_section("Vodafone Collections vs Mambu")
        add_row(ITC_MAMBU_STATUS, sc(vvm, ITC_MAMBU_STATUS), sa(vvm, ITC_MAMBU_STATUS))
        add_row(ITC_UNIDENTIFIED_STATUS, sc(vvm, ITC_UNIDENTIFIED_STATUS), sa(vvm, ITC_UNIDENTIFIED_STATUS))
        add_row(VODAFONE_WRITE_OFF_STATUS, sc(vvm, VODAFONE_WRITE_OFF_STATUS), sa(vvm, VODAFONE_WRITE_OFF_STATUS))
        add_row(ITC_NOT_FOUND_STATUS, sc(vvm, ITC_NOT_FOUND_STATUS), sa(vvm, ITC_NOT_FOUND_STATUS))
        add_total(int(vvm.get("total", 0)), vvm.get("source_amount", ""))
        add_blank()

    if "zenith_result" in snapshot:
        r = snapshot["zenith_result"]
        mvz = r.get("mvz", {})
        zvs = r.get("zvs", {})
        add_workflow_title("ZENITH COLLECTION RECONCILIATION")
        add_section("Mambu vs Zenith Bank")
        add_row(ZENITH_MATCHED_STATUS, sc(mvz, ZENITH_MATCHED_STATUS), sa(mvz, ZENITH_MATCHED_STATUS))
        add_row(ZENITH_NOT_FOUND_STATUS, sc(mvz, ZENITH_NOT_FOUND_STATUS), sa(mvz, ZENITH_NOT_FOUND_STATUS))
        add_total(int(mvz.get("total", 0)), mvz.get("source_amount", ""))
        add_blank()
        add_section("Zenith Bank as Source")
        add_row(ZENITH_MAMBU_STATUS, sc(zvs, ZENITH_MAMBU_STATUS), sa(zvs, ZENITH_MAMBU_STATUS))
        add_row(ZENITH_UNIDENTIFIED_STATUS, sc(zvs, ZENITH_UNIDENTIFIED_STATUS), sa(zvs, ZENITH_UNIDENTIFIED_STATUS))
        add_row(ZENITH_WRITE_OFF_STATUS, sc(zvs, ZENITH_WRITE_OFF_STATUS), sa(zvs, ZENITH_WRITE_OFF_STATUS))
        add_row(ZENITH_NOT_FOUND_STATUS, sc(zvs, ZENITH_NOT_FOUND_STATUS), sa(zvs, ZENITH_NOT_FOUND_STATUS))
        add_total(int(zvs.get("total", 0)), zvs.get("source_amount", ""))
        add_blank()

    if "write_off_recon_result" in snapshot:
        r = snapshot["write_off_recon_result"]
        st = _stats(r, "stats")
        add_workflow_title("WRITE-OFF RECONCILIATION")
        add_section("Write Off vs Collection Sources")
        add_row(WRITE_OFF_NSANO_STATUS, sc(st, WRITE_OFF_NSANO_STATUS), sa(st, WRITE_OFF_NSANO_STATUS))
        add_row(WRITE_OFF_ITC_STATUS, sc(st, WRITE_OFF_ITC_STATUS), sa(st, WRITE_OFF_ITC_STATUS))
        add_row(WRITE_OFF_ZENITH_STATUS, sc(st, WRITE_OFF_ZENITH_STATUS), sa(st, WRITE_OFF_ZENITH_STATUS))
        add_row(WRITE_OFF_VODAFONE_STATUS, sc(st, WRITE_OFF_VODAFONE_STATUS), sa(st, WRITE_OFF_VODAFONE_STATUS))
        add_row(WRITE_OFF_NOT_FOUND_STATUS, sc(st, WRITE_OFF_NOT_FOUND_STATUS), sa(st, WRITE_OFF_NOT_FOUND_STATUS))
        add_total(int(st.get("total", 0)), st.get("source_amount", ""))
        add_blank()

    # ── DISBURSEMENT ─────────────────────────────────────────────────────
    has_disb = any(k in snapshot for k in (
        "nsano_disb_result", "itc_disb_result", "mtn_manual_disb_result", "vodafone_manual_disb_result",
    ))
    if has_disb:
        add_workflow_title("DISBURSEMENT WORKFLOWS")
        add_blank()

    if "nsano_disb_result" in snapshot:
        r = snapshot["nsano_disb_result"]
        mvn = r.get("mvn", {})
        nvm = r.get("nvm", {})
        add_workflow_title("NSANO DISBURSEMENT RECONCILIATION")
        add_section("Mambu vs Nsano")
        add_row(DISB_NSANO_STATUS, sc(mvn, DISB_NSANO_STATUS), sa(mvn, DISB_NSANO_STATUS))
        add_row(DISB_NOT_FOUND_STATUS, sc(mvn, DISB_NOT_FOUND_STATUS), sa(mvn, DISB_NOT_FOUND_STATUS))
        add_total(int(mvn.get("total", 0)), mvn.get("source_amount", ""))
        add_blank()
        add_section("Nsano vs Mambu")
        add_row(DISB_MAMBU_STATUS, sc(nvm, DISB_MAMBU_STATUS), sa(nvm, DISB_MAMBU_STATUS))
        add_row(DISB_NOT_FOUND_STATUS, sc(nvm, DISB_NOT_FOUND_STATUS), sa(nvm, DISB_NOT_FOUND_STATUS))
        add_total(int(nvm.get("total", 0)), nvm.get("source_amount", ""))
        add_blank()

    if "itc_disb_result" in snapshot:
        r = snapshot["itc_disb_result"]
        mvi = r.get("mvi", {})
        ivm = r.get("ivm", {})
        add_workflow_title("ITC WALLET vs MAMBU RECONCILIATION")
        add_section("Mambu vs ITC Wallet")
        add_row(ITC_DISB_STATUS, sc(mvi, ITC_DISB_STATUS), sa(mvi, ITC_DISB_STATUS))
        add_row(ITC_DISB_NOT_FOUND_STATUS, sc(mvi, ITC_DISB_NOT_FOUND_STATUS), sa(mvi, ITC_DISB_NOT_FOUND_STATUS))
        add_total(int(mvi.get("total", 0)), mvi.get("source_amount", ""))
        add_blank()
        add_section("ITC Wallet vs Mambu")
        add_row(DISB_MAMBU_STATUS, sc(ivm, DISB_MAMBU_STATUS), sa(ivm, DISB_MAMBU_STATUS))
        add_row(ITC_DISB_REFERRAL_BONUS_STATUS, sc(ivm, ITC_DISB_REFERRAL_BONUS_STATUS), sa(ivm, ITC_DISB_REFERRAL_BONUS_STATUS))
        add_row(ITC_DISB_UPSALES_REFOUND_STATUS, sc(ivm, ITC_DISB_UPSALES_REFOUND_STATUS), sa(ivm, ITC_DISB_UPSALES_REFOUND_STATUS))
        add_row(ITC_DISB_SAVINGS_REWARD_STATUS, sc(ivm, ITC_DISB_SAVINGS_REWARD_STATUS), sa(ivm, ITC_DISB_SAVINGS_REWARD_STATUS))
        add_row(ITC_DISB_NOT_FOUND_STATUS, sc(ivm, ITC_DISB_NOT_FOUND_STATUS), sa(ivm, ITC_DISB_NOT_FOUND_STATUS))
        add_total(int(ivm.get("total", 0)), ivm.get("source_amount", ""))
        add_blank()

    if "mtn_manual_disb_result" in snapshot:
        r = snapshot["mtn_manual_disb_result"]
        mvm = r.get("mvm", {})
        mtv = r.get("mtv", {})
        add_workflow_title("MTN MANUAL DISBURSEMENT RECONCILIATION")
        add_section("Mambu vs MTN MANUAL")
        add_row(MTN_MANUAL_MATCHED_STATUS, sc(mvm, MTN_MANUAL_MATCHED_STATUS), sa(mvm, MTN_MANUAL_MATCHED_STATUS))
        add_row(MTN_MANUAL_NOT_FOUND_STATUS, sc(mvm, MTN_MANUAL_NOT_FOUND_STATUS), sa(mvm, MTN_MANUAL_NOT_FOUND_STATUS))
        add_total(int(mvm.get("total", 0)), mvm.get("source_amount", ""))
        add_blank()
        add_section("MTN MANUAL as Source")
        add_row(MTN_MANUAL_MAMBU_STATUS, sc(mtv, MTN_MANUAL_MAMBU_STATUS), sa(mtv, MTN_MANUAL_MAMBU_STATUS))
        add_row(MTN_MANUAL_REFUND_STATUS, sc(mtv, MTN_MANUAL_REFUND_STATUS), sa(mtv, MTN_MANUAL_REFUND_STATUS))
        add_row(MTN_MANUAL_NOT_FOUND_STATUS, sc(mtv, MTN_MANUAL_NOT_FOUND_STATUS), sa(mtv, MTN_MANUAL_NOT_FOUND_STATUS))
        add_total(int(mtv.get("total", 0)), mtv.get("source_amount", ""))
        add_blank()

    if "vodafone_manual_disb_result" in snapshot:
        r = snapshot["vodafone_manual_disb_result"]
        mvv = r.get("mvv", {})
        vtv = r.get("vtv", {})
        add_workflow_title("VODAFONE MANUAL DISBURSEMENT RECONCILIATION")
        add_section("Mambu vs VODAFONE MANUAL")
        add_row(VODAFONE_MANUAL_MATCHED_STATUS, sc(mvv, VODAFONE_MANUAL_MATCHED_STATUS), sa(mvv, VODAFONE_MANUAL_MATCHED_STATUS))
        add_row(VODAFONE_MANUAL_NOT_FOUND_STATUS, sc(mvv, VODAFONE_MANUAL_NOT_FOUND_STATUS), sa(mvv, VODAFONE_MANUAL_NOT_FOUND_STATUS))
        add_total(int(mvv.get("total", 0)), mvv.get("source_amount", ""))
        add_blank()
        add_section("VODAFONE MANUAL as Source")
        add_row(VODAFONE_MANUAL_MAMBU_STATUS, sc(vtv, VODAFONE_MANUAL_MAMBU_STATUS), sa(vtv, VODAFONE_MANUAL_MAMBU_STATUS))
        add_row(VODAFONE_MANUAL_REFUND_STATUS, sc(vtv, VODAFONE_MANUAL_REFUND_STATUS), sa(vtv, VODAFONE_MANUAL_REFUND_STATUS))
        add_row(VODAFONE_MANUAL_NOT_FOUND_STATUS, sc(vtv, VODAFONE_MANUAL_NOT_FOUND_STATUS), sa(vtv, VODAFONE_MANUAL_NOT_FOUND_STATUS))
        add_total(int(vtv.get("total", 0)), vtv.get("source_amount", ""))
        add_blank()

    # ── WALLET LEDGER ─────────────────────────────────────────────────────
    has_wallet = any(k in snapshot for k in ("nsano_collection_ledger_result", "nsano_wallet_ledger_result", "itc_wallet_ledger_result"))
    if has_wallet:
        add_workflow_title("WALLET LEDGER WORKFLOWS")
        add_blank()

    if "nsano_collection_ledger_result" in snapshot:
        r = snapshot["nsano_collection_ledger_result"]
        m = r.get("metrics", {})
        cat_counts = m.get("transfer_category_counts", {})
        cat_amounts = m.get("transfer_category_amounts", {})
        add_workflow_title("NSANO COLLECTIONS vs LEDGER")
        add_section("Credit / Available Funds")
        add_row("Total Nsano Mambu Collections", int(m.get("mambu_collection_total", 0)), m.get("topup_collections", ""))
        add_row("Bank to Wallet", int(cat_counts.get(NSANO_WALLET_BANK_TO_WALLET_STATUS, 0)), m.get("bank_to_wallet", ""))
        add_row("Reversal Adjustment", int(cat_counts.get(NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS, 0)), m.get("reversal_adjustment", ""))
        add_row("Other Transfers", int(cat_counts.get(NSANO_WALLET_OTHER_STATUS, 0)), cat_amounts.get(NSANO_WALLET_OTHER_STATUS, ""))
        add_row("Manual: Ledger Balance", "", m.get("ledger_balance", ""))
        add_row("Manual: Delayed Transactions", "", m.get("manual_unidentified", ""))
        add_row("Manual: Recovery from write off", "", m.get("settlement", ""))
        add_row("Manual: Reversal", "", m.get("reversal", ""))
        add_balance("AVAILABLE FUNDS BEFORE DEBIT", m.get("available_funds", ""))
        add_blank()
        add_section("Debit")
        add_row("Transfer to Bank", int(cat_counts.get(NSANO_WALLET_TRANSFER_TO_BANK_STATUS, 0)), m.get("transfer_to_bank", ""))
        add_row("Disbursement (Mambu)", int(m.get("mambu_disb_total", 0)), m.get("disbursement", ""))
        add_row("Charges (Nsano W2A)", int(m.get("nsano_charge_total", 0)), m.get("charges", ""))
        add_row("Delayed Transactions Debit", "", m.get("unidentified_debit", ""))
        add_balance("TOTAL DEBIT", m.get("total_debit", ""))
        add_blank()
        add_section("Balance")
        add_balance("WALLET STATEMENT BALANCE", m.get("wallet_statement_balance", ""))
        add_blank()

    if "nsano_wallet_ledger_result" in snapshot:
        r = snapshot["nsano_wallet_ledger_result"]
        m = r.get("metrics", {})
        cat_counts = m.get("transfer_category_counts", {})
        cat_amounts = m.get("transfer_category_amounts", {})
        add_workflow_title("NSANO DISB WALLET vs LEDGER")
        add_section("Credit / Available Funds")
        add_row("Top-up Collections", int(cat_counts.get(NSANO_WALLET_TOPUP_COLLECTION_STATUS, 0)), m.get("topup_collections", ""))
        add_row("Bank to Wallet", int(cat_counts.get(NSANO_WALLET_BANK_TO_WALLET_STATUS, 0)), m.get("bank_to_wallet", ""))
        add_row("Reversal Adjustment", int(cat_counts.get(NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS, 0)), m.get("reversal_adjustment", ""))
        add_row("Other Transfers", int(cat_counts.get(NSANO_WALLET_OTHER_STATUS, 0)), cat_amounts.get(NSANO_WALLET_OTHER_STATUS, ""))
        add_row("Manual: Ledger Balance", "", m.get("ledger_balance", ""))
        add_row("Manual: Delayed Transactions", "", m.get("manual_unidentified", ""))
        add_row("Manual: Recovery from write off", "", m.get("settlement", ""))
        add_row("Manual: Reversal", "", m.get("reversal", ""))
        add_balance("AVAILABLE FUNDS BEFORE DEBIT", m.get("available_funds", ""))
        add_blank()
        add_section("Debit")
        add_row("Transfer to Bank", int(cat_counts.get(NSANO_WALLET_TRANSFER_TO_BANK_STATUS, 0)), m.get("transfer_to_bank", ""))
        add_row("Disbursement (Mambu)", int(m.get("mambu_total", 0)), m.get("disbursement", ""))
        add_row("Delayed Transactions Debit", "", m.get("unidentified_debit", ""))
        add_balance("TOTAL DEBIT", m.get("total_debit", ""))
        add_blank()
        add_section("Balance")
        add_balance("WALLET STATEMENT BALANCE", m.get("wallet_statement_balance", ""))
        add_blank()

    if "itc_wallet_ledger_result" in snapshot:
        r = snapshot["itc_wallet_ledger_result"]
        m = r.get("metrics", {})
        debit_counts = m.get("debit_status_counts", {})
        credit_counts = m.get("credit_category_counts", {})
        debit_amounts = m.get("debit_status_amounts", {})
        credit_amounts = m.get("credit_category_amounts", {})
        add_workflow_title("ITC WALLET vs LEDGER")
        add_section("Credit / Available Funds")
        add_row("Settlement", int(credit_counts.get(ITC_WALLET_SETTLEMENT_STATUS, 0)), m.get("settlement", ""))
        add_row("Reversal", int(credit_counts.get(ITC_WALLET_REVERSAL_STATUS, 0)), m.get("reversal", ""))
        add_row("Transfer to Wallet", int(credit_counts.get(ITC_WALLET_TRANSFER_TO_WALLET_STATUS, 0)), m.get("transfer_to_wallet", ""))
        add_row("Prepaid Reversal", int(credit_counts.get(ITC_WALLET_PREPAID_REVERSAL_STATUS, 0)), credit_amounts.get(ITC_WALLET_PREPAID_REVERSAL_STATUS, ""))
        add_row("Manual: Ledger Balance", "", m.get("ledger_balance", ""))
        add_row("Manual: Unidentified", "", m.get("manual_unidentified", ""))
        add_row("Unidentified Credit", "", m.get("unidentified_credit", ""))
        add_balance("AVAILABLE FUNDS BEFORE DEBIT", m.get("available_funds", ""))
        add_blank()
        add_section("Debit")
        add_row("Transfer to Bank", int(debit_counts.get(ITC_WALLET_TRANSFERS_FROM_BANK_STATUS, 0)), m.get("transfer_to_bank", ""))
        add_row("Referral Awards", int(debit_counts.get(ITC_WALLET_REFERRAL_AWARD_STATUS, 0)), m.get("referral_awards", ""))
        add_row("Savings", int(debit_counts.get(ITC_WALLET_SAVINGS_STATUS, 0)), m.get("savings", ""))
        add_row("Upsales Refund", int(debit_counts.get(ITC_WALLET_UPSALES_REFUND_STATUS, 0)), m.get("upsales_refund", ""))
        add_row("Disbursement", int(debit_counts.get(ITC_WALLET_DISBURSEMENT_STATUS, 0)), m.get("disbursement", ""))
        add_row("Credit Transfer Fallback", int(debit_counts.get(ITC_WALLET_CREDIT_TRANSFER_STATUS, 0)), debit_amounts.get(ITC_WALLET_CREDIT_TRANSFER_STATUS, ""))
        add_row("Not Found", int(debit_counts.get(ITC_WALLET_NOT_FOUND_STATUS, 0)), debit_amounts.get(ITC_WALLET_NOT_FOUND_STATUS, ""))
        add_row("Unidentified Debit", "", m.get("unidentified_debit", ""))
        add_balance("TOTAL DEBIT", m.get("total_debit", ""))
        add_blank()
        add_section("Balance")
        add_balance("WALLET STATEMENT BALANCE", m.get("wallet_statement_balance", ""))
        add_blank()

    return rows, merges


def _collection_summary_rows(snapshot: Mapping[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    if "nsano_result" in snapshot:
        result = snapshot["nsano_result"]
        rows.append(_recon_row("Nsano Coll Recon", "Mambu vs Nsano", _stats(result, "mvn")))
        rows.append(_recon_row("Nsano Coll Recon", "Nsano vs Mambu", _stats(result, "nvm")))
    if "itc_result" in snapshot:
        result = snapshot["itc_result"]
        rows.append(_recon_row("ITC/Voda Coll Recon", "Mambu vs ITC", _stats(result, "mvi")))
        rows.append(_recon_row("ITC/Voda Coll Recon", "ITC vs Mambu", _stats(result, "ivm")))
        rows.append(_recon_row("ITC/Voda Coll Recon", "Vodafone Collections vs Mambu", _stats(result, "vvm")))
    if "zenith_result" in snapshot:
        result = snapshot["zenith_result"]
        rows.append(_recon_row("Zenith Coll Recon", "Mambu vs Zenith", _stats(result, "mvz")))
        rows.append(_recon_row("Zenith Coll Recon", "Zenith as Source", _stats(result, "zvs")))
    if "write_off_recon_result" in snapshot:
        result = snapshot["write_off_recon_result"]
        stats = _stats(result, "stats")
        notes = _status_notes(
            stats,
            WRITE_OFF_NSANO_STATUS,
            WRITE_OFF_ITC_STATUS,
            WRITE_OFF_ZENITH_STATUS,
            WRITE_OFF_VODAFONE_STATUS,
            WRITE_OFF_NOT_FOUND_STATUS,
        )
        rows.append(_recon_row("Write-off Recon", "Write Off vs Collection Sources", stats, notes))
    return rows or [_empty_recon_row("No collection workflows have been run yet.")]


def _disbursement_summary_rows(snapshot: Mapping[str, Any]) -> list[list[Any]]:
    # Design: produce two compact subtables displayed within the same 10-column sheet:
    #  - "Disbursement Recon - Wallet/Manual as Source": for each external wallet/manual source show how many rows were found in Mambu
    #  - "Disbursement Recon - Mambu as Source": for Mambu as the source show counts found in each disbursement target
    # We still return 10-column rows to keep the sheet layout compatible with the existing workbook writer.
    rows: list[list[Any]] = []

    def section_row(title: str) -> None:
        # place a visible section title in the Workflow column (col 0)
        rows.append([Cell(title, STYLE_TABLE_SECTION), "", "", "", "", "", "", "", "", ""])

    def provider_row(source: str, found_in: str, count: int | str, amount: Any, notes: str = "") -> None:
        # fill the same 10-column shape: workflow, comparison, total, ..., source_amount, ..., notes
        rows.append([
            source,
            found_in,
            int(count or 0),
            "",
            "",
            "",
            _money(amount),
            "",
            "",
            notes,
        ])

    # Wallet/Manual as Source table
    added_any = False
    section_row("Disbursement Recon - Wallet/Manual as Source")
    if "nsano_disb_result" in snapshot:
        added_any = True
        r = snapshot["nsano_disb_result"]
        mvn = _stats(r, "mvn")
        provider_row("Nsano", "Found In Mambu (Nsano Disbursement)", mvn.get("matched", 0), mvn.get("matched_amount", mvn.get("source_amount", "")))
    if "itc_disb_result" in snapshot:
        added_any = True
        r = snapshot["itc_disb_result"]
        mvi = _stats(r, "mvi")
        provider_row("ITC", "Found In Mambu (ITC Disbursement)", mvi.get("matched", mvi.get("total", 0)), mvi.get("matched_amount", mvi.get("source_amount", "")))
    if "mtn_manual_disb_result" in snapshot:
        added_any = True
        r = snapshot["mtn_manual_disb_result"]
        mvm = _stats(r, "mvm")
        provider_row("MTN Manual", "Found In Mambu (MTN Manual Disbursement)", mvm.get("matched", mvm.get("total", 0)), mvm.get("matched_amount", mvm.get("source_amount", "")))
    if "vodafone_manual_disb_result" in snapshot:
        added_any = True
        r = snapshot["vodafone_manual_disb_result"]
        mvv = _stats(r, "mvv")
        provider_row("Vodafone Manual", "Found In Mambu (Vodafone Manual Disbursement)", mvv.get("matched", mvv.get("total", 0)), mvv.get("matched_amount", mvv.get("source_amount", "")))

    if not added_any:
        rows.append(["", "", "", "", "", "", "", "", "", "No disbursement workflows have been run yet."])

    # Blank spacer
    rows.append(["", "", "", "", "", "", "", "", "", ""])

    # Mambu as Source table
    section_row("Disbursement Recon - Mambu as Source")
    added_mambu = False
    # For each disbursement workflow, use the 'other-side' stats where applicable
    if "nsano_disb_result" in snapshot:
        added_mambu = True
        r = snapshot["nsano_disb_result"]
        nvm = _stats(r, "nvm")
        # nvm typically contains how Nsano rows map into Mambu; 'mambu' key counts Mambu matches
        provider_row("Mambu", "Found In Nsano Disbursement", nvm.get("mambu", nvm.get("matched", 0)), nvm.get("mambu_amount", nvm.get("matched_amount", nvm.get("source_amount", ""))))
    if "itc_disb_result" in snapshot:
        added_mambu = True
        r = snapshot["itc_disb_result"]
        ivm = _stats(r, "ivm")
        provider_row("Mambu", "Found In ITC Disbursement", ivm.get("mambu", ivm.get("matched", 0)), ivm.get("mambu_amount", ivm.get("matched_amount", ivm.get("source_amount", ""))))
    if "mtn_manual_disb_result" in snapshot:
        added_mambu = True
        r = snapshot["mtn_manual_disb_result"]
        mtv = _stats(r, "mtv")
        provider_row("Mambu", "Found In MTN Manual Disbursement", mtv.get("mambu", mtv.get("matched", 0)), mtv.get("mambu_amount", mtv.get("matched_amount", mtv.get("source_amount", ""))))
    if "vodafone_manual_disb_result" in snapshot:
        added_mambu = True
        r = snapshot["vodafone_manual_disb_result"]
        vtv = _stats(r, "vtv")
        provider_row("Mambu", "Found In Vodafone Manual Disbursement", vtv.get("mambu", vtv.get("matched", 0)), vtv.get("mambu_amount", vtv.get("matched_amount", vtv.get("source_amount", ""))))

    # Analysis row: compare wallet-found totals vs mambu-found totals
    try:
        wallet_total = sum(int(row[2] or 0) for row in rows if row[0] in ("Nsano", "ITC", "MTN Manual", "Vodafone Manual"))
        mambu_total = sum(int(row[2] or 0) for row in rows if row[0] == "Mambu")
        wallet_amount = sum((row[6] or 0) for row in rows if isinstance(row[6], Decimal))
        mambu_amount = sum((row[6] or 0) for row in rows if row[0] == "Mambu" and isinstance(row[6], Decimal))
        rows.append(["", "", "", "", "", "", "", "", "", ""])
        rows.append([Cell("Analysis", STYLE_TABLE_SECTION), "", "", "", "", "", "", "", "", ""])
        rows.append(["Wallet - Mambu", "", wallet_total - mambu_total, "", "", "", _money(wallet_amount - mambu_amount), "", "", ""])
    except Exception:
        # keep graceful fallback if any values missing
        pass

    return rows


def _wallet_summary_rows(snapshot: Mapping[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    if "nsano_collection_ledger_result" in snapshot:
        rows.extend(_nsano_collection_ledger_rows(snapshot["nsano_collection_ledger_result"].get("metrics", {})))
    if "nsano_wallet_ledger_result" in snapshot:
        rows.extend(_nsano_wallet_rows(snapshot["nsano_wallet_ledger_result"].get("metrics", {})))
    if "itc_wallet_ledger_result" in snapshot:
        rows.extend(_itc_wallet_rows(snapshot["itc_wallet_ledger_result"].get("metrics", {})))
    return rows or [["", "", "No wallet ledger workflows have been run yet.", "", "", ""]]


def _recon_row(workflow: str, comparison: str, stats: Mapping[str, Any], notes: str = "") -> list[Any]:
    return [
        workflow,
        comparison,
        int(stats.get("total", 0)),
        int(stats.get("matched", 0)),
        int(stats.get("not_found", 0)),
        Cell(float(stats.get("match_rate", 0) or 0), STYLE_PERCENT),
        _money(stats.get("source_amount", "")),
        _money(stats.get("matched_amount", "")),
        _money(stats.get("not_found_amount", "")),
        notes,
    ]


def _empty_recon_row(note: str) -> list[Any]:
    return ["", "", "", "", "", "", "", "", "", note]


def _nsano_wallet_rows(metrics: Mapping[str, Any]) -> list[list[Any]]:
    counts = metrics.get("transfer_category_counts", {})
    category_amounts = metrics.get("transfer_category_amounts", {})
    return [
        _wallet_row("Nsano Disb", "Credit / Available Funds", "Top-up Collections", _count(counts, NSANO_WALLET_TOPUP_COLLECTION_STATUS), metrics.get("topup_collections", "")),
        _wallet_row("Nsano Disb", "Credit / Available Funds", "Bank to Wallet", _count(counts, NSANO_WALLET_BANK_TO_WALLET_STATUS), metrics.get("bank_to_wallet", "")),
        _wallet_row("Nsano Disb", "Credit / Available Funds", "Reversal Adjustment", _count(counts, NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS), metrics.get("reversal_adjustment", "")),
        _wallet_row("Nsano Disb", "Credit / Available Funds", "Other Transfer Rows", _count(counts, NSANO_WALLET_OTHER_STATUS), _status_money(category_amounts, NSANO_WALLET_OTHER_STATUS)),
        _wallet_row("Nsano Disb", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _wallet_row("Nsano Disb", "Manual Ledger Inputs", "Manual Delayed Transactions", "", metrics.get("manual_unidentified", "")),
        _wallet_row("Nsano Disb", "Manual Ledger Inputs", "Recovery from write off", "", metrics.get("settlement", "")),
        _wallet_row("Nsano Disb", "Manual Ledger Inputs", "Reversal", "", metrics.get("reversal", "")),
        _wallet_row("Nsano Disb", "Debit", "Transfer to Bank", _count(counts, NSANO_WALLET_TRANSFER_TO_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _wallet_row("Nsano Disb", "Debit", "Disbursement", int(metrics.get("mambu_total", 0)), metrics.get("disbursement", ""), "Mambu disbursement rows."),
        _wallet_row("Nsano Disb", "Debit", "Delayed Transactions", "", metrics.get("unidentified_debit", ""), "Negative manual delayed transactions, shown as a positive debit."),
        _wallet_row("Nsano Disb", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _wallet_row("Nsano Disb", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _nsano_collection_ledger_rows(metrics: Mapping[str, Any]) -> list[list[Any]]:
    counts = metrics.get("transfer_category_counts", {})
    category_amounts = metrics.get("transfer_category_amounts", {})
    return [
        _wallet_row("Nsano Collections", "Credit / Available Funds", "Total Nsano Mambu Collections", int(metrics.get("mambu_collection_total", 0)), metrics.get("topup_collections", "")),
        _wallet_row("Nsano Collections", "Credit / Available Funds", "Bank to Wallet", _count(counts, NSANO_WALLET_BANK_TO_WALLET_STATUS), metrics.get("bank_to_wallet", "")),
        _wallet_row("Nsano Collections", "Credit / Available Funds", "Reversal Adjustment", _count(counts, NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS), metrics.get("reversal_adjustment", "")),
        _wallet_row("Nsano Collections", "Credit / Available Funds", "Top-up Transfer Rows", _count(counts, NSANO_WALLET_TOPUP_COLLECTION_STATUS), metrics.get("transfer_topup_collections", ""), "Retained for audit; Mambu collections feed the credit line."),
        _wallet_row("Nsano Collections", "Credit / Available Funds", "Other Transfer Rows", _count(counts, NSANO_WALLET_OTHER_STATUS), _status_money(category_amounts, NSANO_WALLET_OTHER_STATUS)),
        _wallet_row("Nsano Collections", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _wallet_row("Nsano Collections", "Manual Ledger Inputs", "Manual Delayed Transactions", "", metrics.get("manual_unidentified", "")),
        _wallet_row("Nsano Collections", "Manual Ledger Inputs", "Recovery from write off", "", metrics.get("settlement", "")),
        _wallet_row("Nsano Collections", "Manual Ledger Inputs", "Reversal", "", metrics.get("reversal", "")),
        _wallet_row("Nsano Collections", "Debit", "Transfer to Bank", _count(counts, NSANO_WALLET_TRANSFER_TO_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _wallet_row("Nsano Collections", "Debit", "Disbursement", int(metrics.get("mambu_disb_total", 0)), metrics.get("disbursement", ""), "Mambu disbursement rows from the Mambu Filtered Disbursement upload."),
        _wallet_row("Nsano Collections", "Debit", "Charges", int(metrics.get("nsano_charge_total", 0)), metrics.get("charges", ""), "Nsano Filtered Data W2A column U (Charge), optional upload."),
        _wallet_row("Nsano Collections", "Debit", "Delayed Transactions", "", metrics.get("unidentified_debit", ""), "Negative manual delayed transactions, shown as a positive debit."),
        _wallet_row("Nsano Collections", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _wallet_row("Nsano Collections", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _itc_wallet_rows(metrics: Mapping[str, Any]) -> list[list[Any]]:
    debit_counts = metrics.get("debit_status_counts", {})
    credit_counts = metrics.get("credit_category_counts", {})
    return [
        _wallet_row("ITC", "Credit / Available Funds", "Settlement", _count(credit_counts, ITC_WALLET_SETTLEMENT_STATUS), metrics.get("settlement", "")),
        _wallet_row("ITC", "Credit / Available Funds", "Reversal", _count(credit_counts, ITC_WALLET_REVERSAL_STATUS), metrics.get("reversal", "")),
        _wallet_row("ITC", "Credit / Available Funds", "Transfer to Wallet", _count(credit_counts, ITC_WALLET_TRANSFER_TO_WALLET_STATUS), metrics.get("transfer_to_wallet", "")),
        _wallet_row("ITC", "Credit / Available Funds", "Prepaid Reversal", _count(credit_counts, ITC_WALLET_PREPAID_REVERSAL_STATUS), _status_money(metrics.get("credit_category_amounts", {}), ITC_WALLET_PREPAID_REVERSAL_STATUS)),
        _wallet_row("ITC", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _wallet_row("ITC", "Manual Ledger Inputs", "Manual Unidentified", "", metrics.get("manual_unidentified", "")),
        _wallet_row("ITC", "Manual Ledger Inputs", "Unidentified Credit", "", metrics.get("unidentified_credit", "")),
        _wallet_row("ITC", "Manual Ledger Inputs", "Unidentified Debit", "", metrics.get("unidentified_debit", "")),
        _wallet_row("ITC", "Debit", "Transfer to Bank", _count(debit_counts, ITC_WALLET_TRANSFERS_FROM_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _wallet_row("ITC", "Debit", "Referral Awards", _count(debit_counts, ITC_WALLET_REFERRAL_AWARD_STATUS), metrics.get("referral_awards", "")),
        _wallet_row("ITC", "Debit", "Savings", _count(debit_counts, ITC_WALLET_SAVINGS_STATUS), metrics.get("savings", "")),
        _wallet_row("ITC", "Debit", "Upsales Refund", _count(debit_counts, ITC_WALLET_UPSALES_REFUND_STATUS), metrics.get("upsales_refund", "")),
        _wallet_row("ITC", "Debit", "Disbursement", _count(debit_counts, ITC_WALLET_DISBURSEMENT_STATUS), metrics.get("disbursement", "")),
        _wallet_row("ITC", "Debit", "Credit Transfer Fallback", _count(debit_counts, ITC_WALLET_CREDIT_TRANSFER_STATUS), _status_money(metrics.get("debit_status_amounts", {}), ITC_WALLET_CREDIT_TRANSFER_STATUS)),
        _wallet_row("ITC", "Debit", "Not Found", _count(debit_counts, ITC_WALLET_NOT_FOUND_STATUS), _status_money(metrics.get("debit_status_amounts", {}), ITC_WALLET_NOT_FOUND_STATUS)),
        _wallet_row("ITC", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _wallet_row("ITC", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _wallet_row(wallet: str, group: str, label: str, count: int | str, amount: Any, notes: str = "") -> list[Any]:
    return [wallet, group, label, count, _money(amount), notes]


def _stats(result: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    value = result.get(key, {})
    return value if isinstance(value, Mapping) else {}


def _money(value: Any) -> Decimal | str:
    if value in ("", None):
        return ""
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return ""


def _status_money(values: Any, key: str) -> Decimal | str:
    if not isinstance(values, Mapping):
        return ""
    return _money(values.get(key, ""))


def _count(values: Any, key: str) -> int:
    if not isinstance(values, Mapping):
        return 0
    return int(values.get(key, 0) or 0)


def _status_notes(stats: Mapping[str, Any], *statuses: str) -> str:
    counts = stats.get("status_counts", {})
    if not isinstance(counts, Mapping):
        return ""
    return "; ".join(f"{status}: {int(counts.get(status, 0) or 0):,}" for status in statuses)


def _notes(*items: tuple[str, Any]) -> str:
    return "; ".join(f"{label}: {int(value or 0):,}" for label, value in items)
