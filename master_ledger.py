#!/usr/bin/env python3
"""Master ledger workbook for wallet-to-ledger reconciliation results."""

from __future__ import annotations

import os
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Mapping

from build_reconciliation_template import (
    FAST_XLSX_COMPRESSLEVEL,
    STYLE_HEADER,
    STYLE_MONEY,
    STYLE_TABLE_MONEY,
    STYLE_TABLE_SECTION,
    STYLE_TITLE,
    Cell,
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
    NSANO_WALLET_BANK_TO_WALLET_STATUS,
    NSANO_WALLET_OTHER_STATUS,
    NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS,
    NSANO_WALLET_TOPUP_COLLECTION_STATUS,
    NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
)


MASTER_LEDGER_RESULT_KEYS = (
    "nsano_collection_ledger_result",
    "nsano_wallet_ledger_result",
    "itc_wallet_ledger_result",
)

OVERVIEW_HEADERS = [
    "Wallet",
    "Ledger Balance (GHC)",
    "Manual Inputs (GHC)",
    "Available Funds Before Debit (GHC)",
    "Total Debit (GHC)",
    "Wallet Statement Balance (GHC)",
    "Rows",
    "Notes",
]

DETAIL_HEADERS = [
    "Wallet",
    "Group",
    "Line Item",
    "Count",
    "Amount (GHC)",
    "Notes",
]


def master_ledger_snapshot(state: Mapping[str, Any]) -> dict[str, Any]:
    snapshot: dict[str, Any] = {}
    for key in MASTER_LEDGER_RESULT_KEYS:
        result = state.get(key)
        if not isinstance(result, Mapping):
            continue
        metrics = result.get("metrics")
        if isinstance(metrics, Mapping):
            snapshot[key] = {"metrics": dict(metrics)}
    return snapshot


def build_master_ledger_bytes(snapshot: Mapping[str, Any], output_path: Path) -> bytes:
    build_master_ledger_workbook(output_path, snapshot)
    return output_path.read_bytes()


def build_master_ledger_workbook(output_path: Path, snapshot: Mapping[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Ledger Overview", "Ledger Detail"]

    overview_rows = _sheet_rows(
        "MASTER LEDGER SUMMARY",
        OVERVIEW_HEADERS,
        [_overview_row_values(row) for row in master_ledger_overview_rows(snapshot)],
    )
    detail_rows = _sheet_rows(
        "MASTER LEDGER DETAIL",
        DETAIL_HEADERS,
        [_detail_row_values(row) for row in master_ledger_detail_rows(snapshot)],
    )

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
            overview_rows,
            len(overview_rows),
            len(OVERVIEW_HEADERS),
            [22, 20, 20, 30, 18, 28, 24, 46],
            freeze_top_row=False,
            autofilter=False,
            merges=[f"A1:{column_letter(len(OVERVIEW_HEADERS))}1"],
            style_func=_master_ledger_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            detail_rows,
            len(detail_rows),
            len(DETAIL_HEADERS),
            [18, 26, 32, 14, 18, 52],
            freeze_top_row=False,
            autofilter=False,
            merges=[f"A1:{column_letter(len(DETAIL_HEADERS))}1"],
            style_func=_master_ledger_style,
        )

    os.replace(temp_path, output_path)


def master_ledger_overview_rows(snapshot: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if "nsano_collection_ledger_result" in snapshot:
        rows.append(_nsano_collection_overview_row(snapshot["nsano_collection_ledger_result"].get("metrics", {})))
    if "nsano_wallet_ledger_result" in snapshot:
        rows.append(_nsano_overview_row(snapshot["nsano_wallet_ledger_result"].get("metrics", {})))
    if "itc_wallet_ledger_result" in snapshot:
        rows.append(_itc_overview_row(snapshot["itc_wallet_ledger_result"].get("metrics", {})))
    if len(rows) > 1:
        rows.append(_total_overview_row(rows))
    return rows


def master_ledger_detail_rows(snapshot: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if "nsano_collection_ledger_result" in snapshot:
        rows.extend(_nsano_collection_detail_rows(snapshot["nsano_collection_ledger_result"].get("metrics", {})))
    if "nsano_wallet_ledger_result" in snapshot:
        rows.extend(_nsano_detail_rows(snapshot["nsano_wallet_ledger_result"].get("metrics", {})))
    if "itc_wallet_ledger_result" in snapshot:
        rows.extend(_itc_detail_rows(snapshot["itc_wallet_ledger_result"].get("metrics", {})))
    return rows


def _nsano_overview_row(metrics: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "Wallet": "Nsano Disb",
        "Ledger Balance (GHC)": _money(metrics.get("ledger_balance", "")),
        "Manual Inputs (GHC)": _money(
            _decimal_amount(metrics.get("manual_unidentified"))
            + _decimal_amount(metrics.get("settlement"))
            + _decimal_amount(metrics.get("reversal"))
        ),
        "Available Funds Before Debit (GHC)": _money(metrics.get("available_funds", "")),
        "Total Debit (GHC)": _money(metrics.get("total_debit", "")),
        "Wallet Statement Balance (GHC)": _money(metrics.get("wallet_statement_balance", "")),
        "Rows": _notes(
            ("Mambu", metrics.get("mambu_total", 0)),
            ("Nsano disb", metrics.get("nsano_disb_total", 0)),
            ("Transfers", metrics.get("transfer_total", 0)),
        ),
        "Notes": _month_note(metrics),
    }


def _nsano_collection_overview_row(metrics: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "Wallet": "Nsano Collections",
        "Ledger Balance (GHC)": _money(metrics.get("ledger_balance", "")),
        "Manual Inputs (GHC)": _money(
            _decimal_amount(metrics.get("manual_unidentified"))
            + _decimal_amount(metrics.get("settlement"))
            + _decimal_amount(metrics.get("reversal"))
        ),
        "Available Funds Before Debit (GHC)": _money(metrics.get("available_funds", "")),
        "Total Debit (GHC)": _money(metrics.get("total_debit", "")),
        "Wallet Statement Balance (GHC)": _money(metrics.get("wallet_statement_balance", "")),
        "Rows": _notes(
            ("Mambu collections", metrics.get("mambu_collection_total", 0)),
            ("Mambu disb", metrics.get("mambu_disb_total", 0)),
            ("Transfers", metrics.get("transfer_total", 0)),
            ("Nsano charges", metrics.get("nsano_charge_total", 0)),
        ),
        "Notes": _month_note(metrics),
    }


def _itc_overview_row(metrics: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "Wallet": "ITC",
        "Ledger Balance (GHC)": _money(metrics.get("ledger_balance", "")),
        "Manual Inputs (GHC)": _money(metrics.get("manual_unidentified", "")),
        "Available Funds Before Debit (GHC)": _money(metrics.get("available_funds", "")),
        "Total Debit (GHC)": _money(metrics.get("total_debit", "")),
        "Wallet Statement Balance (GHC)": _money(metrics.get("wallet_statement_balance", "")),
        "Rows": _notes(
            ("Statement", metrics.get("statement_total", 0)),
            ("Debit", metrics.get("debit_total", 0)),
            ("Credit", metrics.get("credit_total", 0)),
        ),
        "Notes": _month_note(metrics),
    }


def _total_overview_row(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "Wallet": "Total",
        "Ledger Balance (GHC)": _sum_row_amounts(rows, "Ledger Balance (GHC)"),
        "Manual Inputs (GHC)": _sum_row_amounts(rows, "Manual Inputs (GHC)"),
        "Available Funds Before Debit (GHC)": _sum_row_amounts(rows, "Available Funds Before Debit (GHC)"),
        "Total Debit (GHC)": _sum_row_amounts(rows, "Total Debit (GHC)"),
        "Wallet Statement Balance (GHC)": _sum_row_amounts(rows, "Wallet Statement Balance (GHC)"),
        "Rows": "",
        "Notes": "Combined wallet ledger results.",
    }


def _nsano_detail_rows(metrics: Mapping[str, Any]) -> list[dict[str, Any]]:
    counts = metrics.get("transfer_category_counts", {})
    amounts = metrics.get("transfer_category_amounts", {})
    return [
        _detail_row("Nsano Disb", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _detail_row("Nsano Disb", "Manual Ledger Inputs", "Manual Delayed Transactions", "", metrics.get("manual_unidentified", "")),
        _detail_row("Nsano Disb", "Manual Ledger Inputs", "Manual Recovery from write off", "", metrics.get("settlement", "")),
        _detail_row("Nsano Disb", "Manual Ledger Inputs", "Manual Reversal", "", metrics.get("reversal", "")),
        _detail_row("Nsano Disb", "Manual Ledger Inputs", "Starting Ledger Balance", "", metrics.get("starting_ledger_balance", "")),
        _detail_row("Nsano Disb", "Credit / Available Funds", "Top-up Collections", _count(counts, NSANO_WALLET_TOPUP_COLLECTION_STATUS), metrics.get("topup_collections", "")),
        _detail_row("Nsano Disb", "Credit / Available Funds", "Bank to Wallet", _count(counts, NSANO_WALLET_BANK_TO_WALLET_STATUS), metrics.get("bank_to_wallet", "")),
        _detail_row("Nsano Disb", "Credit / Available Funds", "Reversal Adjustment", _count(counts, NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS), metrics.get("reversal_adjustment", "")),
        _detail_row("Nsano Disb", "Credit / Available Funds", "Other Transfers", _count(counts, NSANO_WALLET_OTHER_STATUS), _status_money(amounts, NSANO_WALLET_OTHER_STATUS)),
        _detail_row("Nsano Disb", "Credit / Available Funds", "Available Funds Before Debit", "", metrics.get("available_funds", "")),
        _detail_row("Nsano Disb", "Debit", "Transfer to Bank", _count(counts, NSANO_WALLET_TRANSFER_TO_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _detail_row("Nsano Disb", "Debit", "Disbursement", int(metrics.get("mambu_total", 0) or 0), metrics.get("disbursement", ""), "Mambu disbursement rows."),
        _detail_row("Nsano Disb", "Debit", "Delayed Transactions", "", metrics.get("unidentified_debit", ""), "Negative manual delayed transactions, shown as a positive debit."),
        _detail_row("Nsano Disb", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _detail_row("Nsano Disb", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _nsano_collection_detail_rows(metrics: Mapping[str, Any]) -> list[dict[str, Any]]:
    counts = metrics.get("transfer_category_counts", {})
    amounts = metrics.get("transfer_category_amounts", {})
    return [
        _detail_row("Nsano Collections", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _detail_row("Nsano Collections", "Manual Ledger Inputs", "Manual Delayed Transactions", "", metrics.get("manual_unidentified", "")),
        _detail_row("Nsano Collections", "Manual Ledger Inputs", "Manual Recovery from write off", "", metrics.get("settlement", "")),
        _detail_row("Nsano Collections", "Manual Ledger Inputs", "Manual Reversal", "", metrics.get("reversal", "")),
        _detail_row("Nsano Collections", "Manual Ledger Inputs", "Starting Ledger Balance", "", metrics.get("starting_ledger_balance", "")),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Total Nsano Mambu Collections", int(metrics.get("mambu_collection_total", 0) or 0), metrics.get("topup_collections", "")),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Bank to Wallet", _count(counts, NSANO_WALLET_BANK_TO_WALLET_STATUS), metrics.get("bank_to_wallet", "")),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Reversal Adjustment", _count(counts, NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS), metrics.get("reversal_adjustment", "")),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Top-up Transfer Rows", _count(counts, NSANO_WALLET_TOPUP_COLLECTION_STATUS), metrics.get("transfer_topup_collections", ""), "Retained for audit only."),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Other Transfers", _count(counts, NSANO_WALLET_OTHER_STATUS), _status_money(amounts, NSANO_WALLET_OTHER_STATUS)),
        _detail_row("Nsano Collections", "Credit / Available Funds", "Available Funds Before Debit", "", metrics.get("available_funds", "")),
        _detail_row("Nsano Collections", "Debit", "Transfer to Bank", _count(counts, NSANO_WALLET_TRANSFER_TO_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _detail_row("Nsano Collections", "Debit", "Disbursement", int(metrics.get("mambu_disb_total", 0) or 0), metrics.get("disbursement", ""), "Mambu disbursement rows from the Mambu Filtered Disbursement upload."),
        _detail_row("Nsano Collections", "Debit", "Charges", int(metrics.get("nsano_charge_total", 0) or 0), metrics.get("charges", ""), "Nsano Filtered Data W2A column U (Charge), optional upload."),
        _detail_row("Nsano Collections", "Debit", "Delayed Transactions", "", metrics.get("unidentified_debit", ""), "Negative manual delayed transactions, shown as a positive debit."),
        _detail_row("Nsano Collections", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _detail_row("Nsano Collections", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _itc_detail_rows(metrics: Mapping[str, Any]) -> list[dict[str, Any]]:
    debit_counts = metrics.get("debit_status_counts", {})
    credit_counts = metrics.get("credit_category_counts", {})
    debit_amounts = metrics.get("debit_status_amounts", {})
    credit_amounts = metrics.get("credit_category_amounts", {})
    return [
        _detail_row("ITC", "Manual Ledger Inputs", "Ledger Balance", "", metrics.get("ledger_balance", "")),
        _detail_row("ITC", "Manual Ledger Inputs", "Manual Unidentified", "", metrics.get("manual_unidentified", "")),
        _detail_row("ITC", "Manual Ledger Inputs", "Unidentified Credit", "", metrics.get("unidentified_credit", "")),
        _detail_row("ITC", "Manual Ledger Inputs", "Unidentified Debit", "", metrics.get("unidentified_debit", "")),
        _detail_row("ITC", "Credit / Available Funds", "Settlement", _count(credit_counts, ITC_WALLET_SETTLEMENT_STATUS), metrics.get("settlement", "")),
        _detail_row("ITC", "Credit / Available Funds", "Reversal", _count(credit_counts, ITC_WALLET_REVERSAL_STATUS), metrics.get("reversal", "")),
        _detail_row("ITC", "Credit / Available Funds", "Transfer to Wallet", _count(credit_counts, ITC_WALLET_TRANSFER_TO_WALLET_STATUS), metrics.get("transfer_to_wallet", "")),
        _detail_row("ITC", "Credit / Available Funds", "Prepaid Reversal", _count(credit_counts, ITC_WALLET_PREPAID_REVERSAL_STATUS), _status_money(credit_amounts, ITC_WALLET_PREPAID_REVERSAL_STATUS)),
        _detail_row("ITC", "Credit / Available Funds", "Available Funds Before Debit", "", metrics.get("available_funds", "")),
        _detail_row("ITC", "Debit", "Transfer to Bank", _count(debit_counts, ITC_WALLET_TRANSFERS_FROM_BANK_STATUS), metrics.get("transfer_to_bank", "")),
        _detail_row("ITC", "Debit", "Referral Awards", _count(debit_counts, ITC_WALLET_REFERRAL_AWARD_STATUS), metrics.get("referral_awards", "")),
        _detail_row("ITC", "Debit", "Savings", _count(debit_counts, ITC_WALLET_SAVINGS_STATUS), metrics.get("savings", "")),
        _detail_row("ITC", "Debit", "Upsales Refund", _count(debit_counts, ITC_WALLET_UPSALES_REFUND_STATUS), metrics.get("upsales_refund", "")),
        _detail_row("ITC", "Debit", "Disbursement", _count(debit_counts, ITC_WALLET_DISBURSEMENT_STATUS), metrics.get("disbursement", "")),
        _detail_row("ITC", "Debit", "Credit Transfer Fallback", _count(debit_counts, ITC_WALLET_CREDIT_TRANSFER_STATUS), _status_money(debit_amounts, ITC_WALLET_CREDIT_TRANSFER_STATUS)),
        _detail_row("ITC", "Debit", "Not Found", _count(debit_counts, ITC_WALLET_NOT_FOUND_STATUS), _status_money(debit_amounts, ITC_WALLET_NOT_FOUND_STATUS)),
        _detail_row("ITC", "Debit", "Total Debit", "", metrics.get("total_debit", "")),
        _detail_row("ITC", "Balance", "Wallet Statement Balance", "", metrics.get("wallet_statement_balance", "")),
    ]


def _detail_row(wallet: str, group: str, line_item: str, count: int | str, amount: Any, notes: str = "") -> dict[str, Any]:
    return {
        "Wallet": wallet,
        "Group": group,
        "Line Item": line_item,
        "Count": count,
        "Amount (GHC)": _money(amount),
        "Notes": notes,
    }


def _sheet_rows(title: str, headers: list[str], rows: list[list[Any]]) -> list[list[Any]]:
    col_count = len(headers)
    return [
        [Cell(title, STYLE_TITLE), *("" for _ in range(col_count - 1))],
        [*("" for _ in range(col_count))],
        [Cell(header, STYLE_HEADER) for header in headers],
        *rows,
    ]


def _overview_row_values(row: Mapping[str, Any]) -> list[Any]:
    values = [row.get(header, "") for header in OVERVIEW_HEADERS]
    if str(row.get("Wallet", "")).casefold() == "total":
        return [
            Cell(values[0], STYLE_TABLE_SECTION),
            *(Cell(value, STYLE_TABLE_MONEY if isinstance(value, Decimal) else STYLE_TABLE_SECTION) for value in values[1:]),
        ]
    return values


def _detail_row_values(row: Mapping[str, Any]) -> list[Any]:
    return [row.get(header, "") for header in DETAIL_HEADERS]


def _master_ledger_style(row_number: int, _col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    if row_number == 3:
        return STYLE_HEADER
    if isinstance(value, Decimal):
        return STYLE_MONEY
    return None


def _sum_row_amounts(rows: list[Mapping[str, Any]], key: str) -> Decimal:
    return sum((_decimal_amount(row.get(key)) for row in rows), Decimal("0"))


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


def _decimal_amount(value: Any) -> Decimal:
    amount = _money(value)
    return amount if isinstance(amount, Decimal) else Decimal("0")


def _count(values: Any, key: str) -> int:
    if not isinstance(values, Mapping):
        return 0
    return int(values.get(key, 0) or 0)


def _notes(*items: tuple[str, Any]) -> str:
    return "; ".join(f"{label}: {int(value or 0):,}" for label, value in items)


def _month_note(metrics: Mapping[str, Any]) -> str:
    month = str(metrics.get("month_title") or "").strip()
    return f"Wallet month: {month}" if month else ""
