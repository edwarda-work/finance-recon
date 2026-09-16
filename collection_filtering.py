#!/usr/bin/env python3
"""Filtering automations for collection source files."""

from __future__ import annotations

import os
import csv
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Iterable
from xml.etree.ElementTree import iterparse

from build_reconciliation_template import (
    CELL_REF_RE,
    FAST_XLSX_COMPRESSLEVEL,
    STYLE_HEADER,
    STYLE_MONEY,
    STYLE_WALLET_BALANCE_LABEL,
    STYLE_WALLET_BALANCE_MONEY,
    STYLE_WALLET_COMPANY,
    STYLE_WALLET_FINAL_LABEL,
    STYLE_WALLET_FINAL_MONEY,
    STYLE_WALLET_HEADER,
    STYLE_WALLET_LINE_ITEM,
    STYLE_WALLET_MONEY,
    STYLE_WALLET_SIGNATURE,
    STYLE_WALLET_TITLE,
    VODAFONE_CHARGE_DETAIL,
    Cell,
    app_xml,
    cell_value,
    column_letter,
    col_index,
    compute_widths,
    content_types_xml,
    core_xml,
    data_style,
    first_worksheet_path,
    iter_csv_dicts,
    iter_xlsx_dicts,
    read_shared_strings,
    root_rels_xml,
    set_csv_field_limit,
    styles_xml,
    workbook_rels_xml,
    workbook_xml,
    worksheet_paths,
    write_worksheet,
)


@dataclass(frozen=True)
class CollectionFilter:
    name: str
    sheet_name: str
    predicate: Callable[[dict[str, str]], bool]


@dataclass
class CollectionFilteringResult:
    input_rows: int
    filtered_rows: int
    filter_counts: dict[str, int]
    filter_amounts: dict[str, Decimal]
    metrics: dict[str, Any] | None = None


@dataclass(frozen=True)
class VodafoneCleanupInspection:
    is_clean: bool
    message: str
    headers: list[str]


AIRTEL_TIGO_WALLETS: tuple[tuple[str, str], ...] = (
    ("Airtel Repayment", "0266081377"),
    ("Airtel Disbursement", "0261046218"),
    ("AirtelTigo Disbursement", "0266930184"),
    ("Tigo Repayment", "027322026"),
    ("Tigo Disbursement", "0278796639"),
)


def _norm(value: Any) -> str:
    return str(value or "").strip().casefold()


def _compact(value: Any) -> str:
    return " ".join(_norm(value).replace("(", " ").replace(")", " ").split())


def _row_value(row: dict[str, str], header: str) -> str:
    if header in row:
        return row[header]
    wanted = header.casefold()
    for key, value in row.items():
        # csv.DictReader uses a None key when a data row contains more values
        # than its header. Ignore that overflow entry during header lookup.
        if key is not None and str(key).casefold() == wanted:
            return value
    return ""


def _channel_is(expected: str) -> Callable[[dict[str, str]], bool]:
    expected_norm = _norm(expected)
    return lambda row: _norm(_row_value(row, "Channel")) == expected_norm


def _channel_startswith(prefix: str) -> Callable[[dict[str, str]], bool]:
    prefix_norm = _norm(prefix)
    return lambda row: _norm(_row_value(row, "Channel")).startswith(prefix_norm)


def _column_is(header: str, expected: str) -> Callable[[dict[str, str]], bool]:
    expected_norm = _norm(expected)
    return lambda row: _norm(_row_value(row, header)) == expected_norm


def _column_contains(header: str, text: str) -> Callable[[dict[str, str]], bool]:
    text_norm = _compact(text)
    return lambda row: text_norm in _compact(_row_value(row, header))


def _column_contains_all(header: str, *parts: str) -> Callable[[dict[str, str]], bool]:
    part_norms = [_compact(part) for part in parts]
    return lambda row: all(part in _compact(_row_value(row, header)) for part in part_norms)


def _column_matches_any(header: str, *values: str) -> Callable[[dict[str, str]], bool]:
    value_norms = {_compact(value) for value in values}
    return lambda row: _compact(_row_value(row, header)) in value_norms


def _all_of(*predicates: Callable[[dict[str, str]], bool]) -> Callable[[dict[str, str]], bool]:
    return lambda row: all(predicate(row) for predicate in predicates)


def _any_of(*predicates: Callable[[dict[str, str]], bool]) -> Callable[[dict[str, str]], bool]:
    return lambda row: any(predicate(row) for predicate in predicates)


VODAFONE_COLLECTION_CHANNEL = "Vodafone Coll. Account - vodafone"

ITC_VODAFONE_COLLECTION_CHANNELS: tuple[str, ...] = (
    "AirtelTigo Coll. Account - ITC Fido",
    "Business Loans (AirtelTigo) - Client Collection - ITC",
    "Business Loans (MTN) - Client Collection - ITC",
    "Business Loans (Vodafone) - Client Collection - ITC",
    "ITC Collections Account - GIP",
    "MTN Coll. Account - ITC Fido",
    "Vodafone Coll. Account - ITC Fido",
    VODAFONE_COLLECTION_CHANNEL,
    "Vodafone Collections Account - GIP",
)

ITC_COLLECTION_CHANNELS: tuple[str, ...] = tuple(
    channel
    for channel in ITC_VODAFONE_COLLECTION_CHANNELS
    if channel != VODAFONE_COLLECTION_CHANNEL
)

MAMBU_COLLECTION_FILTERS: list[CollectionFilter] = [
    CollectionFilter(
        "ITC / Vodafone Collection",
        "ITC Vodafone Collection",
        _column_matches_any(
            "Channel",
            *ITC_VODAFONE_COLLECTION_CHANNELS,
        ),
    ),
    CollectionFilter(
        "ITC Collection",
        "ITC Collection",
        _column_matches_any(
            "Channel",
            *ITC_COLLECTION_CHANNELS,
        ),
    ),
    CollectionFilter(
        "Vodafone Collection",
        "Vodafone Collection",
        _column_is("Channel", VODAFONE_COLLECTION_CHANNEL),
    ),
    CollectionFilter(
        "Nsano Client Collection",
        "Nsano Client Collection",
        _column_matches_any(
            "Channel",
            "Nsano - Client Collections",
            "Nsano - Client Collection",
        ),
    ),
    CollectionFilter(
        "Zenith Collections",
        "Zenith Collections",
        _column_is(
            "Channel",
            "Zenith Coll. Account - 6010159660",
        ),
    ),
]


ITC_COLLECTION_FILTERS: list[CollectionFilter] = [
    CollectionFilter(
        "Inflow",
        "Inflow",
        _column_is("transaction_type", "inflow"),
    ),
    CollectionFilter(
        "Outflow",
        "Outflow",
        _column_is("transaction_type", "outflow"),
    ),
]


NSANO_COLLECTION_FILTERS: list[CollectionFilter] = [
    CollectionFilter(
        "Successful W2A",
        "Successful W2A",
        _all_of(
            _column_is("Result", "Successful"),
            _column_is("Type", "W2A"),
        ),
    ),
    CollectionFilter(
        "Successful A2W",
        "Successful A2W",
        _all_of(
            _column_is("Result", "Successful"),
            _column_is("Type", "A2W"),
        ),
    ),
]


MAMBU_DISBURSEMENT_FILTERS: list[CollectionFilter] = [
    CollectionFilter(
        "ITC Disbursement",
        "ITC Disbursement",
        _column_matches_any(
            "Channel",
            "AirtelTigo Disb. Account - Auto - 0266930184",
            "Business Loans (AirtelTigo) - Client Disbursement - ITC",
            "Business Loans (MTN) - Client Disbursement - ITC",
            "Business Loans (Vodafone) - Client Disbursement - ITC",
            "MTN Disb. Account - Auto - ITC",
            "Vodafone Disb. Account - Auto - ITC",
        ),
    ),
    CollectionFilter(
        "Nsano Client Disbursement",
        "Nsano Client Disb",
        _column_matches_any("Channel", "Nsano - Client Disbursement"),
    ),
    CollectionFilter(
        "Maxbuy Client Disbursement",
        "Maxbuy Client Disbursement",
        _column_matches_any("Channel", "Maxbuy - Client Disbursement"),
    ),
    CollectionFilter(
        "MTN Manual Disbursement",
        "MTN Manual Disb",
        _all_of(
            _column_contains("Channel", "MTN"),
            _column_contains("Channel", "Manual"),
            _column_contains("Channel", "Disbursement"),
        ),
    ),
    CollectionFilter(
        "Vodafone Manual Disbursement",
        "Vodafone Manual Disb",
        _all_of(
            _column_contains("Channel", "Vodafone"),
            _column_contains("Channel", "Manual"),
            _column_contains("Channel", "Disbursement"),
        ),
    ),
]


VODAFONE_COLLECTION_CLEANUP_FILTERS: list[CollectionFilter] = [
    CollectionFilter("Cleaned Vodafone Collection", "Cleaned Voda Coll", lambda _row: True),
]

VODAFONE_CLEANED_COLLECTION_COLUMNS: list[str] = [
    "Receipt No.",
    "Completion Time",
    "Initiation Time",
    "Details",
    "Transaction Status",
    "Currency",
    "Paid In",
    "Withdrawn",
    "Balance",
    "Reason Type",
    "Opposite Party",
]

VODAFONE_LEDGER_TRANSFER_DETAIL = "Funds Transfer from Utility Account to Working Account"


def iter_source_rows(path: Path) -> Iterable[tuple[int, dict[str, str]]]:
    if path.suffix.lower() == ".csv":
        yield from iter_csv_dicts(path)
    else:
        yield from iter_xlsx_dicts(path)


def _add_headers(headers: list[str], new_headers: Iterable[str]) -> None:
    for header in new_headers:
        if header is not None and header not in headers:
            headers.append(header)


def _require_headers(headers: Iterable[str], *required: str) -> None:
    """Raise a useful input error when required filtering columns are absent."""
    available = {str(header).strip().casefold() for header in headers if header is not None}
    missing = [header for header in required if header.casefold() not in available]
    if missing:
        raise ValueError(
            "Nsano filtering requires the following column(s): "
            + ", ".join(missing)
            + ". Check that the file has the correct header row and that data rows align with it."
        )


def load_collection_rows(paths: Iterable[Path]) -> tuple[list[str], list[dict[str, str]]]:
    headers: list[str] = []
    rows: list[dict[str, str]] = []
    for path in paths:
        if not path or not path.exists():
            continue
        for _source_row, row in iter_source_rows(path):
            _add_headers(headers, row.keys())
            rows.append(row)
    return headers, rows


def filter_collection_rows(
    rows: list[dict[str, str]],
    filters: list[CollectionFilter] | None = None,
) -> dict[str, list[dict[str, str]]]:
    rules = filters or MAMBU_COLLECTION_FILTERS
    return {rule.name: [row for row in rows if rule.predicate(row)] for rule in rules}


def amount_to_decimal(value: Any) -> Decimal:
    text = str(value or "").strip().replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _amount_value(row: dict[str, str]) -> str:
    for header in ("Amount", "amount", "Amount (GHC)", "Amount_GHC", "Paid In", "Principal Amount", "net_amount"):
        value = _row_value(row, header)
        if value not in ("", None):
            return value
    return ""


def sum_amounts(rows: Iterable[dict[str, str]]) -> Decimal:
    return sum((amount_to_decimal(_amount_value(row)) for row in rows), Decimal("0"))


def _rows_for_sheet(headers: list[str], rows: list[dict[str, str]]) -> Iterable[list[Any]]:
    yield headers
    for row in rows:
        yield [_row_value(row, header) for header in headers]


def _preview_rows(headers: list[str], rows: list[dict[str, str]], limit: int = 5000) -> Iterable[list[Any]]:
    for row in rows[:limit]:
        yield [_row_value(row, header) for header in headers]


def _all_filtered_rows(filtered: dict[str, list[dict[str, str]]]) -> list[dict[str, str]]:
    seen: set[int] = set()
    combined: list[dict[str, str]] = []
    for rows in filtered.values():
        for row in rows:
            marker = id(row)
            if marker in seen:
                continue
            seen.add(marker)
            combined.append(row)
    return combined


def write_collection_filter_workbook(
    output_path: Path,
    headers: list[str],
    filtered: dict[str, list[dict[str, str]]],
    filters: list[CollectionFilter] | None = None,
    *,
    include_all_filtered: bool = True,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")

    rules = filters or MAMBU_COLLECTION_FILTERS
    all_filtered = _all_filtered_rows(filtered)
    sheet_specs = [
        (rule.sheet_name, filtered.get(rule.name, [])) for rule in rules
    ]
    if include_all_filtered:
        sheet_specs.insert(0, ("All Filtered", all_filtered))
    sheet_names = [name for name, _rows in sheet_specs]

    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=FAST_XLSX_COMPRESSLEVEL, allowZip64=True) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheet_names)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml(sheet_names))
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheet_names)))
        zf.writestr("xl/styles.xml", styles_xml())

        for index, (_sheet_name, rows) in enumerate(sheet_specs, start=1):
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{index}.xml",
                _rows_for_sheet(headers, rows),
                len(rows) + 1,
                len(headers),
                compute_widths(headers, _preview_rows(headers, rows)),
                style_func=data_style(headers),
            )

    os.replace(temp_path, output_path)


def iter_raw_rows(path: Path) -> Iterable[list[str]]:
    if path.suffix.lower() == ".csv":
        set_csv_field_limit()
        with path.open(newline="", encoding="utf-8-sig") as handle:
            yield from csv.reader(handle)
        return

    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        worksheet_path = first_worksheet_path(zf)
        with zf.open(worksheet_path) as handle:
            for _event, elem in iterparse(handle, events=("end",)):
                if not elem.tag.endswith("}row"):
                    continue
                values_by_index: dict[int, str] = {}
                max_index = 0
                for cell in list(elem):
                    if not cell.tag.endswith("}c"):
                        continue
                    ref_match = CELL_REF_RE.match(cell.attrib.get("r", ""))
                    index = col_index(ref_match.group(1)) if ref_match else len(values_by_index) + 1
                    values_by_index[index] = cell_value(cell, shared_strings)
                    max_index = max(max_index, index)
                yield [values_by_index.get(i, "") for i in range(1, max_index + 1)]
                elem.clear()


def inspect_vodafone_collection_cleanup(path: Path) -> VodafoneCleanupInspection:
    """Confirm that a Vodafone export has the cleaned schema used by collection recon."""
    if path.suffix.lower() != ".csv":
        try:
            with zipfile.ZipFile(path) as zf:
                sheet_names = set(worksheet_paths(zf))
        except Exception as exc:
            return VodafoneCleanupInspection(False, f"Could not read the Vodafone workbook: {exc}", [])

        if "cleaned voda coll" not in sheet_names:
            return VodafoneCleanupInspection(
                False,
                "The workbook does not contain the `Cleaned Voda Coll` sheet created by Voda Collection Cleanup.",
                [],
            )

    for raw_row in iter_raw_rows(path):
        normalized = [_norm(value) for value in raw_row]
        if "receipt no." not in normalized:
            continue

        headers = [str(value).strip() for value in raw_row]
        normalized_headers = {_norm(header) for header in headers if header}
        required = {_norm(header) for header in VODAFONE_CLEANED_COLLECTION_COLUMNS}
        missing = sorted(required - normalized_headers)
        if missing:
            return VodafoneCleanupInspection(
                False,
                "The Vodafone file does not match the cleaned collection layout. Missing: "
                + ", ".join(missing),
                headers,
            )

        return VodafoneCleanupInspection(
            True,
            "Vodafone collection cleanup confirmed.",
            headers,
        )

    return VodafoneCleanupInspection(
        False,
        "Could not find the Vodafone header row containing `Receipt No.`.",
        [],
    )


VODAFONE_HEADER_SCAN_LIMIT = 25


def _is_vodafone_collection_header(raw_row: list[str]) -> bool:
    normalized = {_norm(value) for value in raw_row if str(value).strip()}
    return "receipt no." in normalized


def load_vodafone_collection_rows(
    paths: Iterable[Path],
    *,
    keep_withdrawn: bool = False,
) -> tuple[list[str], list[dict[str, str]]]:
    """Combine Vodafone exports while retaining all columns, including Withdrawn."""
    headers: list[str] = []
    rows: list[dict[str, str]] = []

    for path in paths:
        source_headers: list[str] | None = None
        for row_index, raw_row in enumerate(iter_raw_rows(path)):
            if source_headers is None:
                if not _is_vodafone_collection_header(raw_row):
                    if row_index >= VODAFONE_HEADER_SCAN_LIMIT:
                        break
                    continue
                source_headers = [value.strip() for value in raw_row]
                _add_headers(headers, source_headers)
                continue

            if not any(str(value).strip() for value in raw_row):
                continue
            if raw_row and _norm(raw_row[0]) == "receipt no.":
                continue
            cleaned = list(raw_row)
            padded = cleaned + [""] * max(0, len(source_headers) - len(cleaned))
            rows.append(dict(zip(source_headers, padded)))

    return headers, rows


def normalize_vodafone_csv(path: Path | None) -> Path | None:
    """Strip the Vodafone account metadata block from a raw CSV export in place.

    Cleaned CSVs (header already on row 1) are rewritten with the same schema.
    """
    if not path or path.suffix.lower() != ".csv":
        return path

    headers, rows = load_vodafone_collection_rows([path])
    if not headers:
        return path

    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([row.get(header, "") for header in headers])

    return path


def _sum_column(rows: Iterable[dict[str, str]], header: str) -> Decimal:
    return sum((amount_to_decimal(_row_value(row, header)) for row in rows), Decimal("0"))


def _sum_positive_column(rows: Iterable[dict[str, str]], header: str) -> Decimal:
    return sum((abs(amount_to_decimal(_row_value(row, header))) for row in rows), Decimal("0"))


def _sum_withdrawn_by_detail(rows: Iterable[dict[str, str]], detail_text: str) -> tuple[int, Decimal]:
    detail_norm = _compact(detail_text)
    matches = [
        row for row in rows
        if detail_norm in _compact(_row_value(row, "Details"))
    ]
    return len(matches), _sum_positive_column(matches, "Withdrawn")


def _vodafone_wallet_not_found_rows(rows: Iterable[dict[str, str]]) -> list[dict[str, str]]:
    """Return monetary movements not used by any Vodafone ledger line."""
    charge_detail = _compact(VODAFONE_CHARGE_DETAIL)
    transfer_detail = _compact(VODAFONE_LEDGER_TRANSFER_DETAIL)
    not_found: list[dict[str, str]] = []
    for row in rows:
        if amount_to_decimal(_row_value(row, "Paid In")) != 0:
            continue
        if amount_to_decimal(_row_value(row, "Withdrawn")) == 0:
            continue
        detail = _compact(_row_value(row, "Details"))
        if charge_detail in detail or transfer_detail in detail:
            continue
        not_found.append(row)
    return not_found


def _positive_decimal(value: Any) -> Decimal:
    return abs(amount_to_decimal(value))


def _positive_amount_text(value: Any) -> Any:
    text = str(value or "").strip()
    if not text:
        return value
    amount = amount_to_decimal(text)
    if amount == 0 and text.replace(",", "") not in {"0", "0.0", "0.00"}:
        return value
    return str(abs(amount))


def _positive_amount_rows(headers: list[str], rows: list[dict[str, str]]) -> list[dict[str, str]]:
    amount_headers = {"paid in", "withdrawn", "balance"}
    normalized_rows: list[dict[str, str]] = []
    for row in rows:
        normalized = dict(row)
        for header in headers:
            if header.strip().casefold() in amount_headers:
                normalized[header] = _positive_amount_text(_row_value(row, header))
        normalized_rows.append(normalized)
    return normalized_rows


def build_vodafone_wallet_ledger_summary(
    balance: Decimal,
    delayed_transactions: Decimal,
    rows: list[dict[str, str]],
) -> tuple[list[list[Any]], dict[str, Any]]:
    balance = abs(balance)
    collections_rows = [row for row in rows if amount_to_decimal(_row_value(row, "Paid In")) != 0]
    total_collections = _sum_positive_column(rows, "Paid In")
    delayed_credit = abs(delayed_transactions) if delayed_transactions > 0 else Decimal("0")
    delayed_debit = abs(delayed_transactions) if delayed_transactions < 0 else Decimal("0")
    available_funds = balance + total_collections + delayed_credit

    charge_count, total_charges = _sum_withdrawn_by_detail(rows, VODAFONE_CHARGE_DETAIL)
    transfer_count, transfer_to_bank = _sum_withdrawn_by_detail(rows, VODAFONE_LEDGER_TRANSFER_DETAIL)
    total_debit = total_charges + transfer_to_bank + delayed_debit
    wallet_statement_balance = abs(available_funds - total_debit)

    metrics = {
        "balance": balance,
        "delayed_transactions": abs(delayed_transactions),
        "delayed_credit": delayed_credit,
        "delayed_debit": delayed_debit,
        "total_collections": total_collections,
        "available_funds": available_funds,
        "total_charges": total_charges,
        "transfer_to_bank": transfer_to_bank,
        "total_debit": total_debit,
        "wallet_statement_balance": wallet_statement_balance,
        "input_rows": len(rows),
        "collections_count": len(collections_rows),
        "charge_count": charge_count,
        "transfer_count": transfer_count,
    }

    def wh(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_HEADER)

    def wb(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_BALANCE_LABEL)

    def wbm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_BALANCE_MONEY)

    def wm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_MONEY)

    def wf(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_FINAL_LABEL)

    def wfm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_FINAL_MONEY)

    def sig(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SIGNATURE)

    def line(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_LINE_ITEM)

    rows_out = [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell("VODAFONE WALLET VS LEDGER", STYLE_WALLET_TITLE), "", "", ""],
        ["", "", "", ""],
        [wh("Line Item"), wh("Count"), wh("Amount (GHC)"), ""],
        [wb("Balance"), "", wbm(balance), ""],
        [line("Delayed Transactions (Credit)"), "", wm(delayed_credit), ""],
        [sig("Total Collections"), len(collections_rows), wm(total_collections), ""],
        [wb("Available Funds Before Debit"), "", wbm(available_funds), ""],
        ["", "", "", ""],
        [sig("Total Charges"), charge_count, wm(total_charges), ""],
        [line("Transfer to Bank"), transfer_count, wm(transfer_to_bank), ""],
        [line("Delayed Transactions (Debit)"), "", wm(delayed_debit), ""],
        [wb("Total Debit"), "", wbm(total_debit), ""],
        ["", "", "", ""],
        [wf("Balance as per Wallet Statement"), "", wfm(wallet_statement_balance), ""],
        ["", "", "", ""],
        [sig("Prepared By:"), "", sig("Reviewed By:"), ""],
        [sig("Name:"), "", sig("Name:"), ""],
        [sig("Title:"), "", sig("Title:"), ""],
        [sig("Signature:"), "Paste signature image here", sig("Signature:"), "Paste signature image here"],
        [sig("Date:"), "", sig("Date:"), ""],
    ]
    return rows_out, metrics


def build_vodafone_manual_wallet_ledger_summary(
    balance: Decimal,
    delayed_transactions: Decimal,
    rows: list[dict[str, str]],
) -> tuple[list[list[Any]], dict[str, Any]]:
    """Build the Vodafone Manual disbursement wallet-to-ledger summary."""
    balance = abs(balance)
    disbursement_rows = [
        row for row in rows
        if amount_to_decimal(_row_value(row, "Amount")) != 0
    ]
    total_disbursement = _sum_positive_column(disbursement_rows, "Amount")
    delayed_credit = abs(delayed_transactions) if delayed_transactions > 0 else Decimal("0")
    delayed_debit = abs(delayed_transactions) if delayed_transactions < 0 else Decimal("0")
    available_funds = balance + delayed_credit
    total_debit = total_disbursement + delayed_debit
    wallet_statement_balance = available_funds - total_debit

    metrics = {
        "balance": balance,
        "delayed_transactions": abs(delayed_transactions),
        "delayed_credit": delayed_credit,
        "delayed_debit": delayed_debit,
        "total_disbursement": total_disbursement,
        "available_funds": available_funds,
        "total_debit": total_debit,
        "wallet_statement_balance": wallet_statement_balance,
        "input_rows": len(rows),
        "disbursement_count": len(disbursement_rows),
    }

    def wh(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_HEADER)

    def wb(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_BALANCE_LABEL)

    def wbm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_BALANCE_MONEY)

    def wm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_MONEY)

    def wf(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_FINAL_LABEL)

    def wfm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_FINAL_MONEY)

    def line(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_LINE_ITEM)

    def sig(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SIGNATURE)

    rows_out = [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell("VODAFONE MANUAL WALLET VS LEDGER", STYLE_WALLET_TITLE), "", "", ""],
        ["", "", "", ""],
        [wh("Line Item"), wh("Count"), wh("Amount (GHC)"), ""],
        [wb("Balance Per Ledger"), "", wbm(balance), ""],
        [line("Delayed Transactions (Credit)"), "", wm(delayed_credit), ""],
        [wb("Available Funds Before Debit"), "", wbm(available_funds), ""],
        ["", "", "", ""],
        [line("Vodafone Manual Disbursements"), len(disbursement_rows), wm(total_disbursement), ""],
        [line("Delayed Transactions (Debit)"), "", wm(delayed_debit), ""],
        [wb("Total Debit"), "", wbm(total_debit), ""],
        ["", "", "", ""],
        [wf("Balance as per Wallet Statement"), "", wfm(wallet_statement_balance), ""],
        ["", "", "", ""],
        [sig("Prepared By:"), "", sig("Reviewed By:"), ""],
        [sig("Name:"), "", sig("Name:"), ""],
        [sig("Title:"), "", sig("Title:"), ""],
        [sig("Signature:"), "Paste signature image here", sig("Signature:"), "Paste signature image here"],
        [sig("Date:"), "", sig("Date:"), ""],
    ]
    return rows_out, metrics


def _wallet_summary_style(row_number: int, _col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    if row_number == 4:
        return STYLE_HEADER
    if isinstance(value, Decimal):
        return STYLE_MONEY
    return None


def write_vodafone_wallet_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Summary", "Voda Not Found", "Cleaned Voda Coll"]
    cleaned_rows = list(_rows_for_sheet(headers, rows))
    not_found_rows = _vodafone_wallet_not_found_rows(rows)
    not_found_sheet_rows = list(_rows_for_sheet(headers, not_found_rows))

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
            summary_rows,
            len(summary_rows),
            4,
            [34, 14, 14, 14],
            freeze_top_row=False,
            autofilter=False,
            merges=[
                "A1:D1", "A2:D2", "C4:D4",
                *(f"C{row}:D{row}" for row in (5, 6, 7, 8, 10, 11, 12, 13, 15)),
                "A17:B17", "C17:D17",
            ],
            style_func=_wallet_summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            not_found_sheet_rows,
            len(not_found_sheet_rows),
            len(headers),
            compute_widths(headers, _preview_rows(headers, not_found_rows)),
            style_func=data_style(headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            cleaned_rows,
            len(cleaned_rows),
            len(headers),
            compute_widths(headers, _preview_rows(headers, rows)),
            style_func=data_style(headers),
        )

    os.replace(temp_path, output_path)


def build_mambu_collection_filter_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
) -> CollectionFilteringResult:
    headers, rows = load_collection_rows(source_paths)
    filtered = filter_collection_rows(rows)
    write_collection_filter_workbook(
        output_path,
        headers,
        filtered,
        MAMBU_COLLECTION_FILTERS,
        include_all_filtered=False,
    )

    all_filtered = _all_filtered_rows(filtered)
    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(all_filtered),
        filter_counts={name: len(filter_rows) for name, filter_rows in filtered.items()},
        filter_amounts={name: sum_amounts(filter_rows) for name, filter_rows in filtered.items()},
    )


def build_itc_collection_filter_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
) -> CollectionFilteringResult:
    headers, rows = load_collection_rows(source_paths)
    filtered = filter_collection_rows(rows, ITC_COLLECTION_FILTERS)
    write_collection_filter_workbook(
        output_path,
        headers,
        filtered,
        ITC_COLLECTION_FILTERS,
        include_all_filtered=False,
    )

    all_filtered = _all_filtered_rows(filtered)
    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(all_filtered),
        filter_counts={name: len(filter_rows) for name, filter_rows in filtered.items()},
        filter_amounts={name: sum_amounts(filter_rows) for name, filter_rows in filtered.items()},
    )


def build_nsano_collection_filter_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
) -> CollectionFilteringResult:
    headers, rows = load_collection_rows(source_paths)
    _require_headers(headers, "Result", "Type")
    filtered = filter_collection_rows(rows, NSANO_COLLECTION_FILTERS)
    write_collection_filter_workbook(
        output_path,
        headers,
        filtered,
        NSANO_COLLECTION_FILTERS,
        include_all_filtered=False,
    )

    all_filtered = _all_filtered_rows(filtered)
    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(all_filtered),
        filter_counts={name: len(filter_rows) for name, filter_rows in filtered.items()},
        filter_amounts={name: sum_amounts(filter_rows) for name, filter_rows in filtered.items()},
    )


def build_mambu_disbursement_filter_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
) -> CollectionFilteringResult:
    headers, rows = load_collection_rows(source_paths)
    filtered = filter_collection_rows(rows, MAMBU_DISBURSEMENT_FILTERS)
    write_collection_filter_workbook(
        output_path,
        headers,
        filtered,
        MAMBU_DISBURSEMENT_FILTERS,
        include_all_filtered=False,
    )

    all_filtered = _all_filtered_rows(filtered)
    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(all_filtered),
        filter_counts={name: len(filter_rows) for name, filter_rows in filtered.items()},
        filter_amounts={name: sum_amounts(filter_rows) for name, filter_rows in filtered.items()},
    )


def build_vodafone_collection_cleanup_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
) -> CollectionFilteringResult:
    headers, rows = load_vodafone_collection_rows(source_paths)
    if not headers:
        raise ValueError(
            "Vodafone cleanup could not find a header row containing `Receipt No.`."
        )
    cleaned = {"Cleaned Vodafone Collection": rows}
    write_collection_filter_workbook(
        output_path,
        headers,
        cleaned,
        VODAFONE_COLLECTION_CLEANUP_FILTERS,
        include_all_filtered=False,
    )

    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(rows),
        filter_counts={"Cleaned Vodafone Collection": len(rows)},
        filter_amounts={"Cleaned Vodafone Collection": sum_amounts(rows)},
    )


def build_vodafone_wallet_ledger_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
    balance: Decimal,
    delayed_transactions: Decimal,
) -> CollectionFilteringResult:
    headers, rows = load_vodafone_collection_rows(source_paths, keep_withdrawn=True)
    positive_rows = _positive_amount_rows(headers, rows)
    summary_rows, metrics = build_vodafone_wallet_ledger_summary(
        balance,
        delayed_transactions,
        positive_rows,
    )
    write_vodafone_wallet_ledger_workbook(
        output_path,
        summary_rows,
        headers,
        positive_rows,
    )

    return CollectionFilteringResult(
        input_rows=len(positive_rows),
        filtered_rows=len(positive_rows),
        filter_counts={"Cleaned Vodafone Collection": len(positive_rows)},
        filter_amounts={"Cleaned Vodafone Collection": sum_amounts(positive_rows)},
        metrics=metrics,
    )


def load_vodafone_manual_rows(
    paths: Iterable[Path],
) -> tuple[list[str], list[dict[str, str]]]:
    """Load raw or already-cleaned Vodafone Manual disbursement exports."""
    headers: list[str] = []
    rows: list[dict[str, str]] = []
    required = {"id", "date", "amount"}
    for path in paths:
        source_headers: list[str] | None = None
        for row_index, raw_row in enumerate(iter_raw_rows(path)):
            if source_headers is None:
                normalized = {_norm(value) for value in raw_row if str(value).strip()}
                if not required.issubset(normalized):
                    if row_index >= 25:
                        break
                    continue
                source_headers = [str(value or "").strip() for value in raw_row]
                _add_headers(headers, source_headers)
                continue
            if not any(str(value or "").strip() for value in raw_row):
                continue
            normalized_first = _norm(raw_row[0]) if raw_row else ""
            if normalized_first == "id":
                continue
            padded = list(raw_row) + [""] * max(0, len(source_headers) - len(raw_row))
            rows.append(dict(zip(source_headers, padded)))
    if not headers:
        raise ValueError(
            "Vodafone Manual Wallet vs Ledger could not find a header row containing Id, Date, and Amount."
        )
    return headers, rows


def write_vodafone_manual_wallet_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Summary", "Vodafone Manual"]
    source_rows = list(_rows_for_sheet(headers, rows))

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
            summary_rows,
            len(summary_rows),
            4,
            [36, 14, 15, 15],
            freeze_top_row=False,
            autofilter=False,
            merges=[
                "A1:D1", "A2:D2", "C4:D4",
                *(f"C{row}:D{row}" for row in (5, 6, 7, 9, 10, 11, 13)),
                "A15:B15", "C15:D15",
            ],
            style_func=_wallet_summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            source_rows,
            len(source_rows),
            len(headers),
            compute_widths(headers, _preview_rows(headers, rows)),
            style_func=data_style(headers),
        )
    os.replace(temp_path, output_path)


def build_vodafone_manual_wallet_ledger_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
    balance: Decimal,
    delayed_transactions: Decimal,
) -> CollectionFilteringResult:
    headers, rows = load_vodafone_manual_rows(source_paths)
    summary_rows, metrics = build_vodafone_manual_wallet_ledger_summary(
        balance,
        delayed_transactions,
        rows,
    )
    write_vodafone_manual_wallet_ledger_workbook(
        output_path,
        summary_rows,
        headers,
        rows,
    )
    return CollectionFilteringResult(
        input_rows=len(rows),
        filtered_rows=len(rows),
        filter_counts={"Vodafone Manual Disbursements": metrics["disbursement_count"]},
        filter_amounts={"Vodafone Manual Disbursements": metrics["total_disbursement"]},
        metrics=metrics,
    )


def calculate_airtel_tigo_wallet_ledger(values: dict[str, Any]) -> dict[str, Decimal]:
    """Calculate one Airtel/Tigo wallet using directional delayed transactions."""
    opening_balance = amount_to_decimal(values.get("opening_balance", 0))
    total_repayment = amount_to_decimal(values.get("total_repayment", 0))
    interest_received = amount_to_decimal(values.get("interest_received", 0))
    charges = abs(amount_to_decimal(values.get("charges", 0)))
    total_disbursement = abs(amount_to_decimal(values.get("total_disbursement", 0)))
    delayed_transaction = amount_to_decimal(values.get("delayed_transaction", 0))
    delayed_credit = delayed_transaction if delayed_transaction > 0 else Decimal("0")
    delayed_debit = abs(delayed_transaction) if delayed_transaction < 0 else Decimal("0")
    available_funds = opening_balance + total_repayment + interest_received + delayed_credit
    total_debit = total_disbursement + charges + delayed_debit
    return {
        "opening_balance": opening_balance,
        "total_repayment": total_repayment,
        "interest_received": interest_received,
        "charges": charges,
        "total_disbursement": total_disbursement,
        "delayed_transaction": delayed_transaction,
        "delayed_credit": delayed_credit,
        "delayed_debit": delayed_debit,
        "available_funds": available_funds,
        "total_debit": total_debit,
        "wallet_balance": available_funds - total_debit,
    }


def _airtel_tigo_wallet_detail_rows(
    wallet_name: str,
    account_number: str,
    metrics: dict[str, Decimal],
) -> list[list[Any]]:
    def sig(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SIGNATURE)

    return [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell(f"{wallet_name.upper()} WALLET VS LEDGER", STYLE_WALLET_TITLE), "", "", ""],
        [Cell(f"Wallet number: {account_number}", STYLE_WALLET_SIGNATURE), "", "", ""],
        [Cell("Line Item", STYLE_WALLET_HEADER), "", Cell("Amount (GHC)", STYLE_WALLET_HEADER), ""],
        [Cell("Opening Balance as per Ledger", STYLE_WALLET_BALANCE_LABEL), "", Cell(metrics["opening_balance"], STYLE_WALLET_BALANCE_MONEY), ""],
        [Cell("Total Repayment", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["total_repayment"], STYLE_WALLET_MONEY), ""],
        [Cell("Interest Received", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["interest_received"], STYLE_WALLET_MONEY), ""],
        [Cell("Delayed Transaction (Credit)", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["delayed_credit"], STYLE_WALLET_MONEY), ""],
        [Cell("Available Funds Before Debit", STYLE_WALLET_BALANCE_LABEL), "", Cell(metrics["available_funds"], STYLE_WALLET_BALANCE_MONEY), ""],
        ["", "", "", ""],
        [Cell("Total Disbursement", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["total_disbursement"], STYLE_WALLET_MONEY), ""],
        [Cell("Charges", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["charges"], STYLE_WALLET_MONEY), ""],
        [Cell("Delayed Transaction (Debit)", STYLE_WALLET_LINE_ITEM), "", Cell(metrics["delayed_debit"], STYLE_WALLET_MONEY), ""],
        [Cell("Total Debit", STYLE_WALLET_BALANCE_LABEL), "", Cell(metrics["total_debit"], STYLE_WALLET_BALANCE_MONEY), ""],
        ["", "", "", ""],
        [Cell("Balance as per Wallet", STYLE_WALLET_FINAL_LABEL), "", Cell(metrics["wallet_balance"], STYLE_WALLET_FINAL_MONEY), ""],
        ["", "", "", ""],
        [sig("Prepared By:"), "", sig("Reviewed By:"), ""],
        [sig("Name:"), "", sig("Name:"), ""],
        [sig("Title:"), "", sig("Title:"), ""],
        [sig("Signature:"), "Paste signature image here", sig("Signature:"), "Paste signature image here"],
        [sig("Date:"), "", sig("Date:"), ""],
    ]


def build_airtel_tigo_wallet_ledger_workbook(
    output_path: Path,
    wallet_inputs: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Build the consolidated Airtel/Tigo wallet-to-ledger workbook."""
    calculated: dict[str, dict[str, Any]] = {}
    for wallet_name, account_number in AIRTEL_TIGO_WALLETS:
        calculated[wallet_name] = {
            "account_number": account_number,
            **calculate_airtel_tigo_wallet_ledger(wallet_inputs.get(wallet_name, {})),
        }

    summary_rows: list[list[Any]] = [
        [Cell("AIRTEL / TIGO WALLET VS LEDGER — SUMMARY", STYLE_WALLET_TITLE), *([""] * 9)],
        [Cell("Positive delayed transactions are credits; negative delayed transactions are debits.", STYLE_WALLET_SIGNATURE), *([""] * 9)],
        ["" for _ in range(10)],
        [
            Cell("Wallet", STYLE_WALLET_HEADER),
            Cell("Number", STYLE_WALLET_HEADER),
            Cell("Opening Balance", STYLE_WALLET_HEADER),
            Cell("Repayment", STYLE_WALLET_HEADER),
            Cell("Interest", STYLE_WALLET_HEADER),
            Cell("Delayed Credit", STYLE_WALLET_HEADER),
            Cell("Available Funds", STYLE_WALLET_HEADER),
            Cell("Total Debit", STYLE_WALLET_HEADER),
            Cell("Delayed Debit", STYLE_WALLET_HEADER),
            Cell("Wallet Balance", STYLE_WALLET_HEADER),
        ],
    ]
    for wallet_name, account_number in AIRTEL_TIGO_WALLETS:
        metrics = calculated[wallet_name]
        summary_rows.append([
            Cell(wallet_name, STYLE_WALLET_LINE_ITEM),
            account_number,
            Cell(metrics["opening_balance"], STYLE_WALLET_MONEY),
            Cell(metrics["total_repayment"], STYLE_WALLET_MONEY),
            Cell(metrics["interest_received"], STYLE_WALLET_MONEY),
            Cell(metrics["delayed_credit"], STYLE_WALLET_MONEY),
            Cell(metrics["available_funds"], STYLE_WALLET_BALANCE_MONEY),
            Cell(metrics["total_debit"], STYLE_WALLET_BALANCE_MONEY),
            Cell(metrics["delayed_debit"], STYLE_WALLET_MONEY),
            Cell(metrics["wallet_balance"], STYLE_WALLET_FINAL_MONEY),
        ])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Summary", *(name[:31] for name, _number in AIRTEL_TIGO_WALLETS)]
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
            summary_rows,
            len(summary_rows),
            10,
            [27, 17, 19, 17, 17, 18, 20, 18, 18, 20],
            freeze_top_row=False,
            autofilter=False,
            merges=["A1:J1", "A2:J2"],
            style_func=_wallet_summary_style,
            show_gridlines=False,
        )
        for index, (wallet_name, account_number) in enumerate(AIRTEL_TIGO_WALLETS, start=2):
            detail_rows = _airtel_tigo_wallet_detail_rows(
                wallet_name,
                account_number,
                calculated[wallet_name],
            )
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{index}.xml",
                detail_rows,
                len(detail_rows),
                4,
                [25, 16, 20, 20],
                freeze_top_row=False,
                autofilter=False,
                merges=[
                    "A1:D1", "A2:D2", "A3:D3",
                    "A4:B4", "C4:D4",
                    *(f"A{row}:B{row}" for row in (5, 6, 7, 8, 9, 11, 12, 13, 14, 16)),
                    *(f"C{row}:D{row}" for row in (5, 6, 7, 8, 9, 11, 12, 13, 14, 16)),
                    "A18:B18", "C18:D18",
                ],
                style_func=_wallet_summary_style,
                show_gridlines=False,
            )
    os.replace(temp_path, output_path)
    return {
        "wallets": calculated,
        "wallet_count": len(calculated),
        "total_available_funds": sum((m["available_funds"] for m in calculated.values()), Decimal("0")),
        "total_debit": sum((m["total_debit"] for m in calculated.values()), Decimal("0")),
        "total_wallet_balance": sum((m["wallet_balance"] for m in calculated.values()), Decimal("0")),
    }
