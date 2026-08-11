#!/usr/bin/env python3
"""Zenith collection reconciliation workflow.

Zenith collection has no stable transaction identifier, so matching is done
only when the transaction date and amount match at the same time.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import zipfile
from collections import Counter, OrderedDict, defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable
from xml.etree.ElementTree import iterparse

from build_reconciliation_template import (
    CELL_REF_RE,
    FAST_XLSX_COMPRESSLEVEL,
    Cell,
    Record,
    STYLE_HEADER,
    STYLE_ITC_WRITE_OFF_BROWN,
    STYLE_MATCHED,
    STYLE_MONEY,
    STYLE_NOT_FOUND,
    STYLE_TABLE_COUNT,
    STYLE_TABLE_MONEY,
    STYLE_TABLE_SECTION,
    STYLE_TITLE,
    STYLE_UNIDENTIFIED,
    STYLE_WALLET_BALANCE_LABEL,
    STYLE_WALLET_BALANCE_MONEY,
    STYLE_WALLET_COMPANY,
    STYLE_WALLET_FINAL_LABEL,
    STYLE_WALLET_FINAL_MONEY,
    STYLE_WALLET_HEADER,
    STYLE_WALLET_MONEY,
    STYLE_WALLET_SIGNATURE,
    STYLE_WALLET_TITLE,
    add_unique_headers,
    amount_to_cents,
    app_xml,
    cell_value,
    cents_to_decimal,
    clean_sheet_value,
    col_index,
    column_letter,
    compute_widths,
    content_types_xml,
    core_xml,
    data_style,
    first_worksheet_path,
    iter_csv_dicts,
    iter_source_dicts,
    iter_xlsx_dicts,
    not_found_detail_rows,
    read_shared_strings,
    record_rows,
    root_rels_xml,
    set_csv_field_limit,
    styles_xml,
    summary_merges,
    summary_style,
    worksheet_path_by_name,
    workbook_rels_xml,
    workbook_xml,
    write_worksheet,
)


ZENITH_MATCHED_STATUS = "Matched"
ZENITH_NOT_FOUND_STATUS = "Not Found"
ZENITH_MAMBU_STATUS = "Mambu"
ZENITH_UNIDENTIFIED_STATUS = "Unidentified"
ZENITH_WRITE_OFF_STATUS = "Write Off"

ZENITH_MAMBU_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Account Holder ID",
    "Account ID",
    "Account Holder Name",
    "Amount (GHC)",
    "Channel",
    "Mobile Phone (Client)",
    "Identifier",
    "Gender (Client)",
    "Branch Name",
]

ZENITH_BANK_SHEET_COLUMNS: list[str] = [
    "Description",
    "Debit",
    "Credit",
    "Create Date (KEY date - DD/MM/YYYY)",
    "Effect Date",
    "Balance",
]

ZENITH_UNIDENTIFIED_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Identifier",
    "Amount (KEY)",
    "Type",
    "User",
    "Account ID",
    "Notes",
    "Product (Deposit)",
]

ZENITH_WRITE_OFF_SHEET_COLUMNS: list[str] = [
    "Date (KEY)",
    "Channel",
    "Amount (KEY)",
    "Fido Client Name",
    "Account Number",
    "Transaction ID",
    "Loan ID",
    "Status",
]

MAMBU_VS_ZENITH_HEADERS: list[str] = [
    "Mambu Date",
    "Mambu Account Holder",
    "Mambu Account ID",
    "Mambu Amount (GHC)",
    "Mambu Channel",
    "Match Status",
    "Zenith Description",
    "Zenith Date",
    "Zenith Amount (GHC)",
    "ZEN_KEY",
    "MAMBU_KEY",
    "ZEN_MATCH",
]

ZENITH_BANK_AS_SOURCE_HEADERS: list[str] = [
    "Zenith Date",
    "Zenith Description",
    "Zenith Amount (GHC)",
    "Match Status",
    "Matched Name / Identifier",
    "Matched Date",
    "Matched Amount (GHC)",
    "MAMBU_MATCH",
    "UNID_MATCH",
    "WRITEOFF_MATCH",
    "MAMBU_KEY",
    "UNID_KEY",
    "WRITEOFF_KEY",
]


@dataclass(frozen=True)
class ZenithSourceConfig:
    dataset: str
    sheet_name: str
    date_candidates: tuple[str, ...]
    amount_candidates: tuple[str, ...]
    skip_blank_key_rows: bool = False


@dataclass
class ZenithWalletLedgerResult:
    input_rows: int
    filtered_rows: int
    filter_counts: dict[str, int]
    filter_amounts: dict[str, Decimal]
    metrics: dict[str, Any] | None = None


def _normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())


def _pick_existing_key(row: dict[str, str], *candidates: str) -> str:
    normalized = {_normalize_header(key): key for key in row}
    for candidate in candidates:
        if candidate in row:
            return candidate
        key = normalized.get(_normalize_header(candidate))
        if key:
            return key

    candidate_parts = [tuple(part for part in re.split(r"\s+", candidate.casefold()) if part) for candidate in candidates]
    for key in row:
        lowered = key.casefold()
        if any(parts and all(part in lowered for part in parts) for parts in candidate_parts):
            return key
    return candidates[-1] if candidates else ""


def _pick(row: dict[str, str], *candidates: str) -> str:
    key = _pick_existing_key(row, *candidates)
    return row.get(key, "")


def normalize_zenith_date(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""

    try:
        number = float(text)
    except ValueError:
        number = None
    if number is not None and 20000 <= number <= 60000:
        converted = dt.datetime(1899, 12, 30) + dt.timedelta(days=number)
        return converted.strftime("%Y-%m-%d")

    cleaned = (
        text.replace("T", " ")
        .replace("\u00a0", " ")
        .strip()
        .rstrip("Z")
    )
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned_without_offset = re.sub(r"\s*[+-]\d{2}:?\d{2}$", "", cleaned)
    candidates = [cleaned_without_offset, cleaned_without_offset.split(" ", 1)[0]]
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y",
        "%d/%m/%y",
        "%d-%m-%y",
        "%m/%d/%Y",
        "%m-%d-%Y",
    )
    for candidate in candidates:
        for fmt in formats:
            try:
                return dt.datetime.strptime(candidate, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
    return cleaned_without_offset


def zenith_key(date_value: Any, amount_value: Any) -> tuple[str, int | None, str]:
    date_key = normalize_zenith_date(date_value)
    amount_cents = amount_to_cents(amount_value)
    if not date_key or amount_cents is None:
        return "", amount_cents, date_key
    return f"{date_key}|{amount_cents}", amount_cents, date_key


def _date_for_cleaning(header: str, value: Any) -> Any:
    if "date" in header.casefold():
        return normalize_zenith_date(value)
    return clean_sheet_value(header, value)


def _build_zenith_record(
    path: Path,
    source_row: int,
    row: dict[str, str],
    config: ZenithSourceConfig,
) -> Record | None:
    raw_date = _pick(row, *config.date_candidates)
    raw_amount = _pick(row, *config.amount_candidates)
    key, amount_cents, date_key = zenith_key(raw_date, raw_amount)

    if config.skip_blank_key_rows and (not date_key or amount_cents is None):
        return None

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = _date_for_cleaning(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset=config.dataset,
        direction="COLLECTION",
        key=key,
        amount_cents=amount_cents,
        datetime_value=date_key,
        phone="",
        channel_type=config.sheet_name,
        account="",
    )


def _load_zenith_records_from_iter(
    path: Path,
    rows: Iterable[tuple[int, dict[str, str]]],
    config: ZenithSourceConfig,
    fallback_headers: list[str],
) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    for source_row, row in rows:
        add_unique_headers(headers, row.keys())
        record = _build_zenith_record(path, source_row, row, config)
        if record is None:
            continue
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(fallback_headers)


def _source_rows(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    if path.suffix.lower() == ".csv":
        return iter_csv_dicts(path)
    return iter_xlsx_dicts(path, sheet_name)


def _amount_to_decimal(value: Any) -> Decimal:
    text = str(value or "").strip().replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _positive_amount_text(value: Any) -> Any:
    text = str(value or "").strip()
    if not text:
        return value
    amount = _amount_to_decimal(text)
    if amount == 0 and text.replace(",", "") not in {"0", "0.0", "0.00"}:
        return value
    return str(abs(amount))


def _iter_zenith_raw_rows(path: Path, sheet_name: str | None = None) -> Iterable[list[str]]:
    if path.suffix.lower() == ".csv":
        set_csv_field_limit()
        with path.open(newline="", encoding="utf-8-sig") as handle:
            yield from csv.reader(handle)
        return

    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        worksheet_path = worksheet_path_by_name(zf, sheet_name) if sheet_name else first_worksheet_path(zf)
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


def _clean_zenith_raw_rows(raw_rows: list[list[str]]) -> list[list[str]]:
    """Start at the Zenith transaction header for raw or already-cleaned files."""
    for index, raw_row in enumerate(raw_rows):
        headers = {_normalize_header(str(value)) for value in raw_row if str(value).strip()}
        has_description = any(header.startswith("description") for header in headers)
        has_date = any(
            header == "date" or "createdate" in header or "effectdate" in header
            for header in headers
        )
        has_amount = any(
            header == "amount" or header.startswith("credit") or header.startswith("paidin")
            for header in headers
        )
        if has_description and has_date and has_amount:
            return raw_rows[index:]
    return []


def load_cleaned_zenith_wallet_rows(
    paths: Iterable[Path],
    sheet_name: str | None = None,
) -> tuple[list[str], list[dict[str, str]]]:
    headers: list[str] = []
    rows: list[dict[str, str]] = []

    for path in paths:
        if not path or not path.exists():
            continue
        cleaned_raw = _clean_zenith_raw_rows(list(_iter_zenith_raw_rows(path, sheet_name)))
        if not cleaned_raw:
            continue
        source_headers = [str(value).strip() for value in cleaned_raw[0]]
        add_unique_headers(headers, source_headers)
        for row_number, raw_row in enumerate(cleaned_raw[1:], start=8):
            if not any(str(value).strip() for value in raw_row):
                continue
            if raw_row and "total" in str(raw_row[0]).strip().casefold():
                continue
            if [str(value).strip().casefold() for value in raw_row] == [header.casefold() for header in source_headers]:
                continue
            padded = raw_row + [""] * max(0, len(source_headers) - len(raw_row))
            row = {
                header: clean_sheet_value(header, padded[index] if index < len(padded) else "")
                for index, header in enumerate(source_headers)
            }
            row["_Source_Row"] = row_number
            rows.append(row)

    return headers, rows


def _cleaned_zenith_source_rows(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    _headers, rows = load_cleaned_zenith_wallet_rows([path], sheet_name)
    for row in rows:
        source_row = int(row.pop("_Source_Row", 0) or 0)
        yield source_row, row


def _positive_zenith_amount_rows(headers: list[str], rows: list[dict[str, str]]) -> list[dict[str, str]]:
    amount_indexes = {1, 2}
    normalized_rows: list[dict[str, str]] = []
    for row in rows:
        normalized = dict(row)
        for index, header in enumerate(headers):
            if index in amount_indexes:
                normalized[header] = _positive_amount_text(row.get(header, ""))
        normalized_rows.append(normalized)
    return normalized_rows


def _sum_positive_header(rows: Iterable[dict[str, str]], header: str) -> Decimal:
    return sum((abs(_amount_to_decimal(row.get(header, ""))) for row in rows), Decimal("0"))


def build_zenith_wallet_ledger_summary(
    balance_per_ledger: Decimal,
    delayed_transactions: Decimal,
    headers: list[str],
    rows: list[dict[str, str]],
) -> tuple[list[list[Any]], dict[str, Any]]:
    balance = abs(balance_per_ledger)
    debit_header = headers[1] if len(headers) > 1 else ""
    credit_header = headers[2] if len(headers) > 2 else ""

    debit_rows = [row for row in rows if _amount_to_decimal(row.get(debit_header, "")) != 0]
    collection_rows = [row for row in rows if _amount_to_decimal(row.get(credit_header, "")) != 0]
    total_collections = _sum_positive_header(rows, credit_header)
    delayed_credit = abs(delayed_transactions) if delayed_transactions > 0 else Decimal("0")
    delayed_debit = abs(delayed_transactions) if delayed_transactions < 0 else Decimal("0")
    available_funds = balance + total_collections + delayed_credit
    total_debit = _sum_positive_header(rows, debit_header) + delayed_debit
    wallet_statement_balance = abs(available_funds - total_debit)

    metrics = {
        "balance": balance,
        "delayed_transactions": abs(delayed_transactions),
        "delayed_credit": delayed_credit,
        "delayed_debit": delayed_debit,
        "total_collections": total_collections,
        "available_funds": available_funds,
        "total_debit": total_debit,
        "wallet_statement_balance": wallet_statement_balance,
        "input_rows": len(rows),
        "collections_count": len(collection_rows),
        "debit_count": len(debit_rows),
        "debit_header": debit_header,
        "credit_header": credit_header,
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

    summary_rows = [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell("ZENITH WALLET VS LEDGER", STYLE_WALLET_TITLE), "", "", ""],
        ["", "", "", ""],
        [wh("Line Item"), wh("Count"), wh("Amount (GHC)"), wh("Notes")],
        [wb("Balance Per Ledger"), "", wbm(balance), ""],
        [sig("Delayed Transactions (Credit)"), "", wm(delayed_credit), "Positive delayed transactions."],
        [sig("Total Collections"), len(collection_rows), wm(total_collections), "Sum of column C."],
        [wb("Available Funds Before Debit"), "", wbm(available_funds), "Balance Per Ledger + Total Collections + delayed credit."],
        ["", "", "", ""],
        [sig("Total Debit"), len(debit_rows), wm(total_debit), "Sum of column B + delayed debit."],
        [sig("Delayed Transactions (Debit)"), "", wm(delayed_debit), "Negative delayed transactions shown as debit."],
        ["", "", "", ""],
        [wf("Balance as per Wallet Statement"), "", wfm(wallet_statement_balance), "Absolute value of Available Funds Before Debit - Total Debit."],
    ]
    return summary_rows, metrics


def _wallet_summary_style(row_number: int, _col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    if row_number == 4:
        return STYLE_HEADER
    if isinstance(value, Decimal):
        return STYLE_MONEY
    return None


def _rows_from_cleaned_dicts(headers: list[str], rows: list[dict[str, str]]) -> Iterable[list[Any]]:
    yield headers
    for row in rows:
        yield [row.get(header, "") for header in headers]


def write_zenith_wallet_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = ["Summary", "Cleaned Zenith"]
    cleaned_rows = list(_rows_from_cleaned_dicts(headers, rows))

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
            [34, 14, 20, 70],
            freeze_top_row=False,
            autofilter=False,
            merges=[f"A1:{column_letter(4)}1", f"A2:{column_letter(4)}2"],
            style_func=_wallet_summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            cleaned_rows,
            len(cleaned_rows),
            len(headers),
            compute_widths(headers, ([row.get(header, "") for header in headers] for row in rows[:5000])),
            style_func=data_style(headers),
        )

    os.replace(temp_path, output_path)


def build_zenith_wallet_ledger_workbook(
    output_path: Path,
    source_paths: Iterable[Path],
    balance_per_ledger: Decimal,
    delayed_transactions: Decimal,
) -> ZenithWalletLedgerResult:
    headers, rows = load_cleaned_zenith_wallet_rows(source_paths)
    positive_rows = _positive_zenith_amount_rows(headers, rows)
    summary_rows, metrics = build_zenith_wallet_ledger_summary(
        balance_per_ledger,
        delayed_transactions,
        headers,
        positive_rows,
    )
    write_zenith_wallet_ledger_workbook(output_path, summary_rows, headers, positive_rows)
    return ZenithWalletLedgerResult(
        input_rows=len(positive_rows),
        filtered_rows=len(positive_rows),
        filter_counts={"Cleaned Zenith Wallet": len(positive_rows)},
        filter_amounts={"Cleaned Zenith Wallet": metrics.get("total_collections", Decimal("0"))},
        metrics=metrics,
    )


def load_zenith_collection_sources(
    mambu_path: Path | None,
    zenith_bank_path: Path | None,
    unidentified_path: Path | None,
    write_off_path: Path | None,
    mambu_sheet_name: str | None = None,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
]:
    mambu_config = ZenithSourceConfig(
        dataset="Mambu",
        sheet_name="Mambu",
        date_candidates=("Value Date (Entry Date)", "Date/Time", "Date"),
        amount_candidates=("Amount (GHC)", "Amount"),
    )
    zenith_config = ZenithSourceConfig(
        dataset="Zenith Bank",
        sheet_name="Zenith Bank",
        date_candidates=("Create Date (KEY date - DD/MM/YYYY)", "Create Date", "Date"),
        amount_candidates=("Credit", "Credit (KEY amount)", "Paid In", "Amount"),
        skip_blank_key_rows=True,
    )
    unidentified_config = ZenithSourceConfig(
        dataset="Unidentified",
        sheet_name="Unidentified",
        date_candidates=("Value Date (Entry Date)", "Date/Time", "Date"),
        amount_candidates=("Amount (KEY)", "Amount (GHC)", "Amount"),
        skip_blank_key_rows=True,
    )
    write_off_config = ZenithSourceConfig(
        dataset="Write Off",
        sheet_name="Write Off",
        date_candidates=("Date (KEY)", "Date/Time", "Date"),
        amount_candidates=("Amount (KEY)", "Amount (GHC)", "Amount"),
        skip_blank_key_rows=True,
    )

    mambu_records, mambu_headers = (
        _load_zenith_records_from_iter(
            mambu_path,
            _source_rows(mambu_path, mambu_sheet_name),
            mambu_config,
            ZENITH_MAMBU_SHEET_COLUMNS,
        )
        if mambu_path and mambu_path.exists()
        else ([], list(ZENITH_MAMBU_SHEET_COLUMNS))
    )
    zenith_records, zenith_headers = (
        _load_zenith_records_from_iter(
            zenith_bank_path,
            _cleaned_zenith_source_rows(zenith_bank_path),
            zenith_config,
            ZENITH_BANK_SHEET_COLUMNS,
        )
        if zenith_bank_path and zenith_bank_path.exists()
        else ([], list(ZENITH_BANK_SHEET_COLUMNS))
    )
    unidentified_records, unidentified_headers = (
        _load_zenith_records_from_iter(
            unidentified_path,
            _source_rows(unidentified_path),
            unidentified_config,
            ZENITH_UNIDENTIFIED_SHEET_COLUMNS,
        )
        if unidentified_path and unidentified_path.exists()
        else ([], list(ZENITH_UNIDENTIFIED_SHEET_COLUMNS))
    )
    write_off_records, write_off_headers = (
        _load_zenith_records_from_iter(
            write_off_path,
            _source_rows(write_off_path),
            write_off_config,
            ZENITH_WRITE_OFF_SHEET_COLUMNS,
        )
        if write_off_path and write_off_path.exists()
        else ([], list(ZENITH_WRITE_OFF_SHEET_COLUMNS))
    )

    return (
        mambu_records,
        mambu_headers,
        zenith_records,
        zenith_headers,
        unidentified_records,
        unidentified_headers,
        write_off_records,
        write_off_headers,
    )


def load_zenith_collection_workbook(path: Path) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
]:
    mambu_config = ZenithSourceConfig(
        dataset="Mambu",
        sheet_name="Mambu",
        date_candidates=("Value Date (Entry Date)", "Date/Time", "Date"),
        amount_candidates=("Amount (GHC)", "Amount"),
    )
    zenith_config = ZenithSourceConfig(
        dataset="Zenith Bank",
        sheet_name="Zenith Bank",
        date_candidates=("Create Date (KEY date - DD/MM/YYYY)", "Create Date", "Date"),
        amount_candidates=("Credit", "Credit (KEY amount)", "Paid In", "Amount"),
        skip_blank_key_rows=True,
    )
    unidentified_config = ZenithSourceConfig(
        dataset="Unidentified",
        sheet_name="Unidentified",
        date_candidates=("Value Date (Entry Date)", "Date/Time", "Date"),
        amount_candidates=("Amount (KEY)", "Amount (GHC)", "Amount"),
        skip_blank_key_rows=True,
    )
    write_off_config = ZenithSourceConfig(
        dataset="Write Off",
        sheet_name="Write Off",
        date_candidates=("Date (KEY)", "Date/Time", "Date"),
        amount_candidates=("Amount (KEY)", "Amount (GHC)", "Amount"),
        skip_blank_key_rows=True,
    )

    mambu_records, mambu_headers = _load_zenith_records_from_iter(
        path,
        iter_xlsx_dicts(path, "Mambu"),
        mambu_config,
        ZENITH_MAMBU_SHEET_COLUMNS,
    )
    zenith_records, zenith_headers = _load_zenith_records_from_iter(
        path,
        _cleaned_zenith_source_rows(path, "Zenith Bank"),
        zenith_config,
        ZENITH_BANK_SHEET_COLUMNS,
    )
    unidentified_records, unidentified_headers = _load_zenith_records_from_iter(
        path,
        iter_xlsx_dicts(path, "Unidentified"),
        unidentified_config,
        ZENITH_UNIDENTIFIED_SHEET_COLUMNS,
    )
    write_off_records, write_off_headers = _load_zenith_records_from_iter(
        path,
        iter_xlsx_dicts(path, "Write Off"),
        write_off_config,
        ZENITH_WRITE_OFF_SHEET_COLUMNS,
    )
    return (
        mambu_records,
        mambu_headers,
        zenith_records,
        zenith_headers,
        unidentified_records,
        unidentified_headers,
        write_off_records,
        write_off_headers,
    )


def _build_key_index(records: list[Record]) -> dict[str, list[Record]]:
    index: dict[str, list[Record]] = defaultdict(list)
    for record in records:
        if record.key:
            index[record.key].append(record)
    return index


def _first_match(record: Record, index: dict[str, list[Record]]) -> Record | None:
    if not record.key:
        return None
    candidates = index.get(record.key, [])
    return candidates[0] if candidates else None


def _first_value(record: Record | None, *headers: str) -> Any:
    if record is None:
        return ""
    for header in headers:
        value = record.sheet_values.get(header, "")
        if value not in ("", None):
            return value
    return ""


def _match_reason(source: Record, matched: Record | None, target_name: str) -> str:
    if matched is not None:
        return f"date and amount matched in {target_name}"
    if not source.key:
        return "blank date or amount"
    return f"date and amount not found in {target_name}"


def compare_mambu_to_zenith_bank(
    mambu_records: list[Record],
    zenith_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    zenith_index = _build_key_index(zenith_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mambu_records:
        match = _first_match(source, zenith_index)
        status = ZENITH_MATCHED_STATUS if match else ZENITH_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Mambu Date"] = source.datetime_value
        row["Mambu Account Holder"] = _first_value(source, "Account Holder Name", "Name")
        row["Mambu Account ID"] = _first_value(source, "Account ID")
        row["Mambu Amount (GHC)"] = source.amount_decimal
        row["Mambu Channel"] = _first_value(source, "Channel")
        row["Match Status"] = status
        row["Zenith Description"] = _first_value(match, "Description")
        row["Zenith Date"] = match.datetime_value if match else ""
        row["Zenith Amount (GHC)"] = match.amount_decimal if match else ""
        row["ZEN_KEY"] = match.key if match else ""
        row["MAMBU_KEY"] = source.key
        row["ZEN_MATCH"] = match.source_row if match else ""
        row["Match Reason"] = _match_reason(source, match, "Zenith Bank")
        rows.append(row)
    return rows


def compare_zenith_bank_to_sources(
    zenith_records: list[Record],
    mambu_records: list[Record],
    unidentified_records: list[Record],
    write_off_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mambu_index = _build_key_index(mambu_records)
    unidentified_index = _build_key_index(unidentified_records)
    write_off_index = _build_key_index(write_off_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in zenith_records:
        mambu_match = _first_match(source, mambu_index)
        unidentified_match = _first_match(source, unidentified_index)
        write_off_match = _first_match(source, write_off_index)

        status = ZENITH_NOT_FOUND_STATUS
        selected: Record | None = None
        if mambu_match is not None:
            status = ZENITH_MAMBU_STATUS
            selected = mambu_match
        elif unidentified_match is not None:
            status = ZENITH_UNIDENTIFIED_STATUS
            selected = unidentified_match
        elif write_off_match is not None:
            status = ZENITH_WRITE_OFF_STATUS
            selected = write_off_match

        row: OrderedDict[str, Any] = OrderedDict()
        row["Zenith Date"] = source.datetime_value
        row["Zenith Description"] = _first_value(source, "Description")
        row["Zenith Amount (GHC)"] = source.amount_decimal
        row["Match Status"] = status
        row["Matched Name / Identifier"] = _matched_label(status, selected)
        row["Matched Date"] = selected.datetime_value if selected else ""
        row["Matched Amount (GHC)"] = selected.amount_decimal if selected else ""
        row["MAMBU_MATCH"] = mambu_match.source_row if mambu_match else ""
        row["UNID_MATCH"] = unidentified_match.source_row if unidentified_match else ""
        row["WRITEOFF_MATCH"] = write_off_match.source_row if write_off_match else ""
        row["MAMBU_KEY"] = mambu_match.key if mambu_match else ""
        row["UNID_KEY"] = unidentified_match.key if unidentified_match else ""
        row["WRITEOFF_KEY"] = write_off_match.key if write_off_match else ""
        row["Match Reason"] = _zenith_source_reason(source, status)
        rows.append(row)
    return rows


def _matched_label(status: str, record: Record | None) -> str:
    if record is None:
        return ""
    if status == ZENITH_MAMBU_STATUS:
        return str(_first_value(record, "Account Holder Name", "Account ID", "Identifier"))
    if status == ZENITH_UNIDENTIFIED_STATUS:
        return str(_first_value(record, "Identifier", "Account ID", "Notes"))
    if status == ZENITH_WRITE_OFF_STATUS:
        return str(_first_value(record, "Fido Client Name", "Transaction ID", "Loan ID", "Account Number"))
    return ""


def _zenith_source_reason(source: Record, status: str) -> str:
    if not source.key:
        return "blank date or amount"
    if status == ZENITH_MAMBU_STATUS:
        return "date and amount matched in Mambu"
    if status == ZENITH_UNIDENTIFIED_STATUS:
        return "date and amount matched in Unidentified"
    if status == ZENITH_WRITE_OFF_STATUS:
        return "date and amount matched in Write Off"
    return "date and amount not found in Mambu, Unidentified, or Write Off"


def summarize_zenith_compare(
    rows: list[OrderedDict[str, Any]],
    amount_header: str,
    matched_statuses: set[str],
) -> dict[str, Any]:
    status_counts = Counter(str(row.get("Match Status", "")) for row in rows)
    status_amount_cents: defaultdict[str, int] = defaultdict(int)
    source_amount_cents = 0
    matched_amount_cents = 0

    for row in rows:
        status = str(row.get("Match Status", ""))
        cents = amount_to_cents(row.get(amount_header, "")) or 0
        source_amount_cents += cents
        status_amount_cents[status] += cents
        if status in matched_statuses:
            matched_amount_cents += cents

    total = len(rows)
    matched = sum(status_counts[status] for status in matched_statuses)
    return {
        "total": total,
        "matched": matched,
        "not_found": status_counts[ZENITH_NOT_FOUND_STATUS],
        "match_rate": matched / total if total else 0,
        "source_amount": cents_to_decimal(source_amount_cents),
        "matched_amount": cents_to_decimal(matched_amount_cents),
        "not_found_amount": cents_to_decimal(status_amount_cents[ZENITH_NOT_FOUND_STATUS]),
        "status_counts": dict(status_counts),
        "status_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_amount_cents.items()
        },
    }


def sum_zenith_amounts(records: list[Record]) -> Decimal:
    return cents_to_decimal(sum(record.amount_cents or 0 for record in records))  # type: ignore[return-value]


def _count(stats: dict[str, Any], status: str) -> int:
    return int(stats["status_counts"].get(status, 0))


def _amount(stats: dict[str, Any], status: str) -> Decimal | str:
    return stats["status_amounts"].get(status, Decimal("0"))


def build_zenith_summary_rows(
    mambu_vs_zenith_stats: dict[str, Any],
    zenith_vs_sources_stats: dict[str, Any],
) -> list[list[Any]]:
    mvz = mambu_vs_zenith_stats
    zvs = zenith_vs_sources_stats

    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    return [
        [Cell("ZENITH COLLECTION RECONCILIATION - SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu vs Zenith Bank"), th("Count"), th("Amount (GHC)")],
        [Cell(ZENITH_MATCHED_STATUS, STYLE_MATCHED), _count(mvz, ZENITH_MATCHED_STATUS), Cell(_amount(mvz, ZENITH_MATCHED_STATUS), STYLE_MONEY)],
        [Cell(ZENITH_NOT_FOUND_STATUS, STYLE_NOT_FOUND), _count(mvz, ZENITH_NOT_FOUND_STATUS), Cell(_amount(mvz, ZENITH_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mvz["total"]), tm(mvz["source_amount"])],
        ["", "", ""],
        [th("Zenith Bank as Source"), th("Count"), th("Amount (GHC)")],
        [Cell(ZENITH_MAMBU_STATUS, STYLE_MATCHED), _count(zvs, ZENITH_MAMBU_STATUS), Cell(_amount(zvs, ZENITH_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(ZENITH_UNIDENTIFIED_STATUS, STYLE_UNIDENTIFIED), _count(zvs, ZENITH_UNIDENTIFIED_STATUS), Cell(_amount(zvs, ZENITH_UNIDENTIFIED_STATUS), STYLE_MONEY)],
        [Cell(ZENITH_WRITE_OFF_STATUS, STYLE_ITC_WRITE_OFF_BROWN), _count(zvs, ZENITH_WRITE_OFF_STATUS), Cell(_amount(zvs, ZENITH_WRITE_OFF_STATUS), STYLE_MONEY)],
        [Cell(ZENITH_NOT_FOUND_STATUS, STYLE_NOT_FOUND), _count(zvs, ZENITH_NOT_FOUND_STATUS), Cell(_amount(zvs, ZENITH_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(zvs["total"]), tm(zvs["source_amount"])],
    ]


def build_zenith_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("ZENITH COLLECTION RECONCILIATION TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        ["Purpose", "Reconcile Zenith Bank collection transactions against Mambu."],
        ["Required data", "Mambu and Zenith Bank are required. Unidentified and Write Off are optional."],
        ["Mambu vs Zenith Bank", "Mambu column A date and column E amount are matched to Zenith Bank column D date and column C credit amount."],
        ["Zenith Bank as Source", "Zenith Bank column D date and column C credit amount are matched to Mambu first, then Unidentified, then Write Off."],
        ["Match rule", "A transaction only matches when both date and amount match at the same time."],
        ["Zenith note", "Opening balance or non-credit rows with blank date or credit amount are ignored."],
    ]


def _rows_from_dicts(headers: list[str], rows: list[OrderedDict[str, Any]]) -> Iterable[list[Any]]:
    yield headers
    for row in rows:
        yield [row.get(header, "") for header in headers]


def _zenith_compare_style(headers: list[str]):
    amount_headers = {
        "Mambu Amount (GHC)",
        "Zenith Amount (GHC)",
        "Matched Amount (GHC)",
        "Amount (GHC)",
        "Amount (KEY)",
        "Credit (KEY amount)",
        "Debit",
        "Balance",
    }

    def style(row_number: int, col_number: int, value: Any) -> int | None:
        if row_number == 1:
            return STYLE_HEADER
        header = headers[col_number - 1] if col_number - 1 < len(headers) else ""
        if header == "Match Status":
            if value in {ZENITH_MATCHED_STATUS, ZENITH_MAMBU_STATUS}:
                return STYLE_MATCHED
            if value == ZENITH_UNIDENTIFIED_STATUS:
                return STYLE_UNIDENTIFIED
            if value == ZENITH_WRITE_OFF_STATUS:
                return STYLE_ITC_WRITE_OFF_BROWN
            return STYLE_NOT_FOUND
        if header in amount_headers or "amount" in header.casefold() or header.casefold() in {"debit", "credit", "balance"}:
            return STYLE_MONEY
        return None

    return style


def write_zenith_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    zenith_headers: list[str],
    zenith_records: list[Record],
    unidentified_headers: list[str],
    unidentified_records: list[Record],
    write_off_headers: list[str],
    write_off_records: list[Record],
    mambu_vs_zenith_rows: list[OrderedDict[str, Any]],
    zenith_vs_sources_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Mambu vs Zenith Bank",
        "Zenith Bank as Source",
        "Mambu",
        "Zenith Bank",
        "Unidentified",
        "Write Off",
    ]
    instruction_rows = build_zenith_instruction_rows()

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    zenith_preview = ([record.sheet_values.get(header, "") for header in zenith_headers] for record in zenith_records[:5000])
    unidentified_preview = ([record.sheet_values.get(header, "") for header in unidentified_headers] for record in unidentified_records[:5000])
    write_off_preview = ([record.sheet_values.get(header, "") for header in write_off_headers] for record in write_off_records[:5000])
    mambu_vs_preview = ([row.get(header, "") for header in MAMBU_VS_ZENITH_HEADERS] for row in mambu_vs_zenith_rows[:5000])
    zenith_vs_preview = ([row.get(header, "") for header in ZENITH_BANK_AS_SOURCE_HEADERS] for row in zenith_vs_sources_rows[:5000])

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
            instruction_rows,
            len(instruction_rows),
            2,
            [36, 110],
            freeze_top_row=False,
            autofilter=False,
            merges=["A1:B1"],
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            summary_rows,
            len(summary_rows),
            3,
            [34, 14, 22],
            freeze_top_row=False,
            autofilter=False,
            merges=summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            _rows_from_dicts(MAMBU_VS_ZENITH_HEADERS, mambu_vs_zenith_rows),
            len(mambu_vs_zenith_rows) + 1,
            len(MAMBU_VS_ZENITH_HEADERS),
            compute_widths(MAMBU_VS_ZENITH_HEADERS, mambu_vs_preview),
            style_func=_zenith_compare_style(MAMBU_VS_ZENITH_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            _rows_from_dicts(ZENITH_BANK_AS_SOURCE_HEADERS, zenith_vs_sources_rows),
            len(zenith_vs_sources_rows) + 1,
            len(ZENITH_BANK_AS_SOURCE_HEADERS),
            compute_widths(ZENITH_BANK_AS_SOURCE_HEADERS, zenith_vs_preview),
            style_func=_zenith_compare_style(ZENITH_BANK_AS_SOURCE_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=_zenith_compare_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(zenith_headers, zenith_records),
            len(zenith_records) + 1,
            len(zenith_headers),
            compute_widths(zenith_headers, zenith_preview),
            style_func=_zenith_compare_style(zenith_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            record_rows(unidentified_headers, unidentified_records),
            len(unidentified_records) + 1,
            len(unidentified_headers),
            compute_widths(unidentified_headers, unidentified_preview),
            style_func=_zenith_compare_style(unidentified_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet8.xml",
            record_rows(write_off_headers, write_off_records),
            len(write_off_records) + 1,
            len(write_off_headers),
            compute_widths(write_off_headers, write_off_preview),
            style_func=_zenith_compare_style(write_off_headers),
        )

    os.replace(temp_path, output_path)


def build_zenith_collection_reconciliation(
    output_path: Path,
    mambu_records: list[Record],
    mambu_headers: list[str],
    zenith_records: list[Record],
    zenith_headers: list[str],
    unidentified_records: list[Record],
    unidentified_headers: list[str],
    write_off_records: list[Record],
    write_off_headers: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    mambu_vs_zenith_rows = compare_mambu_to_zenith_bank(mambu_records, zenith_records)
    zenith_vs_sources_rows = compare_zenith_bank_to_sources(
        zenith_records,
        mambu_records,
        unidentified_records,
        write_off_records,
    )
    mambu_vs_zenith_stats = summarize_zenith_compare(
        mambu_vs_zenith_rows,
        "Mambu Amount (GHC)",
        {ZENITH_MATCHED_STATUS},
    )
    zenith_vs_sources_stats = summarize_zenith_compare(
        zenith_vs_sources_rows,
        "Zenith Amount (GHC)",
        {ZENITH_MAMBU_STATUS, ZENITH_UNIDENTIFIED_STATUS, ZENITH_WRITE_OFF_STATUS},
    )
    mambu_vs_zenith_stats["not_found_details"] = not_found_detail_rows(
        mambu_records,
        mambu_vs_zenith_rows,
        ZENITH_NOT_FOUND_STATUS,
        "Mambu",
        "Mambu → Wallet",
    )
    zenith_vs_sources_stats["not_found_details"] = not_found_detail_rows(
        zenith_records,
        zenith_vs_sources_rows,
        ZENITH_NOT_FOUND_STATUS,
        "Zenith",
        "Wallet → Mambu",
    )
    summary_rows = build_zenith_summary_rows(mambu_vs_zenith_stats, zenith_vs_sources_stats)
    write_zenith_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        zenith_headers,
        zenith_records,
        unidentified_headers,
        unidentified_records,
        write_off_headers,
        write_off_records,
        mambu_vs_zenith_rows,
        zenith_vs_sources_rows,
    )
    return mambu_vs_zenith_stats, zenith_vs_sources_stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a Zenith collection reconciliation workbook.")
    parser.add_argument("--source", default=None, help="Workbook containing Mambu, Zenith Bank, Unidentified, and Write Off sheets.")
    parser.add_argument("--mambu", default=None, help="Mambu source xlsx/csv path.")
    parser.add_argument("--zenith-bank", default=None, help="Zenith Bank source xlsx/csv path.")
    parser.add_argument("--unidentified", default=None, help="Unidentified source xlsx/csv path.")
    parser.add_argument("--write-off", default=None, help="Write Off source xlsx/csv path.")
    parser.add_argument("--output", default="Zenith Coll Recon.xlsx", help="Output workbook path.")
    args = parser.parse_args()

    if args.source:
        records = load_zenith_collection_workbook(Path(args.source))
    else:
        if not args.mambu or not args.zenith_bank:
            raise SystemExit("Provide either --source or both --mambu and --zenith-bank.")
        records = load_zenith_collection_sources(
            Path(args.mambu),
            Path(args.zenith_bank),
            Path(args.unidentified) if args.unidentified else None,
            Path(args.write_off) if args.write_off else None,
        )

    (
        mambu_records,
        mambu_headers,
        zenith_records,
        zenith_headers,
        unidentified_records,
        unidentified_headers,
        write_off_records,
        write_off_headers,
    ) = records
    mvz, zvs = build_zenith_collection_reconciliation(
        Path(args.output),
        mambu_records,
        mambu_headers,
        zenith_records,
        zenith_headers,
        unidentified_records,
        unidentified_headers,
        write_off_records,
        write_off_headers,
    )
    print(
        f"Mambu vs Zenith Bank: {mvz['matched']:,}/{mvz['total']:,} matched; "
        f"Zenith Bank as Source: {zvs['matched']:,}/{zvs['total']:,} matched"
    )


if __name__ == "__main__":
    main()
