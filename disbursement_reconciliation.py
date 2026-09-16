#!/usr/bin/env python3
"""Nsano disbursement reconciliation workflow.

This module is intentionally separate from the collection reconciliation logic.
Disbursement matching is key-only. Identifier columns are resolved by header
name first, with legacy column positions as a fallback, and Nsano's leading
FIDO prefix is ignored during matching.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import zipfile
from collections import Counter, OrderedDict, defaultdict
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable
from xml.etree.ElementTree import iterparse

from build_reconciliation_template import (
    CELL_REF_RE,
    FAST_XLSX_COMPRESSLEVEL,
    Cell,
    Record,
    STYLE_HEADER,
    STYLE_MATCHED,
    STYLE_MONEY,
    STYLE_NOT_FOUND,
    STYLE_TABLE_COUNT,
    STYLE_TABLE_MONEY,
    STYLE_TABLE_SECTION,
    STYLE_TITLE,
    STYLE_WALLET_BALANCE_LABEL,
    STYLE_WALLET_BALANCE_MONEY,
    STYLE_WALLET_COMPANY,
    STYLE_WALLET_FINAL_LABEL,
    STYLE_WALLET_FINAL_MONEY,
    STYLE_WALLET_HEADER,
    STYLE_WALLET_LINE_ITEM,
    STYLE_WALLET_MONEY,
    STYLE_WALLET_SECTION,
    STYLE_WALLET_SIGNATURE,
    STYLE_WALLET_TITLE,
    add_unique_headers,
    amount_to_cents,
    app_xml,
    clean_sheet_value,
    compute_widths,
    content_types_xml,
    core_xml,
    data_style,
    cell_value,
    col_index,
    first_worksheet_path,
    iter_csv_dicts,
    iter_xlsx_dicts,
    not_found_detail_rows,
    parse_nsano_datetime,
    record_rows,
    root_rels_xml,
    set_csv_field_limit,
    styles_xml,
    summary_merges,
    summary_style,
    read_shared_strings,
    workbook_rels_xml,
    workbook_xml,
    worksheet_path_by_name,
    write_worksheet,
    xlsx_sheet_if_present,
)


DISB_MAMBU_STATUS = "mambu"
DISB_NSANO_STATUS = "nsano"
DISB_NOT_FOUND_STATUS = "not found"
ITC_DISB_STATUS = "ITC"
ITC_DISB_NOT_FOUND_STATUS = "not_found"
ITC_DISB_REFERRAL_BONUS_STATUS = "referral bonus"
ITC_DISB_UPSALES_REFOUND_STATUS = "upsales refund"
ITC_DISB_SAVINGS_REWARD_STATUS = "savings reward"
MTN_MANUAL_MATCHED_STATUS = "Matched"
MTN_MANUAL_MAMBU_STATUS = "Mambu"
MTN_MANUAL_REFUND_STATUS = "Refund"
MTN_MANUAL_NOT_FOUND_STATUS = "Not Found"
VODAFONE_MANUAL_MATCHED_STATUS = "Matched"
VODAFONE_MANUAL_MAMBU_STATUS = "Mambu"
VODAFONE_MANUAL_REFUND_STATUS = "Refund"
VODAFONE_MANUAL_NOT_FOUND_STATUS = "Not Found"
ITC_WALLET_SETTLEMENT_STATUS = "settlement"
ITC_WALLET_REVERSAL_STATUS = "reversal"
ITC_WALLET_PREPAID_REVERSAL_STATUS = "prepaid_reversal"
ITC_WALLET_TRANSFER_TO_WALLET_STATUS = "transfer_to_wallet"
ITC_WALLET_DISBURSEMENT_STATUS = "disbursement"
ITC_WALLET_REFERRAL_AWARD_STATUS = "referral_award"
ITC_WALLET_UPSALES_REFUND_STATUS = "upsales_refund"
ITC_WALLET_SAVINGS_STATUS = "savings"
ITC_WALLET_TRANSFERS_FROM_BANK_STATUS = "transfers_from_bank"
ITC_WALLET_CREDIT_TRANSFER_STATUS = "credit transfer (prepaid_reversal)"
ITC_WALLET_NOT_FOUND_STATUS = "not_found"
NSANO_WALLET_TOPUP_COLLECTION_STATUS = "top_up_through_collections"
NSANO_WALLET_BANK_TO_WALLET_STATUS = "bank_to_wallet"
NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS = "reversal_adjustment"
NSANO_WALLET_TRANSFER_TO_BANK_STATUS = "transfer_to_bank"
NSANO_WALLET_OTHER_STATUS = "other"
NSANO_MAMBU_COLLECTION_CHANNELS = {
    "nsanoclientcollections",
    "nsanoclientcollection",
}
NSANO_MAMBU_DISBURSEMENT_CHANNELS = {
    "nsano-clientdisbursement",
    "nsanoclientdisb",
}

NSANO_DISB_MAMBU_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Account Holder ID",
    "Account Holder Name",
    "Principal Amount",
    "Amount (GHC)",
    "Loan Number",
    "Channel",
    "Mobile Phone (Client)",
    "Identifier",
    "Loan Purpose",
    "Gender (Client)",
    "Birth Date (Client)",
    "Branch Name",
]

NSANO_COLLECTION_MAMBU_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Account Holder ID",
    "Account ID",
    "Account Holder Name",
    "Amount",
    "Channel",
    "Mobile Phone (Client)",
    "Identifier",
    "Gender (Client)",
    "Branch Name",
]

NSANO_DISB_SHEET_COLUMNS: list[str] = [
    "DateTime",
    "Transaction_ID",
    "Type",
    "Author",
    "Amount_GHC",
    "balanceBefore",
    "balanceAfter",
    "SendingHse",
    "SendingHse_ID (KEY)",
    "SendingHse_Account",
    "External_Debit_Reference",
    "SendingHse_Code",
    "SendingHse_Message",
    "ReceivingHse",
    "ReceivingHse_Account",
    "ReceivingHse_ID",
    "External_Credit_Reference",
    "ReceivingHse_Code",
    "ReceivingHse_Message",
    "Result",
    "Charge",
    "Deferred",
    "Service_Label",
    "Service_Transaction_ID",
    "Service_Details",
]

NSANO_WALLET_TRANSFER_SHEET_COLUMNS: list[str] = [
    "Date",
    "VA ID",
    "Map Name",
    "Amount",
    "Purpose",
    "Old Balance",
    "New Balance",
]

ITC_DISB_MAMBU_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Account Holder ID",
    "Account Holder Name",
    "Principal Amount",
    "Amount (GHC)",
    "Loan Number",
    "Channel",
    "Mobile Phone (Client)",
    "Identifier (KEY)",
    "Loan Purpose",
    "Gender (Client)",
    "Birth Date (Client)",
    "Branch Name",
]

ITC_DISB_SHEET_COLUMNS: list[str] = [
    "processor_transaction_id",
    "channel",
    "channel_transaction_id",
    "account_reference",
    "account_name",
    "payer_contact",
    "thirdparty_id (KEY)",
    "processor",
    "transaction_date",
    "transaction_time",
    "transaction_type",
    "source",
    "narration",
    "currency",
    "amount",
    "net_amount",
    "fees",
    "elevy_charge",
    "country_code",
    "created",
    "merchant_product_name",
]

ITC_WALLET_STATEMENT_SHEET_COLUMNS: list[str] = [
    "processor_transaction_id",
    "channel",
    "channel_transaction_id",
    "account_reference",
    "account_name",
    "payer_contact",
    "thirdparty_id",
    "processor",
    "transaction_date",
    "transaction_time",
    "transaction_type",
    "source",
    "narration",
    "currency",
    "amount",
]

ITC_WALLET_DEBIT_TRANSFER_SHEET_COLUMNS: list[str] = [
    "prepaid_transaction_id",
    "merchant_id",
    "third_party_transaction_id",
    "transaction_date",
    "prepaid_transaction_date",
    "amount",
    "narration",
    "created_at",
]

ITC_WALLET_CREDIT_TRANSFER_SHEET_COLUMNS: list[str] = [
    "prepaid_transaction_id",
    "merchant_id",
    "third_party_transaction_id",
    "transaction_date",
    "prepaid_transaction_date",
    "amount",
    "narration",
    "created_at",
]

MTN_MANUAL_MAMBU_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Account Holder ID",
    "Account Holder Name",
    "Principal Amount",
    "Amount",
    "Loan Number",
    "Channel",
    "Mobile Phone (Client)",
    "Identifier",
    "Why do you need a loan?",
    "Gender (Client)",
    "Birth Date (Client)",
    "Branch Name",
]

MTN_MANUAL_SHEET_COLUMNS: list[str] = [
    "Id",
    "External id",
    "Date",
    "Status",
    "Type",
    "Amount",
]

MTN_MANUAL_REFUND_SHEET_COLUMNS: list[str] = [
    "DATE",
    "CHANNEL",
    "NUMBER",
    "TRANSACTION ID",
    "AMOUNT",
    "Comments",
]

VODAFONE_MANUAL_MAMBU_SHEET_COLUMNS: list[str] = list(MTN_MANUAL_MAMBU_SHEET_COLUMNS)

VODAFONE_MANUAL_SHEET_COLUMNS: list[str] = [
    "Id",
    "External id",
    "Date",
    "Status",
    "Type",
    "Amount",
]
VODAFONE_MANUAL_HEADER_SCAN_LIMIT = 25

VODAFONE_MANUAL_REFUND_SHEET_COLUMNS: list[str] = list(MTN_MANUAL_REFUND_SHEET_COLUMNS)

MAMBU_VS_NSANO_HEADERS: list[str] = [
    "Mambu Identifier",
    "Mambu Date",
    "Mambu Account Holder",
    "Mambu Amount (GHC)",
    "Mambu Loan No.",
    "Mambu Mobile",
    "Match Status",
    "Nsano Trans ID",
    "Nsano DateTime",
    "Nsano Amount (GHC)",
    "Nsano Row",
]

NSANO_VS_MAMBU_HEADERS: list[str] = [
    "Nsano Identifier (stripped)",
    "Nsano DateTime",
    "Nsano Trans ID",
    "Nsano Amount (GHC)",
    "Nsano Result",
    "Nsano Receiving Acct",
    "Match Status",
    "Mambu Identifier",
    "Mambu Date",
    "Mambu Amount (GHC)",
    "Mambu Loan No.",
    "Mambu Row",
]

MAMBU_VS_ITC_DISB_HEADERS: list[str] = [
    "Mambu Identifier",
    "Mambu Date",
    "Mambu Account Holder",
    "Mambu Amount (GHC)",
    "Mambu Loan No.",
    "Mambu Mobile",
    "Match Status",
    "ITC Trans ID",
    "ITC Date",
    "ITC Amount (GHC)",
    "ITC Channel",
    "ITC Row",
]

ITC_DISB_VS_MAMBU_HEADERS: list[str] = [
    "ITC thirdparty_id (Key)",
    "ITC Date",
    "ITC Processor Trans ID",
    "ITC Amount (GHC)",
    "ITC Channel",
    "ITC Source",
    "Match Status",
    "Mambu Identifier",
    "Mambu Date",
    "Mambu Amount (GHC)",
    "Mambu Account Holder",
    "Narration",
    "Mambu Row",
]

MAMBU_VS_MTN_MANUAL_HEADERS: list[str] = [
    "Value Date",
    "Account Holder ID",
    "Account Holder Name",
    "Principal Amount",
    "Amount",
    "Loan Number",
    "Channel",
    "Mobile Phone",
    "Identifier",
    "Why Loan?",
    "Gender",
    "Birth Date",
    "Branch",
    "Match Status",
    "MTN_MATCH",
]

MTN_MANUAL_AS_SOURCE_HEADERS: list[str] = [
    "Id",
    "External Id",
    "Date",
    "Status",
    "Type",
    "Amount",
    "Match Status",
    "MAMBU_MATCH",
    "REFUND_MATCH",
]

MAMBU_VS_VODAFONE_MANUAL_HEADERS: list[str] = [
    "Value Date",
    "Account Holder ID",
    "Account Holder Name",
    "Principal Amount",
    "Amount",
    "Loan Number",
    "Channel",
    "Mobile Phone",
    "Identifier",
    "Why Loan?",
    "Gender",
    "Birth Date",
    "Branch",
    "Match Status",
    "VF_MATCH",
]

VODAFONE_MANUAL_AS_SOURCE_HEADERS: list[str] = list(MTN_MANUAL_AS_SOURCE_HEADERS)

ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS: list[str] = [
    "Credit Identifier",
    "Credit Date",
    "Credit Amount (GHC)",
    "Credit Narration",
    "Credit Category",
    "Settlement",
    "Reversal",
    "Prepaid Reversal",
    "Transfer to Wallet",
    "Source Row",
]

ITC_WALLET_DEBIT_COMPARISON_HEADERS: list[str] = [
    "Debit Identifier",
    "Debit Date",
    "Debit Amount (GHC)",
    "Debit Narration",
    "Match Status",
    "Statement Identifier",
    "Statement Narration",
    "Statement Amount (GHC)",
    "Statement Row",
    "Credit Transfer Identifier",
    "Credit Transfer Row",
]

NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS: list[str] = [
    "Transfer Date",
    "Map Name",
    "Purpose",
    "Transfer Amount (GHC)",
    "Transfer Category",
    "Narration Exception",
    "Exception Reason",
    "TOP-UP Through Collection with Date",
    "Bank to Wallet",
    "Reversal Adjustment",
    "Transfer to Bank",
    "Source Row",
]


def normalize_disbursement_identifier(value: Any) -> str:
    text = "".join(str(value or "").split()).strip("'\"").upper()
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?(?:E[+-]?\d+)?", text):
        try:
            number = Decimal(text)
            if number == number.to_integral_value():
                text = str(number.quantize(Decimal("1")))
        except InvalidOperation:
            pass
    if text.startswith("FIDO"):
        return text[4:]
    return text


def _normalized_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _named_row_value(row: dict[str, Any], *headers: str) -> Any:
    normalized = {_normalized_header(key): value for key, value in row.items()}
    for header in headers:
        value = normalized.get(_normalized_header(header), "")
        if value not in ("", None):
            return value
    return ""


def row_value_at(row: dict[str, Any], column_index: int, *fallback_headers: str) -> Any:
    named_value = _named_row_value(row, *fallback_headers)
    if named_value not in ("", None):
        return named_value
    values = list(row.values())
    if len(values) >= column_index and values[column_index - 1] not in ("", None):
        return values[column_index - 1]
    return ""


def pick(row: dict[str, Any], *headers: str) -> Any:
    return _named_row_value(row, *headers)


def pick_key(row: dict[str, Any], *headers: str) -> str:
    normalized = {_normalized_header(key): key for key in row}
    for header in headers:
        if header in row:
            return header
        key = normalized.get(_normalized_header(header))
        if key:
            return key
    return headers[-1]


def _compact_source_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _load_records_from_candidate_sheets(
    path: Path,
    sheet_candidates: tuple[str, ...],
    loader,
) -> tuple[list[Record], list[str]]:
    if path.suffix.lower() == ".csv":
        return loader(path, None)

    records: list[Record] = []
    headers: list[str] = []
    loaded_candidate = False
    for sheet_name in sheet_candidates:
        try:
            sheet_records, sheet_headers = loader(path, sheet_name)
        except KeyError:
            continue
        loaded_candidate = True
        records.extend(sheet_records)
        add_unique_headers(headers, sheet_headers)

    if loaded_candidate:
        return records, headers
    return loader(path, None)


def _records_matching(records: list[Record], matcher, *, fallback_to_all: bool = True) -> list[Record]:
    matched = [record for record in records if matcher(record)]
    return matched if matched or not fallback_to_all else records


def _is_nsano_mambu_collection_record(record: Record) -> bool:
    channel = _compact_source_text(record.channel_type)
    return channel in NSANO_MAMBU_COLLECTION_CHANNELS or (
        "nsano" in channel and "coll" in channel
    )


def _is_nsano_mambu_disbursement_record(record: Record) -> bool:
    channel = _compact_source_text(record.channel_type)
    return channel in NSANO_MAMBU_DISBURSEMENT_CHANNELS or (
        "nsano" in channel and "disb" in channel
    )


def _is_nsano_disbursement_record(record: Record) -> bool:
    return _compact_source_text(record.channel_type) == "a2w"


def _source_cache_key(path: Path, sheet_name: str | None = None) -> tuple[str, int, int, str | None]:
    stat = path.stat()
    return (str(path.resolve()), stat.st_size, stat.st_mtime_ns, sheet_name)


@lru_cache(maxsize=64)
def _cached_source_rows(
    path_text: str,
    size_bytes: int,
    modified_ns: int,
    sheet_name: str | None,
) -> tuple[tuple[int, tuple[tuple[str, str], ...]], ...]:
    path = Path(path_text)
    rows = iter_csv_dicts(path) if path.suffix.lower() == ".csv" else iter_xlsx_dicts(path, sheet_name)
    return tuple((row_number, tuple(row.items())) for row_number, row in rows)


def iter_source_rows(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    if path.suffix.lower() == ".csv":
        sheet_name = None
    for row_number, items in _cached_source_rows(*_source_cache_key(path, sheet_name)):
        yield row_number, dict(items)


def unique_headers(headers: list[str]) -> list[str]:
    counts: Counter[str] = Counter()
    unique: list[str] = []
    for idx, header in enumerate(headers, start=1):
        name = header or f"Column {idx}"
        counts[name] += 1
        unique.append(name if counts[name] == 1 else f"{name} ({counts[name]})")
    return unique


def _iter_raw_source_rows(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, list[str]]]:
    if path.suffix.lower() == ".csv":
        set_csv_field_limit()
        with path.open(newline="", encoding="utf-8-sig") as handle:
            for row_number, row in enumerate(csv.reader(handle), start=1):
                yield row_number, row
        return

    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        worksheet_path = worksheet_path_by_name(zf, sheet_name) if sheet_name else first_worksheet_path(zf)
        with zf.open(worksheet_path) as handle:
            for _event, elem in iterparse(handle, events=("end",)):
                if not elem.tag.endswith("}row"):
                    continue

                row_number = int(elem.attrib.get("r", "0") or 0)
                values_by_index: dict[int, str] = {}
                max_index = 0
                for cell in list(elem):
                    if not cell.tag.endswith("}c"):
                        continue
                    ref_match = CELL_REF_RE.match(cell.attrib.get("r", ""))
                    index = col_index(ref_match.group(1)) if ref_match else len(values_by_index) + 1
                    values_by_index[index] = cell_value(cell, shared_strings)
                    max_index = max(max_index, index)

                yield row_number, [values_by_index.get(i, "") for i in range(1, max_index + 1)]
                elem.clear()


def _is_vodafone_manual_disb_header(raw_row: list[str]) -> bool:
    normalized = {_normalized_header(value) for value in raw_row if str(value).strip()}
    has_amount = bool({"amount", "amountghc", "amountghs", "withdrawn", "paidin"} & normalized)
    return "receiptno" in normalized or ("id" in normalized and "date" in normalized and has_amount)


def _vodafone_manual_disb_rows(
    path: Path,
    sheet_name: str | None = None,
) -> tuple[list[str], list[tuple[int, dict[str, str]]]]:
    source_headers: list[str] | None = None
    source_rows: list[tuple[int, dict[str, str]]] = []

    for scan_index, (row_number, raw_row) in enumerate(_iter_raw_source_rows(path, sheet_name)):
        if source_headers is None:
            if not _is_vodafone_manual_disb_header(raw_row):
                if scan_index + 1 >= VODAFONE_MANUAL_HEADER_SCAN_LIMIT:
                    break
                continue
            source_headers = unique_headers([str(value).strip() for value in raw_row])
            continue

        if not any(str(value).strip() for value in raw_row):
            continue
        if _is_vodafone_manual_disb_header(raw_row):
            continue

        padded = list(raw_row) + [""] * max(0, len(source_headers) - len(raw_row))
        source_rows.append((row_number, dict(zip(source_headers, padded))))

    return source_headers or [], source_rows


def detect_vodafone_manual_disb_headers(path: Path, sheet_name: str | None = None) -> list[str]:
    headers, _rows = _vodafone_manual_disb_rows(path, sheet_name)
    return headers


def iter_xlsx_dicts_unique(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        worksheet_path = worksheet_path_by_name(zf, sheet_name) if sheet_name else first_worksheet_path(zf)
        with zf.open(worksheet_path) as handle:
            headers: list[str] | None = None
            for event, elem in iterparse(handle, events=("end",)):
                if not elem.tag.endswith("}row"):
                    continue

                row_number = int(elem.attrib.get("r", "0") or 0)
                values_by_index: dict[int, str] = {}
                max_index = 0
                for cell in list(elem):
                    if not cell.tag.endswith("}c"):
                        continue
                    ref_match = CELL_REF_RE.match(cell.attrib.get("r", ""))
                    index = col_index(ref_match.group(1)) if ref_match else len(values_by_index) + 1
                    values_by_index[index] = cell_value(cell, shared_strings)
                    max_index = max(max_index, index)

                row_values = [values_by_index.get(i, "") for i in range(1, max_index + 1)]
                if headers is None:
                    headers = unique_headers(row_values)
                else:
                    if not any(v.strip() for v in row_values):
                        elem.clear()
                        continue
                    padded = row_values + [""] * max(0, len(headers) - len(row_values))
                    yield row_number, dict(zip(headers, padded))
                elem.clear()


@lru_cache(maxsize=64)
def _cached_source_rows_unique(
    path_text: str,
    size_bytes: int,
    modified_ns: int,
    sheet_name: str | None,
) -> tuple[tuple[int, tuple[tuple[str, str], ...]], ...]:
    path = Path(path_text)
    rows = iter_csv_dicts(path) if path.suffix.lower() == ".csv" else iter_xlsx_dicts_unique(path, sheet_name)
    return tuple((row_number, tuple(row.items())) for row_number, row in rows)


def iter_source_rows_unique(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    if path.suffix.lower() == ".csv":
        sheet_name = None
    for row_number, items in _cached_source_rows_unique(*_source_cache_key(path, sheet_name)):
        yield row_number, dict(items)


def clear_source_row_cache() -> None:
    _cached_source_rows.cache_clear()
    _cached_source_rows_unique.cache_clear()


def normalize_mambu_disb_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_cents = amount_to_cents(row_value_at(row, 5, "Amount (GHC)", "Amount"))
    raw_key = row_value_at(row, 9, "Identifier (KEY)", "Identifier")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="Mambu",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "Value Date (Entry Date)", "Date/Time")),
        phone=str(pick(row, "Mobile Phone (Client)")),
        channel_type=str(pick(row, "Channel")),
        account=str(pick(row, "Loan Number", "Account Holder ID")),
        transaction_id=str(raw_key or ""),
    )


def normalize_mambu_collection_ledger_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "Amount (GHC)", "Amount", "Reconciliation_Amount")
    amount_cents = amount_to_cents(pick(row, amount_header))
    raw_key = row_value_at(row, 8, "Identifier (KEY)", "Identifier", "Reconciliation_Key")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="Mambu Collection",
        direction="COLLECTION",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "Value Date (Entry Date)", "Date/Time", "Reconciliation_DateTime")),
        phone=str(pick(row, "Mobile Phone (Client)", "Reconciliation_Phone")),
        channel_type=str(pick(row, "Channel", "Source_Dataset")),
        account=str(pick(row, "Account ID", "Account Holder ID")),
        transaction_id=str(raw_key or ""),
    )


def normalize_nsano_disb_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "Amount_GHC", "Amount (GHC)", "Amount")
    amount_cents = amount_to_cents(pick(row, amount_header))
    raw_key = row_value_at(row, 9, "SendingHse_ID (KEY)", "SendingHse_ID")
    key = normalize_disbursement_identifier(raw_key)
    datetime_value = parse_nsano_datetime(pick(row, "DateTime", "Date/Time"))

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="Nsano",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=datetime_value,
        phone=str(pick(row, "ReceivingHse_Account", "SendingHse_Account")),
        channel_type=str(pick(row, "Type")),
        account=str(pick(row, "ReceivingHse_ID", "ReceivingHse_Account")),
        result=str(pick(row, "Result")),
        transaction_id=str(pick(row, "Transaction_ID", "Trans ID")),
    )


def normalize_nsano_w2a_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    charge_value = row_value_at(row, 21, "Charge")
    amount_cents = amount_to_cents(charge_value)
    raw_key = row_value_at(row, 11, "External_Debit_Reference", "Ext Debit Ref (Key)", "Reconciliation_Key")
    key = normalize_disbursement_identifier(raw_key)
    datetime_value = parse_nsano_datetime(pick(row, "DateTime", "Date/Time", "Reconciliation_DateTime"))

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="Nsano Collection",
        direction="COLLECTION",
        key=key,
        amount_cents=amount_cents,
        datetime_value=datetime_value,
        phone=str(pick(row, "SendingHse_Account", "Reconciliation_Phone")),
        channel_type=str(pick(row, "Type", "Direction")),
        account=str(pick(row, "SendingHse_ID", "SendingHse_Account")),
        result=str(pick(row, "Result")),
        transaction_id=str(pick(row, "Transaction_ID", "Trans ID")),
    )


def load_nsano_w2a_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if sheet_name is None:
        sheet_name = xlsx_sheet_if_present(path, "Successful W2A")
    headers_captured = False
    for source_row, row in iter_source_rows(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_nsano_w2a_record(path, source_row, row)
        records.append(record)
    return records, headers


def normalize_itc_disb_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "amount", "Amount", "Amount (GHC)", "ITC Amount (GHC)")
    amount_cents = amount_to_cents(pick(row, amount_header))
    raw_key = row_value_at(row, 7, "thirdparty_id (KEY)", "thirdparty_id", "thirdparty_id (Key)")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="ITC DISB",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "transaction_date", "Transaction Date", "ITC Date")),
        phone=str(pick(row, "payer_contact", "account_reference")),
        channel_type=str(pick(row, "channel")),
        account=str(pick(row, "account_reference")),
        result=str(pick(row, "source")),
        transaction_id=str(pick(row, "processor_transaction_id")),
    )


def normalize_mtn_manual_disb_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "Amount", "AMOUNT", "Amount (GHC)")
    amount_cents = amount_to_cents(pick(row, amount_header))
    if amount_cents is not None:
        amount_cents = abs(amount_cents)
    raw_key = row_value_at(row, 1, "Id", "ID")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)
    if amount_cents is not None:
        values[amount_header] = cents_to_decimal(amount_cents)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="MTN MANUAL",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "Date", "DATE")),
        phone=str(pick(row, "To", "To account", "NUMBER")),
        channel_type=str(pick(row, "Type", "CHANNEL")),
        account=str(pick(row, "To account", "From account")),
        transaction_id=str(raw_key or ""),
    )


def normalize_vodafone_manual_disb_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "Withdrawn", "Paid In", "Amount", "AMOUNT", "Amount (GHC)")
    amount_cents = amount_to_cents(pick(row, amount_header))
    if amount_cents is not None:
        amount_cents = abs(amount_cents)
    raw_key = row_value_at(row, 1, "Receipt No.", "Receipt No", "Id", "ID")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="VODAFONE MANUAL",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "Completion Time", "Initiation Time", "Date", "DATE")),
        phone=str(pick(row, "Opposite Party", "Details", "To", "To account", "NUMBER")),
        channel_type=str(pick(row, "Reason Type", "Type", "CHANNEL")),
        account=str(pick(row, "Opposite Party", "Details", "To account", "From account")),
        transaction_id=str(raw_key or ""),
    )


def normalize_mtn_refund_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    amount_header = pick_key(row, "AMOUNT", "Amount", "Amount (GHC)")
    amount_cents = amount_to_cents(pick(row, amount_header))
    raw_key = row_value_at(row, 4, "TRANSACTION ID", "Transaction ID")
    key = normalize_disbursement_identifier(raw_key)

    values: OrderedDict[str, Any] = OrderedDict()
    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="REFUND",
        direction="DISBURSEMENT",
        key=key,
        amount_cents=amount_cents,
        datetime_value=str(pick(row, "DATE", "Date")),
        phone=str(pick(row, "NUMBER", "Number")),
        channel_type=str(pick(row, "CHANNEL", "Channel")),
        account=str(pick(row, "NUMBER", "Number")),
        transaction_id=str(raw_key or ""),
    )


def load_nsano_disb_sources(
    mambu_path: Path,
    nsano_path: Path,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(mambu_path, mambu_sheet_name)
    nsano_records, nsano_headers = load_nsano_disb_file(nsano_path)
    return mambu_records, mambu_headers, nsano_records, nsano_headers


def load_mambu_disb_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    headers_captured = False
    for source_row, row in iter_source_rows_unique(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_mambu_disb_record(path, source_row, row)
        records.append(record)
    return records, headers


def load_mambu_collection_ledger_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    headers_captured = False
    for source_row, row in iter_source_rows_unique(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_mambu_collection_ledger_record(path, source_row, row)
        records.append(record)
    return records, headers


def load_nsano_mambu_disb_ledger_file(
    path: Path,
    *,
    fallback_to_all: bool = True,
) -> tuple[list[Record], list[str]]:
    records, headers = _load_records_from_candidate_sheets(
        path,
        ("Nsano_Mambu", "Nsano Client Disb", "Nsano Client Disbursement", "Mambu Disbursement", "Mambu Disbursements", "Mambu"),
        load_mambu_disb_file,
    )
    return _records_matching(
        records,
        _is_nsano_mambu_disbursement_record,
        fallback_to_all=fallback_to_all,
    ), headers


def load_nsano_mambu_collection_ledger_file(path: Path) -> tuple[list[Record], list[str]]:
    records, headers = _load_records_from_candidate_sheets(
        path,
        ("Nsano Client Collection", "Nsano Client Collections", "Nsano_Mambu_Collections", "Mambu Collections", "Mambu"),
        load_mambu_collection_ledger_file,
    )
    return _records_matching(records, _is_nsano_mambu_collection_record), headers


def load_nsano_disb_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if sheet_name is None:
        sheet_name = xlsx_sheet_if_present(path, "Successful A2W")
    headers_captured = False
    for source_row, row in iter_source_rows(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_nsano_disb_record(path, source_row, row)
        records.append(record)
    return records, headers


def load_nsano_disb_ledger_file(path: Path) -> tuple[list[Record], list[str]]:
    records, headers = _load_records_from_candidate_sheets(
        path,
        ("Nsano_Disb.Statement", "NSANO DISB", "Nsano Disbursement", "Successful A2W", "Nsano"),
        load_nsano_disb_file,
    )
    return _records_matching(records, _is_nsano_disbursement_record), headers


NSANO_CHARGE_SHEET_CANDIDATES: tuple[str, ...] = (
    "Successful W2A",
    "Nsano Collections",
    "Nsano Collection",
    "Nsano Charges",
    "Nsano",
)


def load_nsano_charge_ledger_file(
    path: Path,
    sheet_name: str | None = None,
) -> tuple[list[Record], list[str]]:
    if sheet_name is not None:
        records, headers = load_nsano_w2a_file(path, sheet_name)
    else:
        records, headers = _load_records_from_candidate_sheets(
            path,
            NSANO_CHARGE_SHEET_CANDIDATES,
            load_nsano_w2a_file,
        )
    return [record for record in records if record.amount_cents], headers


def load_nsano_disb_workbook(path: Path) -> tuple[list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(path, "Mambu")
    nsano_records, nsano_headers = load_nsano_disb_file(path, "NSANO DISB")
    return mambu_records, mambu_headers, nsano_records, nsano_headers


def load_itc_disb_sources(
    mambu_path: Path,
    itc_path: Path,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(mambu_path, mambu_sheet_name)
    itc_records, itc_headers = load_itc_disb_file(itc_path)
    return mambu_records, mambu_headers, itc_records, itc_headers


def load_itc_disb_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if sheet_name is None:
        sheet_name = xlsx_sheet_if_present(path, "Outflow")
    headers_captured = False
    for source_row, row in iter_source_rows(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_itc_disb_record(path, source_row, row)
        records.append(record)
    return records, headers


def load_itc_disb_workbook(path: Path) -> tuple[list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(path, "Mambu")
    itc_records, itc_headers = load_itc_disb_file(path, "ITC DISB")
    return mambu_records, mambu_headers, itc_records, itc_headers


def load_mtn_manual_disb_sources(
    mambu_path: Path,
    mtn_manual_path: Path,
    refund_path: Path | None,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(mambu_path, mambu_sheet_name)
    mtn_records, mtn_headers = load_mtn_manual_disb_file(mtn_manual_path)
    refund_records, refund_headers = (
        load_mtn_refund_file(refund_path)
        if refund_path and refund_path.exists()
        else ([], list(MTN_MANUAL_REFUND_SHEET_COLUMNS))
    )
    return mambu_records, mambu_headers, mtn_records, mtn_headers, refund_records, refund_headers


def load_mtn_manual_disb_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    headers_captured = False
    for source_row, row in iter_source_rows(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_mtn_manual_disb_record(path, source_row, row)
        records.append(record)
    return records, headers or list(MTN_MANUAL_SHEET_COLUMNS)


def load_mtn_refund_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    headers_captured = False
    for source_row, row in iter_source_rows(path, sheet_name):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        record = normalize_mtn_refund_record(path, source_row, row)
        records.append(record)
    return records, headers or list(MTN_MANUAL_REFUND_SHEET_COLUMNS)


def load_mtn_manual_disb_workbook(path: Path) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
]:
    mambu_records, mambu_headers = load_mambu_disb_file(path, "Mambu")
    mtn_records, mtn_headers = load_mtn_manual_disb_file(path, "MTN MANUAL")
    refund_records, refund_headers = load_mtn_refund_file(path, "REFUND")
    return mambu_records, mambu_headers, mtn_records, mtn_headers, refund_records, refund_headers


def load_vodafone_manual_disb_sources(
    mambu_path: Path,
    vodafone_manual_path: Path,
    refund_path: Path | None,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str], list[Record], list[str]]:
    mambu_records, mambu_headers = load_mambu_disb_file(mambu_path, mambu_sheet_name)
    vodafone_records, vodafone_headers = load_vodafone_manual_disb_file(vodafone_manual_path)
    refund_records, refund_headers = (
        load_mtn_refund_file(refund_path)
        if refund_path and refund_path.exists()
        else ([], list(VODAFONE_MANUAL_REFUND_SHEET_COLUMNS))
    )
    return mambu_records, mambu_headers, vodafone_records, vodafone_headers, refund_records, refund_headers


def load_vodafone_manual_disb_file(path: Path, sheet_name: str | None = None) -> tuple[list[Record], list[str]]:
    headers, source_rows = _vodafone_manual_disb_rows(path, sheet_name)
    if not headers:
        raise ValueError("Vodafone Manual cleanup could not find a header row containing Id, Date, and Amount.")

    records: list[Record] = []
    for source_row, row in source_rows:
        record = normalize_vodafone_manual_disb_record(path, source_row, row)
        records.append(record)
    return records, headers or list(VODAFONE_MANUAL_SHEET_COLUMNS)


def load_vodafone_manual_disb_workbook(path: Path) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
]:
    mambu_records, mambu_headers = load_mambu_disb_file(path, "Mambu")
    vodafone_records, vodafone_headers = load_vodafone_manual_disb_file(path, "VODAFONE MANUAL")
    refund_records, refund_headers = load_mtn_refund_file(path, "REFUND")
    return mambu_records, mambu_headers, vodafone_records, vodafone_headers, refund_records, refund_headers


def load_itc_wallet_ledger_sources(
    statement_path: Path,
    debit_transfer_path: Path,
    credit_transfer_path: Path,
) -> tuple[
    list[OrderedDict[str, Any]],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    statement_sheet_name = xlsx_sheet_if_present(statement_path, "Outflow")
    statement_rows, statement_headers = load_generic_rows(statement_path, statement_sheet_name)
    debit_rows, debit_headers = load_generic_rows(debit_transfer_path)
    credit_rows, credit_headers = load_generic_rows(credit_transfer_path)
    return statement_rows, statement_headers, debit_rows, debit_headers, credit_rows, credit_headers


def load_itc_wallet_ledger_workbook(path: Path) -> tuple[
    list[OrderedDict[str, Any]],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    statement_rows, statement_headers = load_generic_rows(path, "ITC Statement")
    debit_rows, debit_headers = load_generic_rows(path, "ITC debit transfer1")
    credit_rows, credit_headers = load_generic_rows(path, "ITC Credit transfer")
    return statement_rows, statement_headers, debit_rows, debit_headers, credit_rows, credit_headers


def load_nsano_wallet_ledger_sources(
    mambu_path: Path,
    nsano_disb_path: Path,
    transfer_path: Path,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    mambu_records, mambu_headers = load_nsano_mambu_disb_ledger_file(mambu_path)
    nsano_disb_records, nsano_disb_headers = load_nsano_disb_ledger_file(nsano_disb_path)
    transfer_rows, transfer_headers = load_generic_rows(transfer_path)
    return mambu_records, mambu_headers, nsano_disb_records, nsano_disb_headers, transfer_rows, transfer_headers


def load_nsano_collection_ledger_sources(
    mambu_collection_path: Path,
    mambu_disb_path: Path,
    transfer_path: Path,
    nsano_charge_path: Path | None = None,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    mambu_collection_records, mambu_collection_headers = load_nsano_mambu_collection_ledger_file(mambu_collection_path)
    mambu_disb_records, mambu_disb_headers = load_nsano_mambu_disb_ledger_file(
        mambu_disb_path,
        fallback_to_all=False,
    )
    if nsano_charge_path is not None:
        nsano_charge_records, nsano_charge_headers = load_nsano_charge_ledger_file(nsano_charge_path)
    else:
        nsano_charge_records, nsano_charge_headers = [], []
    transfer_rows, transfer_headers = load_generic_rows(transfer_path)
    return (
        mambu_collection_records,
        mambu_collection_headers,
        mambu_disb_records,
        mambu_disb_headers,
        nsano_charge_records,
        nsano_charge_headers,
        transfer_rows,
        transfer_headers,
    )


def load_nsano_wallet_ledger_workbook(
    path: Path,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    mambu_records, mambu_headers = load_nsano_mambu_disb_ledger_file(path)
    nsano_disb_records, nsano_disb_headers = load_nsano_disb_ledger_file(path)

    try:
        transfer_rows, transfer_headers = load_generic_rows(path, "Nsano_Transfers")
    except KeyError:
        transfer_rows, transfer_headers = load_generic_rows(path, "Nsano Transfers")

    return mambu_records, mambu_headers, nsano_disb_records, nsano_disb_headers, transfer_rows, transfer_headers


def load_nsano_collection_ledger_workbook(
    path: Path,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[OrderedDict[str, Any]],
    list[str],
]:
    mambu_collection_records, mambu_collection_headers = load_nsano_mambu_collection_ledger_file(path)
    mambu_disb_records, mambu_disb_headers = load_nsano_mambu_disb_ledger_file(
        path,
        fallback_to_all=False,
    )

    charge_sheet_name = next(
        (name for name in NSANO_CHARGE_SHEET_CANDIDATES if xlsx_sheet_if_present(path, name)),
        None,
    )
    if charge_sheet_name is not None:
        nsano_charge_records, nsano_charge_headers = load_nsano_charge_ledger_file(path, charge_sheet_name)
    else:
        nsano_charge_records, nsano_charge_headers = [], []

    try:
        transfer_rows, transfer_headers = load_generic_rows(path, "Nsano_Transfers")
    except KeyError:
        transfer_rows, transfer_headers = load_generic_rows(path, "Nsano Transfers")

    return (
        mambu_collection_records,
        mambu_collection_headers,
        mambu_disb_records,
        mambu_disb_headers,
        nsano_charge_records,
        nsano_charge_headers,
        transfer_rows,
        transfer_headers,
    )


def load_generic_rows(path: Path, sheet_name: str | None = None) -> tuple[list[OrderedDict[str, Any]], list[str]]:
    headers: list[str] = []
    rows: list[OrderedDict[str, Any]] = []
    for source_row, row in iter_source_rows(path, sheet_name):
        add_unique_headers(headers, row.keys())
        cleaned: OrderedDict[str, Any] = OrderedDict()
        cleaned["_Source_File"] = path.name
        cleaned["_Source_Row"] = source_row
        for header, value in row.items():
            cleaned[header] = clean_sheet_value(header, value)
        rows.append(cleaned)
    return rows, headers


def build_key_index(records: list[Record]) -> dict[str, list[Record]]:
    index: dict[str, list[Record]] = defaultdict(list)
    for record in records:
        if record.key:
            index[record.key].append(record)
    return index


def compare_mambu_to_nsano_disb(
    mambu_records: list[Record],
    nsano_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    nsano_index = build_key_index(nsano_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mambu_records:
        match = nsano_index.get(source.key, [None])[0] if source.key else None
        status = DISB_NSANO_STATUS if match else DISB_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Mambu Identifier"] = first_value(source, "Identifier")
        row["Mambu Date"] = first_value(source, "Value Date (Entry Date)", "Date/Time")
        row["Mambu Account Holder"] = first_value(source, "Account Holder Name")
        row["Mambu Amount (GHC)"] = source.amount_decimal
        row["Mambu Loan No."] = first_value(source, "Loan Number")
        row["Mambu Mobile"] = first_value(source, "Mobile Phone (Client)")
        row["Match Status"] = status
        row["Nsano Trans ID"] = first_value(match, "Transaction_ID", "Trans ID") if match else ""
        row["Nsano DateTime"] = first_value(match, "DateTime", "Date/Time") if match else ""
        row["Nsano Amount (GHC)"] = match.amount_decimal if match else ""
        row["Nsano Row"] = match.source_row if match else ""
        rows.append(row)

    return rows


def compare_nsano_to_mambu_disb(
    nsano_records: list[Record],
    mambu_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mambu_index = build_key_index(mambu_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in nsano_records:
        match = mambu_index.get(source.key, [None])[0] if source.key else None
        status = DISB_MAMBU_STATUS if match else DISB_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Nsano Identifier (stripped)"] = source.key
        row["Nsano DateTime"] = first_value(source, "DateTime", "Date/Time")
        row["Nsano Trans ID"] = first_value(source, "Transaction_ID", "Trans ID")
        row["Nsano Amount (GHC)"] = source.amount_decimal
        row["Nsano Result"] = first_value(source, "Result")
        row["Nsano Receiving Acct"] = first_value(source, "ReceivingHse_Account")
        row["Match Status"] = status
        row["Mambu Identifier"] = first_value(match, "Identifier") if match else ""
        row["Mambu Date"] = first_value(match, "Value Date (Entry Date)", "Date/Time") if match else ""
        row["Mambu Amount (GHC)"] = match.amount_decimal if match else ""
        row["Mambu Loan No."] = first_value(match, "Loan Number") if match else ""
        row["Mambu Row"] = match.source_row if match else ""
        rows.append(row)

    return rows


def compare_mambu_to_itc_disb(
    mambu_records: list[Record],
    itc_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    itc_index = build_key_index(itc_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mambu_records:
        match = itc_index.get(source.key, [None])[0] if source.key else None
        status = ITC_DISB_STATUS if match else ITC_DISB_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Mambu Identifier"] = first_value(source, "Identifier (KEY)", "Identifier")
        row["Mambu Date"] = first_value(source, "Value Date (Entry Date)", "Date/Time")
        row["Mambu Account Holder"] = first_value(source, "Account Holder Name")
        row["Mambu Amount (GHC)"] = source.amount_decimal
        row["Mambu Loan No."] = first_value(source, "Loan Number")
        row["Mambu Mobile"] = first_value(source, "Mobile Phone (Client)")
        row["Match Status"] = status
        row["ITC Trans ID"] = first_value(match, "processor_transaction_id") if match else ""
        row["ITC Date"] = first_value(match, "transaction_date", "Transaction Date", "ITC Date") if match else ""
        row["ITC Amount (GHC)"] = match.amount_decimal if match else ""
        row["ITC Channel"] = first_value(match, "channel") if match else ""
        row["ITC Row"] = match.source_row if match else ""
        rows.append(row)

    return rows


def compare_itc_disb_to_mambu(
    itc_records: list[Record],
    mambu_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mambu_index = build_key_index(mambu_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in itc_records:
        match = mambu_index.get(source.key, [None])[0] if source.key else None
        narration = str(first_value(source, "narration")).strip()
        status = DISB_MAMBU_STATUS if match else itc_disb_not_found_status_for_narration(narration)
        row: OrderedDict[str, Any] = OrderedDict()
        row["ITC thirdparty_id (Key)"] = first_value(source, "thirdparty_id (KEY)", "thirdparty_id")
        row["ITC Date"] = first_value(source, "transaction_date", "Transaction Date", "ITC Date")
        row["ITC Processor Trans ID"] = first_value(source, "processor_transaction_id")
        row["ITC Amount (GHC)"] = source.amount_decimal
        row["ITC Channel"] = first_value(source, "channel")
        row["ITC Source"] = first_value(source, "source")
        row["Match Status"] = status
        row["Mambu Identifier"] = first_value(match, "Identifier (KEY)", "Identifier") if match else ""
        row["Mambu Date"] = first_value(match, "Value Date (Entry Date)", "Date/Time") if match else ""
        row["Mambu Amount (GHC)"] = match.amount_decimal if match else ""
        row["Mambu Account Holder"] = first_value(match, "Account Holder Name") if match else ""
        row["Narration"] = narration
        row["Mambu Row"] = match.source_row if match else ""
        rows.append(row)

    return rows


def compare_mambu_to_mtn_manual_disb(
    mambu_records: list[Record],
    mtn_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mtn_index = build_key_index(mtn_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mambu_records:
        match = mtn_index.get(source.key, [None])[0] if source.key else None
        status = MTN_MANUAL_MATCHED_STATUS if match else MTN_MANUAL_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Value Date"] = first_value(source, "Value Date (Entry Date)", "Date/Time")
        row["Account Holder ID"] = first_value(source, "Account Holder ID")
        row["Account Holder Name"] = first_value(source, "Account Holder Name")
        row["Principal Amount"] = first_value(source, "Principal Amount")
        row["Amount"] = source.amount_decimal
        row["Loan Number"] = first_value(source, "Loan Number")
        row["Channel"] = first_value(source, "Channel")
        row["Mobile Phone"] = first_value(source, "Mobile Phone (Client)")
        row["Identifier"] = first_value(source, "Identifier", "Identifier (KEY)")
        row["Why Loan?"] = first_value(source, "Why do you need a loan?", "Loan Purpose")
        row["Gender"] = first_value(source, "Gender (Client)")
        row["Birth Date"] = first_value(source, "Birth Date (Client)")
        row["Branch"] = first_value(source, "Branch Name")
        row["Match Status"] = status
        row["MTN_MATCH"] = match.source_row if match else 0
        rows.append(row)

    return rows


def compare_mtn_manual_disb_to_sources(
    mtn_records: list[Record],
    mambu_records: list[Record],
    refund_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mambu_index = build_key_index(mambu_records)
    refund_index = build_key_index(refund_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mtn_records:
        mambu_match = mambu_index.get(source.key, [None])[0] if source.key else None
        refund_match = refund_index.get(source.key, [None])[0] if source.key else None
        if mambu_match:
            status = MTN_MANUAL_MAMBU_STATUS
        elif refund_match:
            status = MTN_MANUAL_REFUND_STATUS
        else:
            status = MTN_MANUAL_NOT_FOUND_STATUS

        row: OrderedDict[str, Any] = OrderedDict()
        row["Id"] = first_value(source, "Id", "ID")
        row["External Id"] = first_value(source, "External id", "External Id")
        row["Date"] = first_value(source, "Date")
        row["Status"] = first_value(source, "Status")
        row["Type"] = first_value(source, "Type")
        row["Amount"] = source.amount_decimal
        row["Match Status"] = status
        row["MAMBU_MATCH"] = mambu_match.source_row if mambu_match else 0
        row["REFUND_MATCH"] = refund_match.source_row if refund_match else 0
        rows.append(row)

    return rows


def compare_mambu_to_vodafone_manual_disb(
    mambu_records: list[Record],
    vodafone_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    vodafone_index = build_key_index(vodafone_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in mambu_records:
        match = vodafone_index.get(source.key, [None])[0] if source.key else None
        status = VODAFONE_MANUAL_MATCHED_STATUS if match else VODAFONE_MANUAL_NOT_FOUND_STATUS
        row: OrderedDict[str, Any] = OrderedDict()
        row["Value Date"] = first_value(source, "Value Date (Entry Date)", "Date/Time")
        row["Account Holder ID"] = first_value(source, "Account Holder ID")
        row["Account Holder Name"] = first_value(source, "Account Holder Name")
        row["Principal Amount"] = first_value(source, "Principal Amount")
        row["Amount"] = source.amount_decimal
        row["Loan Number"] = first_value(source, "Loan Number")
        row["Channel"] = first_value(source, "Channel")
        row["Mobile Phone"] = first_value(source, "Mobile Phone (Client)")
        row["Identifier"] = first_value(source, "Identifier", "Identifier (KEY)")
        row["Why Loan?"] = first_value(source, "Why do you need a loan?", "Loan Purpose")
        row["Gender"] = first_value(source, "Gender (Client)")
        row["Birth Date"] = first_value(source, "Birth Date (Client)")
        row["Branch"] = first_value(source, "Branch Name")
        row["Match Status"] = status
        row["VF_MATCH"] = match.source_row if match else 0
        rows.append(row)

    return rows


def compare_vodafone_manual_disb_to_sources(
    vodafone_records: list[Record],
    mambu_records: list[Record],
    refund_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    mambu_index = build_key_index(mambu_records)
    refund_index = build_key_index(refund_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in vodafone_records:
        mambu_match = mambu_index.get(source.key, [None])[0] if source.key else None
        refund_match = refund_index.get(source.key, [None])[0] if source.key else None
        if mambu_match:
            status = VODAFONE_MANUAL_MAMBU_STATUS
        elif refund_match:
            status = VODAFONE_MANUAL_REFUND_STATUS
        else:
            status = VODAFONE_MANUAL_NOT_FOUND_STATUS

        row: OrderedDict[str, Any] = OrderedDict()
        row["Id"] = first_value(source, "Id", "ID")
        row["External Id"] = first_value(source, "External id", "External Id")
        row["Date"] = first_value(source, "Date")
        row["Status"] = first_value(source, "Status")
        row["Type"] = first_value(source, "Type")
        row["Amount"] = source.amount_decimal
        row["Match Status"] = status
        row["MAMBU_MATCH"] = mambu_match.source_row if mambu_match else 0
        row["REFUND_MATCH"] = refund_match.source_row if refund_match else 0
        rows.append(row)

    return rows


def row_value_by_column(row: OrderedDict[str, Any], column_index: int, *fallback_headers: str) -> Any:
    visible_values = [
        value for key, value in row.items()
        if not str(key).startswith("_")
    ]
    if len(visible_values) >= column_index and visible_values[column_index - 1] not in ("", None):
        return visible_values[column_index - 1]
    for header in fallback_headers:
        value = row.get(header, "")
        if value not in ("", None):
            return value
    return ""


def wallet_identifier(value: Any) -> str:
    return normalize_disbursement_identifier(value)


def wallet_row_amount_cents(row: OrderedDict[str, Any], column_index: int, *fallback_headers: str) -> int:
    return amount_to_cents(row_value_by_column(row, column_index, *fallback_headers)) or 0


def wallet_row_amount(row: OrderedDict[str, Any], column_index: int, *fallback_headers: str) -> Decimal:
    return cents_to_decimal(wallet_row_amount_cents(row, column_index, *fallback_headers))


def month_from_value(value: Any) -> tuple[int, int] | None:
    text = str(value or "").strip()
    if not text:
        return None

    iso_match = re.search(r"\b(20\d{2})[-/](\d{1,2})(?:[-/]\d{1,2})?\b", text)
    if iso_match:
        year = int(iso_match.group(1))
        month = int(iso_match.group(2))
        if 1 <= month <= 12:
            return year, month

    day_first_match = re.search(r"\b\d{1,2}[-/](\d{1,2})[-/](20\d{2})\b", text)
    if day_first_match:
        month = int(day_first_match.group(1))
        year = int(day_first_match.group(2))
        if 1 <= month <= 12:
            return year, month

    try:
        serial = float(text)
    except ValueError:
        return None
    if not 20000 <= serial <= 80000:
        return None
    date_value = dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))
    return date_value.year, date_value.month


def detect_itc_wallet_month(
    statement_rows: list[OrderedDict[str, Any]],
    debit_rows: list[OrderedDict[str, Any]],
    credit_rows: list[OrderedDict[str, Any]],
) -> tuple[str, str]:
    month_counts: Counter[tuple[int, int]] = Counter()
    sources = [
        (statement_rows, ((9, "transaction_date"),)),
        (debit_rows, ((4, "transaction_date"), (5, "prepaid_transaction_date"))),
        (credit_rows, ((4, "transaction_date"), (5, "prepaid_transaction_date"))),
    ]
    for rows, columns in sources:
        for row in rows:
            for column_index, header in columns:
                month = month_from_value(row_value_by_column(row, column_index, header))
                if month:
                    month_counts[month] += 1
                    break

    if not month_counts:
        return "MONTH YEAR", "Month Year"

    (year, month), _ = month_counts.most_common(1)[0]
    date_value = dt.date(year, month, 1)
    return date_value.strftime("%B %Y").upper(), date_value.strftime("%b %Y")


def detect_nsano_wallet_month(
    mambu_records: list[Record],
    transfer_rows: list[OrderedDict[str, Any]],
) -> tuple[str, str]:
    month_counts: Counter[tuple[int, int]] = Counter()
    for record in mambu_records:
        month = month_from_value(first_value(record, "Value Date (Entry Date)", "Date/Time"))
        if month:
            month_counts[month] += 1
    for row in transfer_rows:
        month = month_from_value(row_value_by_column(row, 1, "Date"))
        if month:
            month_counts[month] += 1

    if not month_counts:
        return "MONTH YEAR", "Month Year"

    (year, month), _ = month_counts.most_common(1)[0]
    date_value = dt.date(year, month, 1)
    return date_value.strftime("%B %Y").upper(), date_value.strftime("%b %Y")


def narration_suffix_token(value: Any) -> str:
    text = str(value or "").strip().casefold()
    for suffix in ("_1", "_2", "_3", "_4", "_6"):
        if text.endswith(suffix):
            return suffix
    return ""


def classify_itc_credit_transfer(row: OrderedDict[str, Any]) -> str:
    narration = str(row_value_by_column(row, 7, "narration")).strip().casefold()
    compact = re.sub(r"[^a-z0-9]+", "", narration)
    if "settlement" in compact:
        return ITC_WALLET_SETTLEMENT_STATUS
    if "prepaidtransactionreversal" in compact or "prepaidtr" in compact:
        return ITC_WALLET_PREPAID_REVERSAL_STATUS
    if "reversalof" in compact or compact.startswith("reversal"):
        return ITC_WALLET_REVERSAL_STATUS
    return ITC_WALLET_TRANSFER_TO_WALLET_STATUS


def classify_itc_debit_transfer(
    debit_row: OrderedDict[str, Any],
    statement_row: OrderedDict[str, Any] | None,
    credit_row: OrderedDict[str, Any] | None,
) -> str:
    if statement_row is not None:
        suffix = narration_suffix_token(row_value_by_column(statement_row, 13, "narration"))
        if suffix in {"_1", "_4"}:
            return ITC_WALLET_DISBURSEMENT_STATUS
        if suffix == "_2":
            return ITC_WALLET_REFERRAL_AWARD_STATUS
        if suffix == "_3":
            return ITC_WALLET_UPSALES_REFUND_STATUS
        if suffix == "_6":
            return ITC_WALLET_SAVINGS_STATUS
        return ITC_WALLET_DISBURSEMENT_STATUS

    debit_text = " ".join(str(value or "") for value in debit_row.values()).casefold()
    if "withdrawal" in debit_text:
        return ITC_WALLET_TRANSFERS_FROM_BANK_STATUS
    if credit_row is not None:
        return ITC_WALLET_CREDIT_TRANSFER_STATUS
    return ITC_WALLET_NOT_FOUND_STATUS


def compact_wallet_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


NSANO_COLLECTION_PURPOSE_PATTERN = re.compile(
    r"^\s*top[\s-]*up\s+(?:through\s+)?"
    r"(?P<collection>collections?|collectrion|colletion|coll)\b",
    re.IGNORECASE,
)
NSANO_COLLECTION_PURPOSE_EXCEPTIONS = {
    "collectrion": "Spelling variation: collectrion interpreted as collection",
    "colletion": "Spelling variation: colletion interpreted as collection",
    "coll": "Abbreviation: coll interpreted as collection",
}


def nsano_collection_purpose_exception(value: Any) -> str:
    """Return an exception reason for an accepted nonstandard collection purpose."""
    match = NSANO_COLLECTION_PURPOSE_PATTERN.search(str(value or ""))
    if match is None:
        return ""
    collection_word = match.group("collection").casefold()
    return NSANO_COLLECTION_PURPOSE_EXCEPTIONS.get(collection_word, "")


def classify_nsano_transfer_purpose(value: Any) -> str:
    compact = compact_wallet_text(value)
    if NSANO_COLLECTION_PURPOSE_PATTERN.search(str(value or "")):
        return NSANO_WALLET_TOPUP_COLLECTION_STATUS
    if "topup" in compact and "stanbic" in compact:
        return NSANO_WALLET_BANK_TO_WALLET_STATUS
    if "reversaladjustment" in compact:
        return NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS
    if "transfertobank" in compact or "transferbank" in compact:
        return NSANO_WALLET_TRANSFER_TO_BANK_STATUS
    return NSANO_WALLET_OTHER_STATUS


def build_nsano_transfer_classification_rows(
    transfer_rows: list[OrderedDict[str, Any]],
) -> list[OrderedDict[str, Any]]:
    rows: list[OrderedDict[str, Any]] = []
    for source in transfer_rows:
        amount = wallet_row_amount(source, 4, "Amount", "Amount1")
        purpose = row_value_by_column(source, 5, "Purpose")
        category = classify_nsano_transfer_purpose(purpose)
        row: OrderedDict[str, Any] = OrderedDict()
        row["Transfer Date"] = row_value_by_column(source, 1, "Date")
        row["Map Name"] = row_value_by_column(source, 3, "Map Name")
        row["Purpose"] = purpose
        row["Transfer Amount (GHC)"] = amount
        row["Transfer Category"] = category
        exception_reason = nsano_collection_purpose_exception(purpose)
        row["Narration Exception"] = "Yes" if exception_reason else ""
        row["Exception Reason"] = exception_reason
        row["TOP-UP Through Collection with Date"] = (
            amount if category == NSANO_WALLET_TOPUP_COLLECTION_STATUS else ""
        )
        row["Bank to Wallet"] = amount if category == NSANO_WALLET_BANK_TO_WALLET_STATUS else ""
        row["Reversal Adjustment"] = (
            amount if category == NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS else ""
        )
        row["Transfer to Bank"] = amount if category == NSANO_WALLET_TRANSFER_TO_BANK_STATUS else ""
        row["Source Row"] = source.get("_Source_Row", "")
        rows.append(row)
    return rows


def build_itc_credit_transfer_rows(
    credit_rows: list[OrderedDict[str, Any]],
) -> list[OrderedDict[str, Any]]:
    rows: list[OrderedDict[str, Any]] = []
    for source in credit_rows:
        category = classify_itc_credit_transfer(source)
        amount = wallet_row_amount(source, 6, "amount")
        row: OrderedDict[str, Any] = OrderedDict()
        row["Credit Identifier"] = row_value_by_column(source, 3, "third_party_transaction_id")
        row["Credit Date"] = row_value_by_column(source, 4, "transaction_date")
        row["Credit Amount (GHC)"] = amount
        row["Credit Narration"] = row_value_by_column(source, 7, "narration")
        row["Credit Category"] = category
        row["Settlement"] = amount if category == ITC_WALLET_SETTLEMENT_STATUS else ""
        row["Reversal"] = amount if category == ITC_WALLET_REVERSAL_STATUS else ""
        row["Prepaid Reversal"] = amount if category == ITC_WALLET_PREPAID_REVERSAL_STATUS else ""
        row["Transfer to Wallet"] = amount if category == ITC_WALLET_TRANSFER_TO_WALLET_STATUS else ""
        row["Source Row"] = source.get("_Source_Row", "")
        rows.append(row)
    return rows


def compare_itc_debit_transfers_to_statement(
    debit_rows: list[OrderedDict[str, Any]],
    statement_rows: list[OrderedDict[str, Any]],
    credit_rows: list[OrderedDict[str, Any]],
) -> list[OrderedDict[str, Any]]:
    statement_index = {
        wallet_identifier(row_value_by_column(row, 1, "processor_transaction_id")): row
        for row in statement_rows
        if wallet_identifier(row_value_by_column(row, 1, "processor_transaction_id"))
    }
    credit_index = {
        wallet_identifier(row_value_by_column(row, 3, "third_party_transaction_id")): row
        for row in credit_rows
        if wallet_identifier(row_value_by_column(row, 3, "third_party_transaction_id"))
    }
    rows: list[OrderedDict[str, Any]] = []

    for debit in debit_rows:
        identifier = wallet_identifier(row_value_by_column(debit, 3, "third_party_transaction_id"))
        statement = statement_index.get(identifier) if identifier else None
        credit = credit_index.get(identifier) if identifier else None
        status = classify_itc_debit_transfer(debit, statement, credit)
        row: OrderedDict[str, Any] = OrderedDict()
        row["Debit Identifier"] = row_value_by_column(debit, 3, "third_party_transaction_id")
        row["Debit Date"] = row_value_by_column(debit, 4, "transaction_date")
        row["Debit Amount (GHC)"] = wallet_row_amount(debit, 6, "amount")
        row["Debit Narration"] = row_value_by_column(debit, 7, "narration")
        row["Match Status"] = status
        row["Statement Identifier"] = row_value_by_column(statement, 1, "processor_transaction_id") if statement else ""
        row["Statement Narration"] = row_value_by_column(statement, 13, "narration") if statement else ""
        row["Statement Amount (GHC)"] = wallet_row_amount(statement, 15, "amount") if statement else ""
        row["Statement Row"] = statement.get("_Source_Row", "") if statement else ""
        row["Credit Transfer Identifier"] = row_value_by_column(credit, 3, "third_party_transaction_id") if credit else ""
        row["Credit Transfer Row"] = credit.get("_Source_Row", "") if credit else ""
        rows.append(row)
    return rows


def itc_disb_not_found_status_for_narration(narration: str) -> str:
    if narration.endswith("_2"):
        return ITC_DISB_REFERRAL_BONUS_STATUS
    if narration.endswith("_3"):
        return ITC_DISB_UPSALES_REFOUND_STATUS
    if narration.endswith("_6"):
        return ITC_DISB_SAVINGS_REWARD_STATUS
    return ITC_DISB_NOT_FOUND_STATUS


def first_value(record: Record | None, *headers: str) -> Any:
    if record is None:
        return ""
    for header in headers:
        value = record.sheet_values.get(header, "")
        if value not in ("", None):
            return value
    return ""


def summarize_disb_compare(
    rows: list[OrderedDict[str, Any]],
    *,
    matched_status: str,
    amount_header: str,
    not_found_status: str = DISB_NOT_FOUND_STATUS,
) -> dict[str, Any]:
    status_counts = Counter(str(row.get("Match Status", "")) for row in rows)
    status_amount_cents: Counter[str] = Counter()
    total_amount_cents = 0

    for row in rows:
        status = str(row.get("Match Status", ""))
        amount_cents = amount_to_cents(row.get(amount_header, ""))
        if amount_cents is None:
            continue
        total_amount_cents += amount_cents
        status_amount_cents[status] += amount_cents

    total = len(rows)
    matched = status_counts[matched_status]
    not_found = status_counts[not_found_status]
    return {
        "total": total,
        "matched": matched,
        "not_found": not_found,
        "match_rate": (matched / total) if total else 0,
        "source_amount": cents_to_decimal(total_amount_cents),
        "matched_amount": cents_to_decimal(status_amount_cents[matched_status]),
        "not_found_amount": cents_to_decimal(status_amount_cents[not_found_status]),
        "status_counts": dict(status_counts),
        "status_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_amount_cents.items()
        },
    }


def summarize_disb_compare_multi(
    rows: list[OrderedDict[str, Any]],
    *,
    matched_statuses: set[str],
    amount_header: str,
    not_found_status: str,
) -> dict[str, Any]:
    status_counts = Counter(str(row.get("Match Status", "")) for row in rows)
    status_amount_cents: Counter[str] = Counter()
    total_amount_cents = 0

    for row in rows:
        status = str(row.get("Match Status", ""))
        amount_cents = amount_to_cents(row.get(amount_header, ""))
        if amount_cents is None:
            continue
        total_amount_cents += amount_cents
        status_amount_cents[status] += amount_cents

    total = len(rows)
    matched = sum(status_counts[status] for status in matched_statuses)
    matched_amount_cents = sum(status_amount_cents[status] for status in matched_statuses)
    return {
        "total": total,
        "matched": matched,
        "not_found": status_counts[not_found_status],
        "match_rate": (matched / total) if total else 0,
        "source_amount": cents_to_decimal(total_amount_cents),
        "matched_amount": cents_to_decimal(matched_amount_cents),
        "not_found_amount": cents_to_decimal(status_amount_cents[not_found_status]),
        "status_counts": dict(status_counts),
        "status_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_amount_cents.items()
        },
    }


def cents_to_decimal(cents: int | None) -> Decimal:
    if cents is None:
        return Decimal("0")
    return Decimal(cents) / Decimal(100)


def sum_amounts(records: list[Record]) -> Decimal:
    return cents_to_decimal(sum(record.amount_cents or 0 for record in records))


def build_nsano_disb_summary_rows(
    mambu_vs_nsano_stats: dict[str, Any],
    nsano_vs_mambu_stats: dict[str, Any],
) -> list[list[Any]]:
    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    mvn = mambu_vs_nsano_stats
    nvm = nsano_vs_mambu_stats
    return [
        [Cell("NSANO DISBURSEMENT RECONCILIATION - SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu vs Nsano"), th("Count"), th("Amount (GHC)")],
        [DISB_NSANO_STATUS, mvn["matched"], Cell(mvn["matched_amount"], STYLE_MONEY)],
        [DISB_NOT_FOUND_STATUS, mvn["not_found"], Cell(mvn["not_found_amount"], STYLE_MONEY)],
        [th("TOTAL"), tc(mvn["total"]), tm(mvn["source_amount"])],
        ["", "", ""],
        [th("Nsano vs Mambu"), th("Count"), th("Amount (GHC)")],
        [DISB_MAMBU_STATUS, nvm["matched"], Cell(nvm["matched_amount"], STYLE_MONEY)],
        [DISB_NOT_FOUND_STATUS, nvm["not_found"], Cell(nvm["not_found_amount"], STYLE_MONEY)],
        [th("TOTAL"), tc(nvm["total"]), tm(nvm["source_amount"])],
    ]


def narration_suffix(value: Any) -> str:
    text = str(value or "").strip().casefold()
    for suffix in ("_1", "_2", "_3", "_4", "_6"):
        if text.endswith(suffix):
            return suffix
    return "other" if text else "blank"


def summarize_itc_disb_narration_breakdowns(
    itc_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> dict[str, dict[str, dict[str, Any]]]:
    totals: dict[str, dict[str, dict[str, int]]] = defaultdict(
        lambda: defaultdict(lambda: {"count": 0, "amount_cents": 0})
    )

    for row in itc_vs_mambu_rows:
        status = str(row.get("Match Status", ""))
        if status == DISB_MAMBU_STATUS:
            group = "matched"
        elif status == ITC_DISB_NOT_FOUND_STATUS:
            group = "not_found"
        else:
            continue
        key = narration_suffix(row.get("Narration", ""))
        amount_cents = amount_to_cents(row.get("ITC Amount (GHC)", "")) or 0
        totals[group][key]["count"] += 1
        totals[group][key]["amount_cents"] += amount_cents

    return {
        group: {
            key: {
                "count": data["count"],
                "amount": cents_to_decimal(data["amount_cents"]),
            }
            for key, data in group_data.items()
        }
        for group, group_data in totals.items()
    }


def build_itc_disb_summary_rows(
    mambu_vs_itc_stats: dict[str, Any],
    itc_vs_mambu_stats: dict[str, Any],
    itc_breakdowns: dict[str, dict[str, dict[str, Any]]],
) -> list[list[Any]]:
    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    def count(stats: dict[str, Any], status: str) -> int:
        return int(stats["status_counts"].get(status, 0))

    def amount(stats: dict[str, Any], status: str) -> Decimal:
        return stats["status_amounts"].get(status, Decimal("0"))

    def breakdown(group: str, key: str) -> dict[str, Any]:
        return itc_breakdowns.get(group, {}).get(key, {"count": 0, "amount": Decimal("0")})

    def bcount(group: str, key: str) -> int:
        return int(breakdown(group, key)["count"])

    def bamount(group: str, key: str) -> Decimal:
        return breakdown(group, key)["amount"]

    mvi = mambu_vs_itc_stats
    ivm = itc_vs_mambu_stats
    rows: list[list[Any]] = [
        [Cell("ITC WALLET VS MAMBU - SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu vs ITC Wallet"), th("Count"), th("Amount (GHC)")],
        [Cell(ITC_DISB_STATUS, STYLE_MATCHED), count(mvi, ITC_DISB_STATUS), Cell(amount(mvi, ITC_DISB_STATUS), STYLE_MONEY)],
        [Cell(ITC_DISB_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(mvi, ITC_DISB_NOT_FOUND_STATUS), Cell(amount(mvi, ITC_DISB_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mvi["total"]), tm(mvi["source_amount"])],
        ["", "", ""],
        [th("ITC Wallet vs Mambu"), th("Count"), th("Amount (GHC)")],
        [Cell(DISB_MAMBU_STATUS, STYLE_MATCHED), count(ivm, DISB_MAMBU_STATUS), Cell(amount(ivm, DISB_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(ITC_DISB_REFERRAL_BONUS_STATUS, STYLE_NOT_FOUND), count(ivm, ITC_DISB_REFERRAL_BONUS_STATUS), Cell(amount(ivm, ITC_DISB_REFERRAL_BONUS_STATUS), STYLE_MONEY)],
        [Cell(ITC_DISB_UPSALES_REFOUND_STATUS, STYLE_NOT_FOUND), count(ivm, ITC_DISB_UPSALES_REFOUND_STATUS), Cell(amount(ivm, ITC_DISB_UPSALES_REFOUND_STATUS), STYLE_MONEY)],
        [Cell(ITC_DISB_SAVINGS_REWARD_STATUS, STYLE_NOT_FOUND), count(ivm, ITC_DISB_SAVINGS_REWARD_STATUS), Cell(amount(ivm, ITC_DISB_SAVINGS_REWARD_STATUS), STYLE_MONEY)],
        [Cell(ITC_DISB_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(ivm, ITC_DISB_NOT_FOUND_STATUS), Cell(amount(ivm, ITC_DISB_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(ivm["total"]), tm(ivm["source_amount"])],
        ["", "", ""],
        [th("Matched - Narration Breakdown"), th("Count"), th("Amount (GHC)")],
    ]

    for suffix in ("_1", "_2", "_3", "_4", "_6"):
        rows.append([f"Matched - {suffix}", bcount("matched", suffix), Cell(bamount("matched", suffix), STYLE_MONEY)])

    rows.extend([
        ["", "", ""],
        [th("Not Found - Narration Breakdown"), th("Count"), th("Amount (GHC)")],
    ])
    for suffix in ("_1", "_2", "_3", "_4", "_6"):
        rows.append([f"Not Found - {suffix}", bcount("not_found", suffix), Cell(bamount("not_found", suffix), STYLE_MONEY)])

    return rows


def build_mtn_manual_disb_summary_rows(
    mambu_vs_mtn_stats: dict[str, Any],
    mtn_vs_sources_stats: dict[str, Any],
) -> list[list[Any]]:
    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    def count(stats: dict[str, Any], status: str) -> int:
        return int(stats["status_counts"].get(status, 0))

    def amount(stats: dict[str, Any], status: str) -> Decimal:
        return stats["status_amounts"].get(status, Decimal("0"))

    mvm = mambu_vs_mtn_stats
    mtv = mtn_vs_sources_stats
    return [
        [Cell("MTN MANUAL DISBURSEMENT RECONCILIATION - SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu vs MTN MANUAL"), th("Count"), th("Amount (GHC)")],
        [Cell(MTN_MANUAL_MATCHED_STATUS, STYLE_MATCHED), count(mvm, MTN_MANUAL_MATCHED_STATUS), Cell(amount(mvm, MTN_MANUAL_MATCHED_STATUS), STYLE_MONEY)],
        [Cell(MTN_MANUAL_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(mvm, MTN_MANUAL_NOT_FOUND_STATUS), Cell(amount(mvm, MTN_MANUAL_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mvm["total"]), tm(mvm["source_amount"])],
        ["", "", ""],
        [th("MTN MANUAL as Source"), th("Count"), th("Amount (GHC)")],
        [Cell(MTN_MANUAL_MAMBU_STATUS, STYLE_MATCHED), count(mtv, MTN_MANUAL_MAMBU_STATUS), Cell(amount(mtv, MTN_MANUAL_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(MTN_MANUAL_REFUND_STATUS, STYLE_NOT_FOUND), count(mtv, MTN_MANUAL_REFUND_STATUS), Cell(amount(mtv, MTN_MANUAL_REFUND_STATUS), STYLE_MONEY)],
        [Cell(MTN_MANUAL_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(mtv, MTN_MANUAL_NOT_FOUND_STATUS), Cell(amount(mtv, MTN_MANUAL_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mtv["total"]), tm(mtv["source_amount"])],
    ]


def build_vodafone_manual_disb_summary_rows(
    mambu_vs_vodafone_stats: dict[str, Any],
    vodafone_vs_sources_stats: dict[str, Any],
) -> list[list[Any]]:
    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    def count(stats: dict[str, Any], status: str) -> int:
        return int(stats["status_counts"].get(status, 0))

    def amount(stats: dict[str, Any], status: str) -> Decimal:
        return stats["status_amounts"].get(status, Decimal("0"))

    mvv = mambu_vs_vodafone_stats
    vtv = vodafone_vs_sources_stats
    return [
        [Cell("VODAFONE MANUAL DISBURSEMENT RECONCILIATION - SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu vs VODAFONE MANUAL"), th("Count"), th("Amount (GHC)")],
        [Cell(VODAFONE_MANUAL_MATCHED_STATUS, STYLE_MATCHED), count(mvv, VODAFONE_MANUAL_MATCHED_STATUS), Cell(amount(mvv, VODAFONE_MANUAL_MATCHED_STATUS), STYLE_MONEY)],
        [Cell(VODAFONE_MANUAL_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(mvv, VODAFONE_MANUAL_NOT_FOUND_STATUS), Cell(amount(mvv, VODAFONE_MANUAL_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mvv["total"]), tm(mvv["source_amount"])],
        ["", "", ""],
        [th("VODAFONE MANUAL as Source"), th("Count"), th("Amount (GHC)")],
        [Cell(VODAFONE_MANUAL_MAMBU_STATUS, STYLE_MATCHED), count(vtv, VODAFONE_MANUAL_MAMBU_STATUS), Cell(amount(vtv, VODAFONE_MANUAL_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(VODAFONE_MANUAL_REFUND_STATUS, STYLE_NOT_FOUND), count(vtv, VODAFONE_MANUAL_REFUND_STATUS), Cell(amount(vtv, VODAFONE_MANUAL_REFUND_STATUS), STYLE_MONEY)],
        [Cell(VODAFONE_MANUAL_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(vtv, VODAFONE_MANUAL_NOT_FOUND_STATUS), Cell(amount(vtv, VODAFONE_MANUAL_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(vtv["total"]), tm(vtv["source_amount"])],
    ]


def status_amount(
    rows: list[OrderedDict[str, Any]],
    *,
    status: str,
    amount_header: str,
) -> Decimal:
    cents = sum(
        amount_to_cents(row.get(amount_header, "")) or 0
        for row in rows
        if str(row.get("Match Status", "")) == status
    )
    return cents_to_decimal(cents)


def category_amount(
    rows: list[OrderedDict[str, Any]],
    *,
    category: str,
    amount_header: str,
    category_header: str,
) -> Decimal:
    cents = sum(
        amount_to_cents(row.get(amount_header, "")) or 0
        for row in rows
        if str(row.get(category_header, "")) == category
    )
    return cents_to_decimal(cents)


def build_itc_wallet_ledger_summary(
    ledger_balance: Decimal,
    manual_unidentified: Decimal,
    month_title: str,
    month_short: str,
    credit_rows: list[OrderedDict[str, Any]],
    debit_rows: list[OrderedDict[str, Any]],
) -> tuple[list[list[Any]], dict[str, Any]]:
    settlement = category_amount(
        credit_rows,
        category=ITC_WALLET_SETTLEMENT_STATUS,
        amount_header="Credit Amount (GHC)",
        category_header="Credit Category",
    )
    reversal = category_amount(
        credit_rows,
        category=ITC_WALLET_REVERSAL_STATUS,
        amount_header="Credit Amount (GHC)",
        category_header="Credit Category",
    )
    transfer_to_wallet = category_amount(
        credit_rows,
        category=ITC_WALLET_TRANSFER_TO_WALLET_STATUS,
        amount_header="Credit Amount (GHC)",
        category_header="Credit Category",
    )
    unidentified_credit = manual_unidentified if manual_unidentified > 0 else Decimal("0")
    unidentified_debit = abs(manual_unidentified) if manual_unidentified < 0 else Decimal("0")
    available_funds = ledger_balance + settlement + reversal + transfer_to_wallet + unidentified_credit

    transfer_to_bank = status_amount(
        debit_rows,
        status=ITC_WALLET_TRANSFERS_FROM_BANK_STATUS,
        amount_header="Debit Amount (GHC)",
    )
    referral_awards = status_amount(
        debit_rows,
        status=ITC_WALLET_REFERRAL_AWARD_STATUS,
        amount_header="Debit Amount (GHC)",
    )
    savings = status_amount(
        debit_rows,
        status=ITC_WALLET_SAVINGS_STATUS,
        amount_header="Debit Amount (GHC)",
    )
    upsales_refund = status_amount(
        debit_rows,
        status=ITC_WALLET_UPSALES_REFUND_STATUS,
        amount_header="Debit Amount (GHC)",
    )
    disbursement = status_amount(
        debit_rows,
        status=ITC_WALLET_DISBURSEMENT_STATUS,
        amount_header="Debit Amount (GHC)",
    )
    total_debit = transfer_to_bank + referral_awards + savings + upsales_refund + disbursement + unidentified_debit
    wallet_statement_balance = available_funds - total_debit

    metrics = {
        "ledger_balance": ledger_balance,
        "manual_unidentified": manual_unidentified,
        "month_title": month_title,
        "month_short": month_short,
        "unidentified_credit": unidentified_credit,
        "unidentified_debit": unidentified_debit,
        "settlement": settlement,
        "reversal": reversal,
        "transfer_to_wallet": transfer_to_wallet,
        "available_funds": available_funds,
        "transfer_to_bank": transfer_to_bank,
        "referral_awards": referral_awards,
        "savings": savings,
        "upsales_refund": upsales_refund,
        "disbursement": disbursement,
        "total_debit": total_debit,
        "wallet_statement_balance": wallet_statement_balance,
    }

    def wh(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_HEADER)

    def ws(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SECTION)

    def wb(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_BALANCE_LABEL)

    def wm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_MONEY)

    def wbm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_BALANCE_MONEY)

    def wf(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_FINAL_LABEL)

    def wfm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_FINAL_MONEY)

    def sig(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SIGNATURE)

    def line(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_LINE_ITEM)

    rows = [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell(f"WALLET RECONCILIATION {month_title}", STYLE_WALLET_TITLE), "", "", ""],
        ["", "", "", ""],
        [wh(f"Date - {month_short}"), wh("Amount (GHC)"), wh("ITC Disbursement"), ""],
        [wb("Balance as per ledger"), wbm(ledger_balance), "", ""],
        [ws("Credit / Available Funds"), "", "", ""],
        [line("Settlement"), wm(settlement), "", ""],
        [line("Reversal"), wm(reversal), "", ""],
        [line("Transfer from Bank"), wm(transfer_to_wallet), "", ""],
    ]
    if unidentified_credit:
        rows.append([line("Unidentified"), wm(unidentified_credit), "", ""])
    rows.extend([
        [wb("Available Funds Before Debit"), wbm(available_funds), "", ""],
        ["", "", "", ""],
        [line("Transfer to Bank"), wm(transfer_to_bank), "", ""],
        [line("Referral Awards"), wm(referral_awards), "", ""],
        [line("Savings"), wm(savings), "", ""],
        [line("Upsales Refund"), wm(upsales_refund), "", ""],
        [line("Disbursement"), wm(disbursement), "", ""],
    ])
    if unidentified_debit:
        rows.append([line("Unidentified"), wm(unidentified_debit), "", ""])
    rows.extend([
        [wb("TOTAL DEBIT"), wbm(total_debit), "", ""],
        ["", "", "", ""],
        [wf("BALANCE AS PER WALLET STATEMENT"), wfm(wallet_statement_balance), "", ""],
        ["", "", "", ""],
        [sig("Prepared By:"), "", sig("Reviewed By:"), ""],
        [sig("Name:"), "", sig("Name:"), ""],
        [sig("Title:"), "", sig("Title:"), ""],
        [sig("Signature:"), "Paste signature image here", sig("Signature:"), "Paste signature image here"],
        [sig("Date:"), "", sig("Date:"), ""],
    ])
    return rows, metrics


def build_nsano_wallet_ledger_summary(
    ledger_balance: Decimal,
    manual_unidentified: Decimal,
    manual_settlement: Decimal,
    manual_reversal: Decimal,
    month_title: str,
    month_short: str,
    mambu_records: list[Record],
    transfer_classification_rows: list[OrderedDict[str, Any]],
    *,
    credit_collection_amount: Decimal | None = None,
    credit_collection_count: int | None = None,
    credit_collection_label: str = "Top-up through collections",
    wallet_column_label: str = "Nsano Disbursement",
    charge_amount: Decimal | None = None,
    charge_count: int | None = None,
    charge_label: str = "Charges",
) -> tuple[list[list[Any]], dict[str, Any]]:
    transfer_topup_collections = category_amount(
        transfer_classification_rows,
        category=NSANO_WALLET_TOPUP_COLLECTION_STATUS,
        amount_header="Transfer Amount (GHC)",
        category_header="Transfer Category",
    )
    topup_collections = (
        credit_collection_amount
        if credit_collection_amount is not None
        else transfer_topup_collections
    )
    topup_collection_count = (
        int(credit_collection_count)
        if credit_collection_count is not None
        else sum(
            1
            for row in transfer_classification_rows
            if str(row.get("Transfer Category", "")) == NSANO_WALLET_TOPUP_COLLECTION_STATUS
        )
    )
    bank_to_wallet = category_amount(
        transfer_classification_rows,
        category=NSANO_WALLET_BANK_TO_WALLET_STATUS,
        amount_header="Transfer Amount (GHC)",
        category_header="Transfer Category",
    )
    reversal_adjustment = category_amount(
        transfer_classification_rows,
        category=NSANO_WALLET_REVERSAL_ADJUSTMENT_STATUS,
        amount_header="Transfer Amount (GHC)",
        category_header="Transfer Category",
    )
    transfer_to_bank = category_amount(
        transfer_classification_rows,
        category=NSANO_WALLET_TRANSFER_TO_BANK_STATUS,
        amount_header="Transfer Amount (GHC)",
        category_header="Transfer Category",
    )
    disbursement = sum_amounts(mambu_records)
    charges = charge_amount if charge_amount is not None else Decimal("0")
    unidentified_available_funds = manual_unidentified if manual_unidentified > 0 else Decimal("0")
    unidentified_debit = abs(manual_unidentified) if manual_unidentified < 0 else Decimal("0")
    reversal_ledger = manual_reversal if manual_reversal > 0 else Decimal("0")
    reversal_available_funds = abs(manual_reversal) if manual_reversal < 0 else Decimal("0")
    starting_ledger_balance = ledger_balance + reversal_ledger
    available_funds = (
        starting_ledger_balance
        + topup_collections
        + bank_to_wallet
        + reversal_adjustment
        + manual_settlement
        + unidentified_available_funds
        + reversal_available_funds
    )
    total_debit = transfer_to_bank + disbursement + charges + unidentified_debit
    wallet_statement_balance = available_funds - total_debit

    metrics = {
        "ledger_balance": ledger_balance,
        "manual_unidentified": manual_unidentified,
        "manual_settlement": manual_settlement,
        "manual_reversal": manual_reversal,
        "month_title": month_title,
        "month_short": month_short,
        "unidentified_available_funds": unidentified_available_funds,
        "unidentified_debit": unidentified_debit,
        "reversal_ledger": reversal_ledger,
        "reversal_available_funds": reversal_available_funds,
        "starting_ledger_balance": starting_ledger_balance,
        "topup_collections": topup_collections,
        "transfer_topup_collections": transfer_topup_collections,
        "topup_collection_count": topup_collection_count,
        "credit_collection_label": credit_collection_label,
        "bank_to_wallet": bank_to_wallet,
        "reversal_adjustment": reversal_adjustment,
        "settlement": manual_settlement,
        "reversal": manual_reversal,
        "available_funds": available_funds,
        "transfer_to_bank": transfer_to_bank,
        "disbursement": disbursement,
        "charges": charges,
        "charge_count": int(charge_count or 0),
        "charge_label": charge_label,
        "total_debit": total_debit,
        "wallet_statement_balance": wallet_statement_balance,
    }

    def wh(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_HEADER)

    def ws(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SECTION)

    def wb(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_BALANCE_LABEL)

    def wm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_MONEY)

    def wbm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_BALANCE_MONEY)

    def wf(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_FINAL_LABEL)

    def wfm(value: Any) -> Cell:
        return Cell(value, STYLE_WALLET_FINAL_MONEY)

    def sig(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_SIGNATURE)

    def line(label: str) -> Cell:
        return Cell(label, STYLE_WALLET_LINE_ITEM)

    rows = [
        [Cell("FIDO MICRO CREDIT LTD", STYLE_WALLET_COMPANY), "", "", ""],
        [Cell(f"WALLET RECONCILIATION {month_title}", STYLE_WALLET_TITLE), "", "", ""],
        ["", "", "", ""],
        [wh(f"Date - {month_short}"), wh("Amount (GHC)"), wh(wallet_column_label), ""],
        [wb("Balance as per ledger"), wbm(ledger_balance), "", ""],
    ]
    if reversal_ledger:
        rows.append([line("Reversal"), wm(reversal_ledger), "", ""])
    rows.extend([
        [ws("Credit / Available Funds"), "", "", ""],
        [line(credit_collection_label), wm(topup_collections), "", ""],
        [line("Bank to Wallet"), wm(bank_to_wallet), "", ""],
        [line("Reversal Adjustment"), wm(reversal_adjustment), "", ""],
        [line("Recovery from write off"), wm(manual_settlement), "", ""],
    ])
    if unidentified_available_funds:
        rows.append([line("Delayed Transactions"), wm(unidentified_available_funds), "", ""])
    if reversal_available_funds:
        rows.append([line("Reversal"), wm(reversal_available_funds), "", ""])
    rows.extend([
        [wb("Available Funds Before Debit"), wbm(available_funds), "", ""],
        ["", "", "", ""],
        [line("Transfer to Bank"), wm(transfer_to_bank), "", ""],
        [line("Disbursement"), wm(disbursement), "", ""],
    ])
    if charge_amount is not None:
        rows.append([line(charge_label), wm(charges), "", ""])
    if unidentified_debit:
        rows.append([line("Delayed Transactions"), wm(unidentified_debit), "", ""])
    rows.extend([
        [wb("TOTAL DEBIT"), wbm(total_debit), "", ""],
        ["", "", "", ""],
        [wf("BALANCE AS PER WALLET STATEMENT"), wfm(wallet_statement_balance), "", ""],
        ["", "", "", ""],
        [sig("Prepared By:"), "", sig("Reviewed By:"), ""],
        [sig("Name:"), "", sig("Name:"), ""],
        [sig("Title:"), "", sig("Title:"), ""],
        [sig("Signature:"), "Paste signature image here", sig("Signature:"), "Paste signature image here"],
        [sig("Date:"), "", sig("Date:"), ""],
    ])
    return rows, metrics


def instruction_rows() -> list[list[Any]]:
    return [
        [Cell("NSANO DISBURSEMENT RECONCILIATION TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile Nsano disbursement transactions against Mambu records from both directions.",
        ],
        [
            "Matching rule",
            "Column I is used as the identifier on both source sheets. Nsano identifiers are matched after removing a leading FIDO prefix.",
        ],
        [Cell("Data sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu",
            "Paste or upload the Mambu disbursement export. Column I must contain the Mambu Identifier.",
        ],
        [
            "NSANO DISB",
            "Paste or upload the Nsano disbursement export. Column I must contain SendingHse_ID or SendingHse_ID (KEY).",
        ],
        [Cell("Output sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu vs Nsano",
            "Mambu is the source. Match Status is nsano when the identifier is found in Nsano, otherwise not found.",
        ],
        [
            "Nsano vs Mambu",
            "Nsano is the source. Match Status is mambu when the stripped identifier is found in Mambu, otherwise not found.",
        ],
    ]


def itc_disb_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("ITC WALLET VS MAMBU TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile ITC wallet disbursement transactions against Mambu records from both directions.",
        ],
        [
            "Matching rule",
            "ITC Wallet column G is matched directly against Mambu column I.",
        ],
        [Cell("Data sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu",
            "Paste or upload the Mambu disbursement export. Column I must contain the Mambu Identifier.",
        ],
        [
            "ITC Wallet",
            "Paste or upload the ITC wallet export. Column G must contain thirdparty_id.",
        ],
        [Cell("Output sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu vs ITC Wallet",
            "Mambu is the source. Match Status is ITC when the identifier is found in ITC Wallet, otherwise not_found.",
        ],
        [
            "ITC Wallet vs Mambu",
            "ITC Wallet is the source. Match Status is mambu when the identifier is found in Mambu. Unmatched _2, _3, and _6 narrations are classified as referral bonus, upsales refund, and savings reward.",
        ],
    ]


def itc_wallet_ledger_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("ITC WALLET VS LEDGER TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile ITC wallet movement against a manually entered ledger balance.",
        ],
        [
            "Credit transfers",
            "Column G narration is classified by code into settlement, reversal, prepaid reversal, or transfer to wallet.",
        ],
        [
            "Debit transfers vs statement",
            "Debit Transfers column C is matched to ITC Statement column A. Statement narration suffixes classify disbursement, referral awards, upsales refund, and savings.",
        ],
        [
            "Fallback",
            "Unmatched debit transfers with withdrawal are transfers from bank; otherwise debit column C is matched to Credit Transfers column C as credit transfer (prepaid_reversal).",
        ],
        [
            "Summary",
            "Ledger balance and unidentified are entered manually in the UI. Positive unidentified is added to Credit / Available Funds; negative unidentified is shown as a positive debit.",
        ],
    ]


def nsano_wallet_ledger_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("NSANO DISB WALLET VS LEDGER TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile Nsano disbursement wallet movement against a manually entered ledger balance.",
        ],
        [
            "Nsano Transfers",
            "Column E Purpose is classified into Top-up through collections, Bank to Wallet, Reversal Adjustment, and Transfer to Bank. Approved collection spelling variations are included and marked in the Narration Exception columns.",
        ],
        [
            "Mambu",
            "Mambu column E is used as the disbursement amount in the Total Debit section.",
        ],
        [
            "Nsano Disbursement",
            "Uploaded as the Nsano wallet disbursement statement and retained as a source sheet in the output workbook.",
        ],
        [
            "Manual fields",
            "Ledger balance, delayed transactions, recovery from write off, and reversal are entered in the UI. Positive delayed transactions is added to Credit / Available Funds; negative delayed transactions is shown as a positive debit line. Positive reversal is shown with ledger balance; negative reversal is shown as a positive available-funds line.",
        ],
    ]


def nsano_collection_ledger_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("NSANO COLLECTIONS VS LEDGER TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile Nsano collections wallet movement against a manually entered ledger balance.",
        ],
        [
            "Mambu collections",
            "Filtered Mambu Collection upload rows with an Nsano collection channel are summed as Total Nsano Mambu Collections in Credit / Available Funds.",
        ],
        [
            "Mambu disbursements",
            "Filtered Mambu Disbursement upload rows with an Nsano disbursement channel are summed in the Disbursement debit line.",
        ],
        [
            "Nsano Transfers",
            "Column E Purpose is classified into Bank to Wallet, Reversal Adjustment, and Transfer to Bank. Top-up rows, including approved spelling variations, are retained for audit and variations are marked in the Narration Exception columns; they are not used as the collections credit line.",
        ],
        [
            "Nsano Filtered Data (optional)",
            "Successful W2A export. Column U (Charge) is summed across rows with a non-zero charge and added to the Debit section as Charges, included in TOTAL DEBIT. Leave unuploaded to skip this line.",
        ],
        [
            "Manual fields",
            "Ledger balance, delayed transactions, recovery from write off, and reversal are entered in the UI. Positive delayed transactions is added to Credit / Available Funds; negative delayed transactions is shown as a positive debit line. Positive reversal is shown with ledger balance; negative reversal is shown as a positive available-funds line.",
        ],
    ]


def mtn_manual_disb_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("MTN MANUAL DISBURSEMENT RECONCILIATION TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile MTN Manual disbursement transactions against Mambu, with Refund as fallback.",
        ],
        [
            "Mambu vs MTN MANUAL",
            "Mambu column I is matched against MTN MANUAL column A. Match Status is Matched when found, otherwise Not Found.",
        ],
        [
            "MTN MANUAL as Source",
            "MTN MANUAL column A is matched against Mambu column I first. If not found, REFUND column D is checked.",
        ],
        [Cell("Data sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu",
            "Paste or upload the Mambu disbursement export. Column I must contain Identifier.",
        ],
        [
            "MTN MANUAL",
            "Paste or upload the MTN Manual export. Column A must contain Id.",
        ],
        [
            "REFUND",
            "Paste or upload the refund data. Column D must contain TRANSACTION ID.",
        ],
    ]


def vodafone_manual_disb_instruction_rows() -> list[list[Any]]:
    return [
        [Cell("VODAFONE MANUAL DISBURSEMENT RECONCILIATION TEMPLATE - INSTRUCTIONS", STYLE_TITLE), ""],
        [Cell("Overview", STYLE_TABLE_SECTION), ""],
        [
            "Purpose",
            "Reconcile Vodafone Manual disbursement transactions against Mambu, with Refund as fallback.",
        ],
        [
            "Mambu vs VODAFONE MANUAL",
            "Mambu column I is matched against VODAFONE MANUAL column A. Match Status is Matched when found, otherwise Not Found.",
        ],
        [
            "VODAFONE MANUAL as Source",
            "VODAFONE MANUAL column A is matched against Mambu column I first. If not found, REFUND column D is checked.",
        ],
        [Cell("Data sheets", STYLE_TABLE_SECTION), ""],
        [
            "Mambu",
            "Paste or upload the Mambu disbursement export. Column I must contain Identifier.",
        ],
        [
            "VODAFONE MANUAL",
            "Paste or upload the Vodafone Manual export. Column A must contain Id.",
        ],
        [
            "REFUND",
            "Paste or upload the refund data. Column D must contain TRANSACTION ID.",
        ],
    ]


def disb_compare_style(headers: list[str]):
    amount_columns = {
        "Mambu Amount (GHC)",
        "Nsano Amount (GHC)",
        "Amount",
        "Amount_GHC",
        "Amount (GHC)",
        "Principal Amount",
        "balanceBefore",
        "balanceAfter",
        "Charge",
        "Transfer Amount (GHC)",
        "TOP-UP Through Collection with Date",
        "Bank to Wallet",
        "Reversal Adjustment",
        "Transfer to Bank",
    }

    def style(row_number: int, col_number: int, value: Any) -> int | None:
        if row_number == 1:
            return STYLE_HEADER
        header = headers[col_number - 1] if col_number - 1 < len(headers) else ""
        if header == "Match Status":
            if value in {
                DISB_MAMBU_STATUS,
                DISB_NSANO_STATUS,
                ITC_DISB_STATUS,
                MTN_MANUAL_MATCHED_STATUS,
                MTN_MANUAL_MAMBU_STATUS,
                VODAFONE_MANUAL_MATCHED_STATUS,
                VODAFONE_MANUAL_MAMBU_STATUS,
                ITC_WALLET_DISBURSEMENT_STATUS,
                ITC_WALLET_REFERRAL_AWARD_STATUS,
                ITC_WALLET_UPSALES_REFUND_STATUS,
                ITC_WALLET_SAVINGS_STATUS,
                ITC_WALLET_TRANSFERS_FROM_BANK_STATUS,
                ITC_WALLET_CREDIT_TRANSFER_STATUS,
            }:
                return STYLE_MATCHED
            return STYLE_NOT_FOUND
        if header in amount_columns:
            return STYLE_MONEY
        return None

    return style


def rows_from_dicts(headers: list[str], rows: list[OrderedDict[str, Any]]) -> Iterable[list[Any]]:
    yield headers
    for row in rows:
        yield [row.get(header, "") for header in headers]


def cell_plain_value(value: Any) -> Any:
    return value.value if isinstance(value, Cell) else value


def itc_wallet_ledger_summary_merges(rows: list[list[Any]]) -> list[str]:
    merges: list[str] = []
    for row_number, row in enumerate(rows, start=1):
        first = cell_plain_value(row[0]) if len(row) >= 1 else ""
        second = cell_plain_value(row[1]) if len(row) >= 2 else ""
        third = cell_plain_value(row[2]) if len(row) >= 3 else ""
        if isinstance(row[0], Cell) and row[0].style in {STYLE_WALLET_COMPANY, STYLE_WALLET_TITLE}:
            merges.append(f"A{row_number}:D{row_number}")
        elif str(first).startswith("Date -") and third in {"ITC Disbursement", "Nsano Disbursement", "Nsano Collections"}:
            if second != "Amount (GHC)":
                merges.append(f"A{row_number}:B{row_number}")
            merges.append(f"C{row_number}:D{row_number}")
        elif first == "Prepared By:" and third == "Reviewed By:":
            merges.append(f"A{row_number}:B{row_number}")
            merges.append(f"C{row_number}:D{row_number}")
    return merges


def write_nsano_disb_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    nsano_headers: list[str],
    nsano_records: list[Record],
    mambu_vs_nsano_rows: list[OrderedDict[str, Any]],
    nsano_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Mambu vs Nsano",
        "Nsano vs Mambu",
        "Mambu",
        "NSANO DISB",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    nsano_preview = ([record.sheet_values.get(header, "") for header in nsano_headers] for record in nsano_records[:5000])
    mambu_vs_preview = ([row.get(header, "") for header in MAMBU_VS_NSANO_HEADERS] for row in mambu_vs_nsano_rows[:5000])
    nsano_vs_preview = ([row.get(header, "") for header in NSANO_VS_MAMBU_HEADERS] for row in nsano_vs_mambu_rows[:5000])
    instructions = instruction_rows()

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
            instructions,
            len(instructions),
            2,
            [34, 86],
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
            [32, 14, 22],
            freeze_top_row=False,
            autofilter=False,
            merges=summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(MAMBU_VS_NSANO_HEADERS, mambu_vs_nsano_rows),
            len(mambu_vs_nsano_rows) + 1,
            len(MAMBU_VS_NSANO_HEADERS),
            compute_widths(MAMBU_VS_NSANO_HEADERS, mambu_vs_preview),
            style_func=disb_compare_style(MAMBU_VS_NSANO_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(NSANO_VS_MAMBU_HEADERS, nsano_vs_mambu_rows),
            len(nsano_vs_mambu_rows) + 1,
            len(NSANO_VS_MAMBU_HEADERS),
            compute_widths(NSANO_VS_MAMBU_HEADERS, nsano_vs_preview),
            style_func=disb_compare_style(NSANO_VS_MAMBU_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(nsano_headers, nsano_records),
            len(nsano_records) + 1,
            len(nsano_headers),
            compute_widths(nsano_headers, nsano_preview),
            style_func=data_style(nsano_headers),
        )

    os.replace(temp_path, output_path)


def write_itc_disb_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    itc_headers: list[str],
    itc_records: list[Record],
    mambu_vs_itc_rows: list[OrderedDict[str, Any]],
    itc_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Mambu vs ITC Wallet",
        "ITC Wallet vs Mambu",
        "Mambu",
        "ITC Wallet",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    itc_preview = ([record.sheet_values.get(header, "") for header in itc_headers] for record in itc_records[:5000])
    mambu_vs_preview = ([row.get(header, "") for header in MAMBU_VS_ITC_DISB_HEADERS] for row in mambu_vs_itc_rows[:5000])
    itc_vs_preview = ([row.get(header, "") for header in ITC_DISB_VS_MAMBU_HEADERS] for row in itc_vs_mambu_rows[:5000])
    instructions = itc_disb_instruction_rows()

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
            instructions,
            len(instructions),
            2,
            [34, 86],
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
            [38, 14, 22],
            freeze_top_row=False,
            autofilter=False,
            merges=summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(MAMBU_VS_ITC_DISB_HEADERS, mambu_vs_itc_rows),
            len(mambu_vs_itc_rows) + 1,
            len(MAMBU_VS_ITC_DISB_HEADERS),
            compute_widths(MAMBU_VS_ITC_DISB_HEADERS, mambu_vs_preview),
            style_func=disb_compare_style(MAMBU_VS_ITC_DISB_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(ITC_DISB_VS_MAMBU_HEADERS, itc_vs_mambu_rows),
            len(itc_vs_mambu_rows) + 1,
            len(ITC_DISB_VS_MAMBU_HEADERS),
            compute_widths(ITC_DISB_VS_MAMBU_HEADERS, itc_vs_preview),
            style_func=disb_compare_style(ITC_DISB_VS_MAMBU_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(itc_headers, itc_records),
            len(itc_records) + 1,
            len(itc_headers),
            compute_widths(itc_headers, itc_preview),
            style_func=data_style(itc_headers),
        )

    os.replace(temp_path, output_path)


def write_itc_wallet_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    statement_headers: list[str],
    statement_rows: list[OrderedDict[str, Any]],
    debit_headers: list[str],
    debit_rows: list[OrderedDict[str, Any]],
    credit_headers: list[str],
    credit_rows: list[OrderedDict[str, Any]],
    credit_classification_rows: list[OrderedDict[str, Any]],
    debit_comparison_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "ITC Credit Transfers",
        "Debit Transfers vs Statement",
        "ITC Not Found",
        "ITC Statement",
        "ITC Debit Transfers",
        "Raw Credit Transfers",
    ]

    instructions = itc_wallet_ledger_instruction_rows()
    statement_preview = (
        [row.get(header, "") for header in statement_headers]
        for row in statement_rows[:5000]
    )
    debit_preview = (
        [row.get(header, "") for header in debit_headers]
        for row in debit_rows[:5000]
    )
    credit_preview = (
        [row.get(header, "") for header in credit_headers]
        for row in credit_rows[:5000]
    )
    credit_classification_preview = (
        [row.get(header, "") for header in ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS]
        for row in credit_classification_rows[:5000]
    )
    debit_comparison_preview = (
        [row.get(header, "") for header in ITC_WALLET_DEBIT_COMPARISON_HEADERS]
        for row in debit_comparison_rows[:5000]
    )
    not_found_rows = [
        row for row in debit_comparison_rows
        if str(row.get("Match Status", "")) == ITC_WALLET_NOT_FOUND_STATUS
    ]
    not_found_preview = (
        [row.get(header, "") for header in ITC_WALLET_DEBIT_COMPARISON_HEADERS]
        for row in not_found_rows[:5000]
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
            instructions,
            len(instructions),
            2,
            [34, 110],
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
            4,
            [36, 22, 28, 28],
            freeze_top_row=False,
            autofilter=False,
            merges=itc_wallet_ledger_summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS, credit_classification_rows),
            len(credit_classification_rows) + 1,
            len(ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS),
            compute_widths(ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS, credit_classification_preview),
            style_func=disb_compare_style(ITC_WALLET_CREDIT_CLASSIFICATION_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(ITC_WALLET_DEBIT_COMPARISON_HEADERS, debit_comparison_rows),
            len(debit_comparison_rows) + 1,
            len(ITC_WALLET_DEBIT_COMPARISON_HEADERS),
            compute_widths(ITC_WALLET_DEBIT_COMPARISON_HEADERS, debit_comparison_preview),
            style_func=disb_compare_style(ITC_WALLET_DEBIT_COMPARISON_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            rows_from_dicts(ITC_WALLET_DEBIT_COMPARISON_HEADERS, not_found_rows),
            len(not_found_rows) + 1,
            len(ITC_WALLET_DEBIT_COMPARISON_HEADERS),
            compute_widths(ITC_WALLET_DEBIT_COMPARISON_HEADERS, not_found_preview),
            style_func=disb_compare_style(ITC_WALLET_DEBIT_COMPARISON_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            rows_from_dicts(statement_headers, statement_rows),
            len(statement_rows) + 1,
            len(statement_headers),
            compute_widths(statement_headers, statement_preview),
            style_func=data_style(statement_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            rows_from_dicts(debit_headers, debit_rows),
            len(debit_rows) + 1,
            len(debit_headers),
            compute_widths(debit_headers, debit_preview),
            style_func=data_style(debit_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet8.xml",
            rows_from_dicts(credit_headers, credit_rows),
            len(credit_rows) + 1,
            len(credit_headers),
            compute_widths(credit_headers, credit_preview),
            style_func=data_style(credit_headers),
        )

    os.replace(temp_path, output_path)


def write_nsano_wallet_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    nsano_disb_headers: list[str],
    nsano_disb_records: list[Record],
    transfer_headers: list[str],
    transfer_rows: list[OrderedDict[str, Any]],
    transfer_classification_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Nsano Transfers Classified",
        "Nsano Not Found",
        "Mambu",
        "Nsano Disbursement",
        "Nsano Transfers",
    ]

    instructions = nsano_wallet_ledger_instruction_rows()
    mambu_preview = (
        [record.sheet_values.get(header, "") for header in mambu_headers]
        for record in mambu_records[:5000]
    )
    nsano_disb_preview = (
        [record.sheet_values.get(header, "") for header in nsano_disb_headers]
        for record in nsano_disb_records[:5000]
    )
    transfer_preview = (
        [row.get(header, "") for header in transfer_headers]
        for row in transfer_rows[:5000]
    )
    classification_preview = (
        [row.get(header, "") for header in NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS]
        for row in transfer_classification_rows[:5000]
    )
    not_found_rows = [
        row for row in transfer_classification_rows
        if str(row.get("Transfer Category", "")) == NSANO_WALLET_OTHER_STATUS
    ]
    not_found_preview = (
        [row.get(header, "") for header in NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS]
        for row in not_found_rows[:5000]
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
            instructions,
            len(instructions),
            2,
            [34, 110],
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
            4,
            [36, 22, 28, 28],
            freeze_top_row=False,
            autofilter=False,
            merges=itc_wallet_ledger_summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, transfer_classification_rows),
            len(transfer_classification_rows) + 1,
            len(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
            compute_widths(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, classification_preview),
            style_func=disb_compare_style(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, not_found_rows),
            len(not_found_rows) + 1,
            len(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
            compute_widths(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, not_found_preview),
            style_func=disb_compare_style(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(nsano_disb_headers, nsano_disb_records),
            len(nsano_disb_records) + 1,
            len(nsano_disb_headers),
            compute_widths(nsano_disb_headers, nsano_disb_preview),
            style_func=data_style(nsano_disb_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            rows_from_dicts(transfer_headers, transfer_rows),
            len(transfer_rows) + 1,
            len(transfer_headers),
            compute_widths(transfer_headers, transfer_preview),
            style_func=data_style(transfer_headers),
        )

    os.replace(temp_path, output_path)


def write_nsano_collection_ledger_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_collection_headers: list[str],
    mambu_collection_records: list[Record],
    mambu_disb_headers: list[str],
    mambu_disb_records: list[Record],
    nsano_charge_headers: list[str],
    nsano_charge_records: list[Record],
    transfer_headers: list[str],
    transfer_rows: list[OrderedDict[str, Any]],
    transfer_classification_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    mambu_collection_headers = mambu_collection_headers or list(NSANO_COLLECTION_MAMBU_SHEET_COLUMNS)
    mambu_disb_headers = mambu_disb_headers or list(NSANO_DISB_MAMBU_SHEET_COLUMNS)
    nsano_charge_headers = nsano_charge_headers or list(NSANO_DISB_SHEET_COLUMNS)
    transfer_headers = transfer_headers or list(NSANO_WALLET_TRANSFER_SHEET_COLUMNS)
    sheet_names = [
        "Instructions",
        "Summary",
        "Nsano Transfers Classified",
        "Nsano Not Found",
        "Mambu Collections",
        "Mambu Disbursements",
        "Nsano Charges",
        "Nsano Transfers",
    ]

    instructions = nsano_collection_ledger_instruction_rows()
    mambu_collection_preview = (
        [record.sheet_values.get(header, "") for header in mambu_collection_headers]
        for record in mambu_collection_records[:5000]
    )
    mambu_disb_preview = (
        [record.sheet_values.get(header, "") for header in mambu_disb_headers]
        for record in mambu_disb_records[:5000]
    )
    nsano_charge_preview = (
        [record.sheet_values.get(header, "") for header in nsano_charge_headers]
        for record in nsano_charge_records[:5000]
    )
    transfer_preview = (
        [row.get(header, "") for header in transfer_headers]
        for row in transfer_rows[:5000]
    )
    classification_preview = (
        [row.get(header, "") for header in NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS]
        for row in transfer_classification_rows[:5000]
    )
    not_found_rows = [
        row for row in transfer_classification_rows
        if str(row.get("Transfer Category", "")) == NSANO_WALLET_OTHER_STATUS
    ]
    not_found_preview = (
        [row.get(header, "") for header in NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS]
        for row in not_found_rows[:5000]
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
            instructions,
            len(instructions),
            2,
            [34, 110],
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
            4,
            [36, 22, 28, 28],
            freeze_top_row=False,
            autofilter=False,
            merges=itc_wallet_ledger_summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, transfer_classification_rows),
            len(transfer_classification_rows) + 1,
            len(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
            compute_widths(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, classification_preview),
            style_func=disb_compare_style(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, not_found_rows),
            len(not_found_rows) + 1,
            len(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
            compute_widths(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS, not_found_preview),
            style_func=disb_compare_style(NSANO_WALLET_TRANSFER_CLASSIFICATION_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_collection_headers, mambu_collection_records),
            len(mambu_collection_records) + 1,
            len(mambu_collection_headers),
            compute_widths(mambu_collection_headers, mambu_collection_preview),
            style_func=data_style(mambu_collection_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(mambu_disb_headers, mambu_disb_records),
            len(mambu_disb_records) + 1,
            len(mambu_disb_headers),
            compute_widths(mambu_disb_headers, mambu_disb_preview),
            style_func=data_style(mambu_disb_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            record_rows(nsano_charge_headers, nsano_charge_records),
            len(nsano_charge_records) + 1,
            len(nsano_charge_headers),
            compute_widths(nsano_charge_headers, nsano_charge_preview),
            style_func=data_style(nsano_charge_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet8.xml",
            rows_from_dicts(transfer_headers, transfer_rows),
            len(transfer_rows) + 1,
            len(transfer_headers),
            compute_widths(transfer_headers, transfer_preview),
            style_func=data_style(transfer_headers),
        )

    os.replace(temp_path, output_path)


def write_mtn_manual_disb_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    mtn_headers: list[str],
    mtn_records: list[Record],
    refund_headers: list[str],
    refund_records: list[Record],
    mambu_vs_mtn_rows: list[OrderedDict[str, Any]],
    mtn_vs_sources_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Mambu vs MTN MANUAL",
        "MTN MANUAL as Source",
        "Mambu",
        "MTN MANUAL",
        "REFUND",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    mtn_preview = ([record.sheet_values.get(header, "") for header in mtn_headers] for record in mtn_records[:5000])
    refund_preview = ([record.sheet_values.get(header, "") for header in refund_headers] for record in refund_records[:5000])
    mambu_vs_preview = ([row.get(header, "") for header in MAMBU_VS_MTN_MANUAL_HEADERS] for row in mambu_vs_mtn_rows[:5000])
    mtn_vs_preview = ([row.get(header, "") for header in MTN_MANUAL_AS_SOURCE_HEADERS] for row in mtn_vs_sources_rows[:5000])
    instructions = mtn_manual_disb_instruction_rows()

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
            instructions,
            len(instructions),
            2,
            [34, 94],
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
            rows_from_dicts(MAMBU_VS_MTN_MANUAL_HEADERS, mambu_vs_mtn_rows),
            len(mambu_vs_mtn_rows) + 1,
            len(MAMBU_VS_MTN_MANUAL_HEADERS),
            compute_widths(MAMBU_VS_MTN_MANUAL_HEADERS, mambu_vs_preview),
            style_func=disb_compare_style(MAMBU_VS_MTN_MANUAL_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(MTN_MANUAL_AS_SOURCE_HEADERS, mtn_vs_sources_rows),
            len(mtn_vs_sources_rows) + 1,
            len(MTN_MANUAL_AS_SOURCE_HEADERS),
            compute_widths(MTN_MANUAL_AS_SOURCE_HEADERS, mtn_vs_preview),
            style_func=disb_compare_style(MTN_MANUAL_AS_SOURCE_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(mtn_headers, mtn_records),
            len(mtn_records) + 1,
            len(mtn_headers),
            compute_widths(mtn_headers, mtn_preview),
            style_func=data_style(mtn_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            record_rows(refund_headers, refund_records),
            len(refund_records) + 1,
            len(refund_headers),
            compute_widths(refund_headers, refund_preview),
            style_func=data_style(refund_headers),
        )

    os.replace(temp_path, output_path)


def write_vodafone_manual_disb_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    vodafone_headers: list[str],
    vodafone_records: list[Record],
    refund_headers: list[str],
    refund_records: list[Record],
    mambu_vs_vodafone_rows: list[OrderedDict[str, Any]],
    vodafone_vs_sources_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Instructions",
        "Summary",
        "Mambu vs VODAFONE MANUAL",
        "VODAFONE MANUAL as Source",
        "Mambu",
        "VODAFONE MANUAL",
        "REFUND",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    vodafone_preview = ([record.sheet_values.get(header, "") for header in vodafone_headers] for record in vodafone_records[:5000])
    refund_preview = ([record.sheet_values.get(header, "") for header in refund_headers] for record in refund_records[:5000])
    mambu_vs_preview = ([row.get(header, "") for header in MAMBU_VS_VODAFONE_MANUAL_HEADERS] for row in mambu_vs_vodafone_rows[:5000])
    vodafone_vs_preview = ([row.get(header, "") for header in VODAFONE_MANUAL_AS_SOURCE_HEADERS] for row in vodafone_vs_sources_rows[:5000])
    instructions = vodafone_manual_disb_instruction_rows()

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
            instructions,
            len(instructions),
            2,
            [34, 98],
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
            [38, 14, 22],
            freeze_top_row=False,
            autofilter=False,
            merges=summary_merges(summary_rows),
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(MAMBU_VS_VODAFONE_MANUAL_HEADERS, mambu_vs_vodafone_rows),
            len(mambu_vs_vodafone_rows) + 1,
            len(MAMBU_VS_VODAFONE_MANUAL_HEADERS),
            compute_widths(MAMBU_VS_VODAFONE_MANUAL_HEADERS, mambu_vs_preview),
            style_func=disb_compare_style(MAMBU_VS_VODAFONE_MANUAL_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(VODAFONE_MANUAL_AS_SOURCE_HEADERS, vodafone_vs_sources_rows),
            len(vodafone_vs_sources_rows) + 1,
            len(VODAFONE_MANUAL_AS_SOURCE_HEADERS),
            compute_widths(VODAFONE_MANUAL_AS_SOURCE_HEADERS, vodafone_vs_preview),
            style_func=disb_compare_style(VODAFONE_MANUAL_AS_SOURCE_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(vodafone_headers, vodafone_records),
            len(vodafone_records) + 1,
            len(vodafone_headers),
            compute_widths(vodafone_headers, vodafone_preview),
            style_func=data_style(vodafone_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            record_rows(refund_headers, refund_records),
            len(refund_records) + 1,
            len(refund_headers),
            compute_widths(refund_headers, refund_preview),
            style_func=data_style(refund_headers),
        )

    os.replace(temp_path, output_path)


def build_nsano_disb_reconciliation(
    output_path: Path,
    mambu_records: list[Record],
    mambu_headers: list[str],
    nsano_records: list[Record],
    nsano_headers: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    mambu_vs_nsano_rows = compare_mambu_to_nsano_disb(mambu_records, nsano_records)
    nsano_vs_mambu_rows = compare_nsano_to_mambu_disb(nsano_records, mambu_records)
    mambu_vs_nsano_stats = summarize_disb_compare(
        mambu_vs_nsano_rows,
        matched_status=DISB_NSANO_STATUS,
        amount_header="Mambu Amount (GHC)",
    )
    nsano_vs_mambu_stats = summarize_disb_compare(
        nsano_vs_mambu_rows,
        matched_status=DISB_MAMBU_STATUS,
        amount_header="Nsano Amount (GHC)",
    )
    mambu_vs_nsano_stats["not_found_details"] = not_found_detail_rows(
        mambu_records,
        mambu_vs_nsano_rows,
        DISB_NOT_FOUND_STATUS,
        "Mambu",
        "Mambu → Wallet",
    )
    nsano_vs_mambu_stats["not_found_details"] = not_found_detail_rows(
        nsano_records,
        nsano_vs_mambu_rows,
        DISB_NOT_FOUND_STATUS,
        "Nsano",
        "Wallet → Mambu",
    )
    summary_rows = build_nsano_disb_summary_rows(mambu_vs_nsano_stats, nsano_vs_mambu_stats)
    write_nsano_disb_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        nsano_headers,
        nsano_records,
        mambu_vs_nsano_rows,
        nsano_vs_mambu_rows,
    )
    return mambu_vs_nsano_stats, nsano_vs_mambu_stats


def build_itc_disb_reconciliation(
    output_path: Path,
    mambu_records: list[Record],
    mambu_headers: list[str],
    itc_records: list[Record],
    itc_headers: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    mambu_vs_itc_rows = compare_mambu_to_itc_disb(mambu_records, itc_records)
    itc_vs_mambu_rows = compare_itc_disb_to_mambu(itc_records, mambu_records)
    mambu_vs_itc_stats = summarize_disb_compare(
        mambu_vs_itc_rows,
        matched_status=ITC_DISB_STATUS,
        amount_header="Mambu Amount (GHC)",
        not_found_status=ITC_DISB_NOT_FOUND_STATUS,
    )
    itc_vs_mambu_stats = summarize_disb_compare(
        itc_vs_mambu_rows,
        matched_status=DISB_MAMBU_STATUS,
        amount_header="ITC Amount (GHC)",
        not_found_status=ITC_DISB_NOT_FOUND_STATUS,
    )
    mambu_vs_itc_stats["not_found_details"] = not_found_detail_rows(
        mambu_records,
        mambu_vs_itc_rows,
        ITC_DISB_NOT_FOUND_STATUS,
        "Mambu",
        "Mambu → Wallet",
    )
    itc_vs_mambu_stats["not_found_details"] = not_found_detail_rows(
        itc_records,
        itc_vs_mambu_rows,
        ITC_DISB_NOT_FOUND_STATUS,
        "ITC",
        "Wallet → Mambu",
    )
    itc_breakdowns = summarize_itc_disb_narration_breakdowns(itc_vs_mambu_rows)
    summary_rows = build_itc_disb_summary_rows(mambu_vs_itc_stats, itc_vs_mambu_stats, itc_breakdowns)
    write_itc_disb_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        itc_headers,
        itc_records,
        mambu_vs_itc_rows,
        itc_vs_mambu_rows,
    )
    return mambu_vs_itc_stats, itc_vs_mambu_stats


def build_itc_wallet_ledger_reconciliation(
    output_path: Path,
    ledger_balance: Decimal,
    manual_unidentified: Decimal,
    statement_rows: list[OrderedDict[str, Any]],
    statement_headers: list[str],
    debit_rows: list[OrderedDict[str, Any]],
    debit_headers: list[str],
    credit_rows: list[OrderedDict[str, Any]],
    credit_headers: list[str],
) -> dict[str, Any]:
    credit_classification_rows = build_itc_credit_transfer_rows(credit_rows)
    debit_comparison_rows = compare_itc_debit_transfers_to_statement(
        debit_rows,
        statement_rows,
        credit_rows,
    )
    month_title, month_short = detect_itc_wallet_month(statement_rows, debit_rows, credit_rows)
    summary_rows, metrics = build_itc_wallet_ledger_summary(
        ledger_balance,
        manual_unidentified,
        month_title,
        month_short,
        credit_classification_rows,
        debit_comparison_rows,
    )

    debit_status_counts = Counter(str(row.get("Match Status", "")) for row in debit_comparison_rows)
    credit_category_counts = Counter(str(row.get("Credit Category", "")) for row in credit_classification_rows)
    debit_status_amounts = {
        status: status_amount(debit_comparison_rows, status=status, amount_header="Debit Amount (GHC)")
        for status in debit_status_counts
    }
    credit_category_amounts = {
        category: category_amount(
            credit_classification_rows,
            category=category,
            amount_header="Credit Amount (GHC)",
            category_header="Credit Category",
        )
        for category in credit_category_counts
    }

    write_itc_wallet_ledger_workbook(
        output_path,
        summary_rows,
        statement_headers,
        statement_rows,
        debit_headers,
        debit_rows,
        credit_headers,
        credit_rows,
        credit_classification_rows,
        debit_comparison_rows,
    )

    return {
        **metrics,
        "statement_total": len(statement_rows),
        "debit_total": len(debit_rows),
        "credit_total": len(credit_rows),
        "debit_status_counts": dict(debit_status_counts),
        "credit_category_counts": dict(credit_category_counts),
        "debit_status_amounts": debit_status_amounts,
        "credit_category_amounts": credit_category_amounts,
    }


def build_nsano_wallet_ledger_reconciliation(
    output_path: Path,
    ledger_balance: Decimal,
    manual_unidentified: Decimal,
    manual_settlement: Decimal,
    manual_reversal: Decimal,
    mambu_records: list[Record],
    mambu_headers: list[str],
    nsano_disb_records: list[Record],
    nsano_disb_headers: list[str],
    transfer_rows: list[OrderedDict[str, Any]],
    transfer_headers: list[str],
) -> dict[str, Any]:
    transfer_classification_rows = build_nsano_transfer_classification_rows(transfer_rows)
    month_title, month_short = detect_nsano_wallet_month(mambu_records, transfer_rows)
    summary_rows, metrics = build_nsano_wallet_ledger_summary(
        ledger_balance,
        manual_unidentified,
        manual_settlement,
        manual_reversal,
        month_title,
        month_short,
        mambu_records,
        transfer_classification_rows,
    )

    transfer_category_counts = Counter(
        str(row.get("Transfer Category", ""))
        for row in transfer_classification_rows
    )
    transfer_category_amounts = {
        category: category_amount(
            transfer_classification_rows,
            category=category,
            amount_header="Transfer Amount (GHC)",
            category_header="Transfer Category",
        )
        for category in transfer_category_counts
    }

    write_nsano_wallet_ledger_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        nsano_disb_headers,
        nsano_disb_records,
        transfer_headers,
        transfer_rows,
        transfer_classification_rows,
    )

    return {
        **metrics,
        "workflow_label": "Nsano Disb Wallet vs Ledger",
        "mambu_total": len(mambu_records),
        "nsano_disb_total": len(nsano_disb_records),
        "transfer_total": len(transfer_rows),
        "transfer_category_counts": dict(transfer_category_counts),
        "transfer_category_amounts": transfer_category_amounts,
    }


def build_nsano_collection_ledger_reconciliation(
    output_path: Path,
    ledger_balance: Decimal,
    manual_unidentified: Decimal,
    manual_settlement: Decimal,
    manual_reversal: Decimal,
    mambu_collection_records: list[Record],
    mambu_collection_headers: list[str],
    mambu_disb_records: list[Record],
    mambu_disb_headers: list[str],
    transfer_rows: list[OrderedDict[str, Any]],
    transfer_headers: list[str],
    nsano_charge_records: list[Record] | None = None,
    nsano_charge_headers: list[str] | None = None,
) -> dict[str, Any]:
    nsano_charge_records = nsano_charge_records or []
    nsano_charge_headers = nsano_charge_headers or []
    transfer_classification_rows = build_nsano_transfer_classification_rows(transfer_rows)
    month_title, month_short = detect_nsano_wallet_month(
        [*mambu_collection_records, *mambu_disb_records],
        transfer_rows,
    )
    mambu_collections = sum_amounts(mambu_collection_records)
    charges = sum_amounts(nsano_charge_records)
    summary_rows, metrics = build_nsano_wallet_ledger_summary(
        ledger_balance,
        manual_unidentified,
        manual_settlement,
        manual_reversal,
        month_title,
        month_short,
        mambu_disb_records,
        transfer_classification_rows,
        credit_collection_amount=mambu_collections,
        credit_collection_count=len(mambu_collection_records),
        credit_collection_label="Total Nsano Mambu Collections",
        wallet_column_label="Nsano Collections",
        charge_amount=charges,
        charge_count=len(nsano_charge_records),
    )

    transfer_category_counts = Counter(
        str(row.get("Transfer Category", ""))
        for row in transfer_classification_rows
    )
    transfer_category_amounts = {
        category: category_amount(
            transfer_classification_rows,
            category=category,
            amount_header="Transfer Amount (GHC)",
            category_header="Transfer Category",
        )
        for category in transfer_category_counts
    }

    write_nsano_collection_ledger_workbook(
        output_path,
        summary_rows,
        mambu_collection_headers,
        mambu_collection_records,
        mambu_disb_headers,
        mambu_disb_records,
        nsano_charge_headers,
        nsano_charge_records,
        transfer_headers,
        transfer_rows,
        transfer_classification_rows,
    )

    return {
        **metrics,
        "workflow_label": "Nsano Collections vs Ledger",
        "mambu_collection_total": len(mambu_collection_records),
        "mambu_collection_amount": mambu_collections,
        "mambu_disb_total": len(mambu_disb_records),
        "nsano_charge_total": len(nsano_charge_records),
        "transfer_total": len(transfer_rows),
        "transfer_category_counts": dict(transfer_category_counts),
        "transfer_category_amounts": transfer_category_amounts,
    }


def build_mtn_manual_disb_reconciliation(
    output_path: Path,
    mambu_records: list[Record],
    mambu_headers: list[str],
    mtn_records: list[Record],
    mtn_headers: list[str],
    refund_records: list[Record],
    refund_headers: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    mambu_vs_mtn_rows = compare_mambu_to_mtn_manual_disb(mambu_records, mtn_records)
    mtn_vs_sources_rows = compare_mtn_manual_disb_to_sources(mtn_records, mambu_records, refund_records)
    mambu_vs_mtn_stats = summarize_disb_compare(
        mambu_vs_mtn_rows,
        matched_status=MTN_MANUAL_MATCHED_STATUS,
        amount_header="Amount",
        not_found_status=MTN_MANUAL_NOT_FOUND_STATUS,
    )
    mtn_vs_sources_stats = summarize_disb_compare_multi(
        mtn_vs_sources_rows,
        matched_statuses={MTN_MANUAL_MAMBU_STATUS, MTN_MANUAL_REFUND_STATUS},
        amount_header="Amount",
        not_found_status=MTN_MANUAL_NOT_FOUND_STATUS,
    )
    mambu_vs_mtn_stats["not_found_details"] = not_found_detail_rows(
        mambu_records,
        mambu_vs_mtn_rows,
        MTN_MANUAL_NOT_FOUND_STATUS,
        "Mambu",
        "Mambu → Wallet",
    )
    mtn_vs_sources_stats["not_found_details"] = not_found_detail_rows(
        mtn_records,
        mtn_vs_sources_rows,
        MTN_MANUAL_NOT_FOUND_STATUS,
        "MTN Manual",
        "Wallet → Mambu",
    )
    summary_rows = build_mtn_manual_disb_summary_rows(mambu_vs_mtn_stats, mtn_vs_sources_stats)
    write_mtn_manual_disb_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        mtn_headers,
        mtn_records,
        refund_headers,
        refund_records,
        mambu_vs_mtn_rows,
        mtn_vs_sources_rows,
    )
    return mambu_vs_mtn_stats, mtn_vs_sources_stats


def build_vodafone_manual_disb_reconciliation(
    output_path: Path,
    mambu_records: list[Record],
    mambu_headers: list[str],
    vodafone_records: list[Record],
    vodafone_headers: list[str],
    refund_records: list[Record],
    refund_headers: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    mambu_vs_vodafone_rows = compare_mambu_to_vodafone_manual_disb(mambu_records, vodafone_records)
    vodafone_vs_sources_rows = compare_vodafone_manual_disb_to_sources(vodafone_records, mambu_records, refund_records)
    mambu_vs_vodafone_stats = summarize_disb_compare(
        mambu_vs_vodafone_rows,
        matched_status=VODAFONE_MANUAL_MATCHED_STATUS,
        amount_header="Amount",
        not_found_status=VODAFONE_MANUAL_NOT_FOUND_STATUS,
    )
    vodafone_vs_sources_stats = summarize_disb_compare_multi(
        vodafone_vs_sources_rows,
        matched_statuses={VODAFONE_MANUAL_MAMBU_STATUS, VODAFONE_MANUAL_REFUND_STATUS},
        amount_header="Amount",
        not_found_status=VODAFONE_MANUAL_NOT_FOUND_STATUS,
    )
    mambu_vs_vodafone_stats["not_found_details"] = not_found_detail_rows(
        mambu_records,
        mambu_vs_vodafone_rows,
        VODAFONE_MANUAL_NOT_FOUND_STATUS,
        "Mambu",
        "Mambu → Wallet",
    )
    vodafone_vs_sources_stats["not_found_details"] = not_found_detail_rows(
        vodafone_records,
        vodafone_vs_sources_rows,
        VODAFONE_MANUAL_NOT_FOUND_STATUS,
        "Vodafone Manual",
        "Wallet → Mambu",
    )
    summary_rows = build_vodafone_manual_disb_summary_rows(mambu_vs_vodafone_stats, vodafone_vs_sources_stats)
    write_vodafone_manual_disb_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        vodafone_headers,
        vodafone_records,
        refund_headers,
        refund_records,
        mambu_vs_vodafone_rows,
        vodafone_vs_sources_rows,
    )
    return mambu_vs_vodafone_stats, vodafone_vs_sources_stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Nsano disbursement reconciliation workbook.")
    parser.add_argument("--profile", choices=["nsano", "itc", "itc-ledger", "mtn", "vodafone"], default="nsano", help="Disbursement reconciliation profile.")
    parser.add_argument("--source", default=None, help="Workbook containing Mambu and NSANO DISB sheets.")
    parser.add_argument("--mambu", default=None, help="Mambu disbursement xlsx/csv path.")
    parser.add_argument("--nsano", default=None, help="Nsano disbursement xlsx/csv path.")
    parser.add_argument("--itc", default=None, help="ITC disbursement xlsx/csv path.")
    parser.add_argument("--itc-statement", default=None, help="ITC Statement xlsx/csv path for the ITC Wallet vs Ledger profile.")
    parser.add_argument("--debit-transfers", default=None, help="ITC Debit Transfers xlsx/csv path for the ITC Wallet vs Ledger profile.")
    parser.add_argument("--credit-transfers", default=None, help="ITC Credit Transfers xlsx/csv path for the ITC Wallet vs Ledger profile.")
    parser.add_argument("--ledger-balance", default="0", help="Manual ledger balance for the ITC Wallet vs Ledger profile.")
    parser.add_argument("--unidentified", default="0", help="Manual unidentified amount for the ITC Wallet vs Ledger profile. Positive values are credit; negative values are debit.")
    parser.add_argument("--mtn-manual", default=None, help="MTN Manual disbursement xlsx/csv path.")
    parser.add_argument("--vodafone-manual", default=None, help="Vodafone Manual disbursement xlsx/csv path.")
    parser.add_argument("--refund", default=None, help="Refund xlsx/csv path.")
    parser.add_argument("--output", default=None, help="Output xlsx path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    default_outputs = {
        "itc": "ITC_Wallet_vs_Mambu.xlsx",
        "itc-ledger": "ITC_Wallet_vs_Ledger.xlsx",
        "mtn": "MTN_MANUAL_DISB_Recon.xlsx",
        "vodafone": "VF_MANUAL_DISB_Recon.xlsx",
        "nsano": "Nsano_DISB_Recon.xlsx",
    }
    default_output = default_outputs[args.profile]
    output_path = Path(args.output or default_output)

    if args.profile == "itc":
        if args.source:
            mambu_records, mambu_headers, itc_records, itc_headers = load_itc_disb_workbook(Path(args.source))
        else:
            if not args.mambu or not args.itc:
                raise SystemExit("Provide either --source or both --mambu and --itc for the ITC profile.")
            mambu_records, mambu_headers, itc_records, itc_headers = load_itc_disb_sources(
                Path(args.mambu),
                Path(args.itc),
            )

        mambu_vs_itc_stats, itc_vs_mambu_stats = build_itc_disb_reconciliation(
            output_path,
            mambu_records,
            mambu_headers,
            itc_records,
            itc_headers,
        )
        print(f"Wrote {output_path}")
        print(
            f"Mambu vs ITC Wallet: {mambu_vs_itc_stats['matched']:,}/{mambu_vs_itc_stats['total']:,} matched; "
            f"ITC Wallet vs Mambu: {itc_vs_mambu_stats['matched']:,}/{itc_vs_mambu_stats['total']:,} matched"
        )
        return

    if args.profile == "itc-ledger":
        try:
            ledger_balance = Decimal(str(args.ledger_balance or "0").replace(",", "").strip())
        except InvalidOperation as exc:
            raise SystemExit("--ledger-balance must be a valid amount.") from exc
        try:
            manual_unidentified = Decimal(str(args.unidentified or "0").replace(",", "").strip())
        except InvalidOperation as exc:
            raise SystemExit("--unidentified must be a valid amount.") from exc

        if args.source:
            (
                statement_rows,
                statement_headers,
                debit_rows,
                debit_headers,
                credit_rows,
                credit_headers,
            ) = load_itc_wallet_ledger_workbook(Path(args.source))
        else:
            if not args.itc_statement or not args.debit_transfers or not args.credit_transfers:
                raise SystemExit(
                    "Provide either --source or all of --itc-statement, --debit-transfers, and --credit-transfers "
                    "for the ITC Wallet vs Ledger profile."
                )
            (
                statement_rows,
                statement_headers,
                debit_rows,
                debit_headers,
                credit_rows,
                credit_headers,
            ) = load_itc_wallet_ledger_sources(
                Path(args.itc_statement),
                Path(args.debit_transfers),
                Path(args.credit_transfers),
            )

        metrics = build_itc_wallet_ledger_reconciliation(
            output_path,
            ledger_balance,
            manual_unidentified,
            statement_rows,
            statement_headers,
            debit_rows,
            debit_headers,
            credit_rows,
            credit_headers,
        )
        print(f"Wrote {output_path}")
        print(
            f"Debit Transfers: {metrics['debit_total']:,}; "
            f"Available Funds Before Debit: {metrics['available_funds']:,.2f}; "
            f"TOTAL DEBIT: {metrics['total_debit']:,.2f}; "
            f"Wallet Statement Balance: {metrics['wallet_statement_balance']:,.2f}"
        )
        return

    if args.profile == "mtn":
        if args.source:
            (
                mambu_records,
                mambu_headers,
                mtn_records,
                mtn_headers,
                refund_records,
                refund_headers,
            ) = load_mtn_manual_disb_workbook(Path(args.source))
        else:
            if not args.mambu or not args.mtn_manual:
                raise SystemExit("Provide either --source or both --mambu and --mtn-manual for the MTN profile.")
            (
                mambu_records,
                mambu_headers,
                mtn_records,
                mtn_headers,
                refund_records,
                refund_headers,
            ) = load_mtn_manual_disb_sources(
                Path(args.mambu),
                Path(args.mtn_manual),
                Path(args.refund) if args.refund else None,
            )

        mambu_vs_mtn_stats, mtn_vs_sources_stats = build_mtn_manual_disb_reconciliation(
            output_path,
            mambu_records,
            mambu_headers,
            mtn_records,
            mtn_headers,
            refund_records,
            refund_headers,
        )
        print(f"Wrote {output_path}")
        print(
            f"Mambu vs MTN MANUAL: {mambu_vs_mtn_stats['matched']:,}/{mambu_vs_mtn_stats['total']:,} matched; "
            f"MTN MANUAL as Source: {mtn_vs_sources_stats['matched']:,}/{mtn_vs_sources_stats['total']:,} matched"
        )
        return

    if args.profile == "vodafone":
        if args.source:
            (
                mambu_records,
                mambu_headers,
                vodafone_records,
                vodafone_headers,
                refund_records,
                refund_headers,
            ) = load_vodafone_manual_disb_workbook(Path(args.source))
        else:
            if not args.mambu or not args.vodafone_manual:
                raise SystemExit("Provide either --source or both --mambu and --vodafone-manual for the Vodafone profile.")
            (
                mambu_records,
                mambu_headers,
                vodafone_records,
                vodafone_headers,
                refund_records,
                refund_headers,
            ) = load_vodafone_manual_disb_sources(
                Path(args.mambu),
                Path(args.vodafone_manual),
                Path(args.refund) if args.refund else None,
            )

        mambu_vs_vodafone_stats, vodafone_vs_sources_stats = build_vodafone_manual_disb_reconciliation(
            output_path,
            mambu_records,
            mambu_headers,
            vodafone_records,
            vodafone_headers,
            refund_records,
            refund_headers,
        )
        print(f"Wrote {output_path}")
        print(
            f"Mambu vs VODAFONE MANUAL: {mambu_vs_vodafone_stats['matched']:,}/{mambu_vs_vodafone_stats['total']:,} matched; "
            f"VODAFONE MANUAL as Source: {vodafone_vs_sources_stats['matched']:,}/{vodafone_vs_sources_stats['total']:,} matched"
        )
        return

    if args.source:
        mambu_records, mambu_headers, nsano_records, nsano_headers = load_nsano_disb_workbook(Path(args.source))
    else:
        if not args.mambu or not args.nsano:
            raise SystemExit("Provide either --source or both --mambu and --nsano.")
        mambu_records, mambu_headers, nsano_records, nsano_headers = load_nsano_disb_sources(
            Path(args.mambu),
            Path(args.nsano),
        )

    mambu_vs_nsano_stats, nsano_vs_mambu_stats = build_nsano_disb_reconciliation(
        output_path,
        mambu_records,
        mambu_headers,
        nsano_records,
        nsano_headers,
    )
    print(f"Wrote {output_path}")
    print(
        f"Mambu vs Nsano: {mambu_vs_nsano_stats['matched']:,}/{mambu_vs_nsano_stats['total']:,} matched; "
        f"Nsano vs Mambu: {nsano_vs_mambu_stats['matched']:,}/{nsano_vs_mambu_stats['total']:,} matched"
    )


if __name__ == "__main__":
    main()
