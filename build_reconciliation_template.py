#!/usr/bin/env python3
"""
Build a seven-sheet Excel reconciliation template for Mambu vs Nsano.

Default use in this folder:
    python3 build_reconciliation_template.py

Reusable examples:
    python3 build_reconciliation_template.py \
        --mambu "Mambu-*.xlsx" \
        --nsano "Nsano.csv" \
        --output "outputs/mambu_nsano_reconciliation.xlsx"

For other file pairs, update the PROFILE block near the top or pass the
input/output paths on the command line. The matching engine is driven by
normalized reconciliation key + direction + amount rounded to 2 decimals.

The script intentionally uses only Python's standard library, so it can run
on machines without pandas/openpyxl/xlsxwriter installed.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import os
import re
import sys
import tomllib
import zipfile
from collections import Counter, OrderedDict, defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Iterable
from xml.etree.ElementTree import fromstring, iterparse
from xml.sax.saxutils import escape


MATCHED = "✓ MATCHED"
NOT_FOUND = "✗ NOT FOUND"
MAMBU_STATUS = "MAMBU"
WRITE_OFF_STATUS = "WRITE-OFF"
UNIDENTIFIED_STATUS = "UNIDENTIFIED"
NSANO_NOT_FOUND_STATUS = "NOT FOUND"
ITC_STATUS = "ITC"
VODA_COLL_STATUS = "Voda_Coll"
ITC_MAMBU_STATUS = "mambu"
ITC_UNIDENTIFIED_STATUS = "unidentified"
ITC_WRITE_OFF_STATUS = "write-off"
ITC_UPSALE_STATUS = "upsale"
VODAFONE_WRITE_OFF_STATUS = "write off"
ITC_NOT_FOUND_STATUS = "not_found"
VODAFONE_CHARGE_DETAIL = "Pay Bill Charge"
WRITE_OFF_NSANO_STATUS = "Nsano"
WRITE_OFF_ITC_STATUS = "ITC"
WRITE_OFF_ZENITH_STATUS = "Zenith"
WRITE_OFF_VODAFONE_STATUS = "Vodafone Collections"
WRITE_OFF_NOT_FOUND_STATUS = "Not Found"
WRITE_OFF_AMBIGUOUS_STATUS = "Ambiguous"
REPORT_LOGIC_VERSION = "2026.07.28-r1"
WRITE_OFF_IDENTIFIER_WALLETS = (
    WRITE_OFF_NSANO_STATUS,
    WRITE_OFF_ITC_STATUS,
    WRITE_OFF_VODAFONE_STATUS,
)
FAST_XLSX_COMPRESSLEVEL = 1
WIDTH_SAMPLE_ROWS = 1000
XLSX_ROW_BATCH_SIZE = 2000

MAMBU_BASE_HEADERS = [
    "Source_File",
    "Source_Row",
    "Source_Dataset",
    "Direction",
    "Reconciliation_Key",
    "Reconciliation_Amount",
    "Reconciliation_DateTime",
    "Reconciliation_Phone",
    "Duplicate_Source_Rows",
    "Duplicate_Source_Files",
]

NSANO_BASE_HEADERS = [
    "Source_File",
    "Source_Row",
    "Direction",
    "Reconciliation_Key",
    "Reconciliation_Amount",
    "Reconciliation_DateTime",
    "Reconciliation_Phone",
]

COMPARE_HEADERS = [
    "Source_File",
    "Source_Row",
    "Source_Dataset",
    "Direction",
    "Source_DateTime",
    "Source_Key",
    "Source_Amount",
    "Source_Channel_Type",
    "Source_Account",
    "Source_Phone",
    "Match_Status",
    "Match_Reason",
    "Counterparty_File",
    "Counterparty_Row",
    "Counterparty_Dataset",
    "Counterparty_Key",
    "Counterparty_Amount",
    "Amount_Difference",
    "Counterparty_Result",
    "Counterparty_Transaction_ID",
]

COMPARE_RESULT_HEADERS = [
    "Match_Status",
    "Match_Reason",
    "Counterparty_Dataset",
    "Counterparty_Key",
    "Counterparty_Amount",
    "Amount_Difference",
    "Counterparty_File",
    "Counterparty_Row",
]

WRITE_OFF_RECON_HEADERS = [
    "Write Off Date",
    "Write Off Identifier",
    "Write Off Amount (GHC)",
    "Write Off Channel",
    "Routed Wallet",
    "Write Off Account",
    "Match Status",
    "Investigation Status",
    "Match Reason",
    "Candidate Wallets",
    "Zenith Date Check",
    "Zenith Amount Check",
    "Zenith Mismatch Detail",
    "Zenith Candidate Date",
    "Zenith Candidate Amount (GHC)",
    "Matched Identifier",
    "Wallet Amount (GHC)",
    "Amount Variance (GHC)",
    "Matched Source Row",
    "Matched Source File",
]

WRITE_OFF_HEADERS = [
    "Date",
    "Channel",
    "Amount",
    "Fido Client Name",
    "Account Number",
    "Transaction ID",
    "Loan ID",
    "Staus",
]

UNIDENTIFIED_HEADERS = [
    "Value Date (Entry Date)",
    "Identifier",
    "Amount",
    "Type",
    "User",
    "Account ID",
    "Notes",
    "Product (Deposit)",
]

MAMBU_SHEET_COLUMNS: list[str] = [
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

NSANO_SHEET_COLUMNS: list[str] = [
    "DateTime",
    "Transaction_ID",
    "Type",
    "Author",
    "Amount_GHC",
    "balanceBefore",
    "balanceAfter",
    "SendingHse",
    "SendingHse_ID",
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
    "Defferred",
    "Service_Label",
    "Service_Transaction_ID",
    "Service_Details",
]

ITC_SHEET_COLUMNS: list[str] = [
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
    "net_amount",
    "fees",
    "elevy_charge",
    "country_code",
    "created",
    "merchant_product_name",
]

VODAFONE_COLLECTIONS_SHEET_COLUMNS: list[str] = [
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

WRITE_OFF_SHEET_COLUMNS: list[str] = [
    "Date",
    "Channel",
    "Amount",
    "Fido Client Name",
    "Account Number",
    "Transaction ID",
    "Loan ID",
    "Staus",
]

UNIDENTIFIED_SHEET_COLUMNS: list[str] = [
    "Value Date (Entry Date)",
    "Identifier",
    "Amount",
    "Type",
    "User",
    "Account ID",
    "Notes",
    "Product (Deposit)",
]


@dataclass
class Cell:
    value: Any
    style: int | None = None


@dataclass
class Record:
    sheet_values: OrderedDict[str, Any]
    raw: dict[str, Any]
    source_file: str
    source_row: int
    source_dataset: str
    direction: str
    key: str
    amount_cents: int | None
    datetime_value: str
    phone: str
    channel_type: str
    account: str
    result: str = ""
    transaction_id: str = ""
    duplicate_rows: int = 1
    duplicate_files: set[str] | None = None

    @property
    def amount_decimal(self) -> Decimal | str:
        if self.amount_cents is None:
            return ""
        return Decimal(self.amount_cents) / Decimal(100)


@lru_cache(maxsize=4096)
def amount_to_cents(value: Any) -> int | None:
    text = str(value or "").strip().replace(",", "")
    if text == "":
        return None
    try:
        decimal_value = Decimal(text)
    except InvalidOperation:
        return None
    return int((decimal_value * Decimal(100)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def cents_to_decimal(cents: int | None) -> Decimal | str:
    if cents is None:
        return ""
    return Decimal(cents) / Decimal(100)


@lru_cache(maxsize=8192)
def normalize_key(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").strip()).upper()


@lru_cache(maxsize=8192)
def clean_reconciliation_identifier(value: Any) -> str:
    """Normalize transaction references used for reconciliation.

    Examples:
    - 0000012606078397 -> 12606078397
    - ‘0000013086073858 -> 13086073858
    - 0000012590103982/9080748521/233507192026/ESTHER AGYEMANG -> 12590103982
    """
    first_part = str(value or "").strip().split("/", 1)[0].strip()
    compact = re.sub(r"\s+", "", first_part).lstrip("’’")
    if compact.isdigit():
        return compact.lstrip("0") or "0"
    return compact


@lru_cache(maxsize=8192)
def normalize_reconciliation_identifier(value: Any) -> str:
    return normalize_key(clean_reconciliation_identifier(value))


@lru_cache(maxsize=4096)
def normalize_phone(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or ""))


@lru_cache(maxsize=4096)
def excel_serial_to_iso(value: Any, *, date_only: bool = False) -> str:
    text = str(value or "").strip()
    if text == "":
        return ""
    try:
        number = float(text)
    except ValueError:
        return text
    if not 20000 <= number <= 60000:
        return text
    converted = dt.datetime(1899, 12, 30) + dt.timedelta(days=number)
    if date_only:
        return converted.strftime("%Y-%m-%d")
    return converted.strftime("%Y-%m-%d %H:%M:%S")


def normalize_date_amount_match_date(value: Any) -> str:
    """Normalize dates used by date+amount matching branches."""
    text = str(value or "").strip()
    if not text:
        return ""

    excel_date = excel_serial_to_iso(text, date_only=True)
    if excel_date != text:
        return excel_date

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


def date_amount_match_key(date_value: Any, amount_value: Any) -> str:
    date_key = normalize_date_amount_match_date(date_value)
    amount_cents = amount_to_cents(amount_value)
    if not date_key or amount_cents is None:
        return ""
    return f"{date_key}|{amount_cents}"


@lru_cache(maxsize=4096)
def parse_nsano_datetime(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text_without_tz = text.split(" GMT", 1)[0]
    try:
        return dt.datetime.strptime(text_without_tz, "%a %b %d %Y %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return text


@lru_cache(maxsize=128)
def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def col_index(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char.upper()) - 64
    return value


def xml_text(value: Any) -> str:
    s = str(value)
    if "&" not in s and "<" not in s and ">" not in s and '"' not in s:
        return s
    return escape(s, {'"': "&quot;"})


def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:31]
    return cleaned or "Sheet"


def set_csv_field_limit() -> None:
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            return
        except OverflowError:
            limit = int(limit / 10)


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        handle = zf.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    values: list[str] = []
    with handle:
        for event, elem in iterparse(handle, events=("end",)):
            if elem.tag.endswith("}si"):
                values.append("".join(t.text or "" for t in elem.iter() if t.tag.endswith("}t")))
                elem.clear()
    return values


def worksheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = fromstring(zf.read("xl/workbook.xml"))
    rels = fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheets = next(elem for elem in workbook.iter() if elem.tag.endswith("}sheets"))
    paths: dict[str, str] = {}
    for sheet in sheets:
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rid_to_target[rid].lstrip("/")
        paths[sheet.attrib["name"].casefold()] = target if target.startswith("xl/") else f"xl/{target}"
    return paths


def first_worksheet_path(zf: zipfile.ZipFile) -> str:
    return next(iter(worksheet_paths(zf).values()))


def worksheet_path_by_name(zf: zipfile.ZipFile, sheet_name: str) -> str:
    paths = worksheet_paths(zf)
    key = sheet_name.casefold()
    if key in paths:
        return paths[key]
    # Try normalizing hyphens/spaces so "Write Off" and "Write-Off" both match
    normalized = re.sub(r"[-\s]+", "", key)
    for path_key, path_val in paths.items():
        if re.sub(r"[-\s]+", "", path_key) == normalized:
            return path_val
    available = ", ".join(paths)
    raise KeyError(f"Sheet '{sheet_name}' not found. Available sheets: {available}")


def xlsx_sheet_if_present(path: Path, sheet_name: str) -> str | None:
    """Return ``sheet_name`` if the workbook at ``path`` contains it, else None.

    Lets callers target a specific sheet (e.g. "Inflow"/"Outflow" from the ITC
    Collection Filtering output) while falling back to default sheet selection
    for CSVs and workbooks that don't have that sheet.
    """
    if path.suffix.lower() == ".csv":
        return None
    try:
        with zipfile.ZipFile(path) as zf:
            sheets = worksheet_paths(zf)
    except Exception:
        return None
    return sheet_name if sheet_name.casefold() in sheets else None


def cell_value(cell: Any, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter() if node.tag.endswith("}t"))

    value_node = next((child for child in cell if child.tag.endswith("}v")), None)
    if value_node is None:
        return ""
    value = value_node.text or ""
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return value
    return value


CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


def iter_xlsx_dicts(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
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
                    headers = row_values
                else:
                    if not any(v.strip() for v in row_values):
                        elem.clear()
                        continue
                    padded = row_values + [""] * max(0, len(headers) - len(row_values))
                    yield row_number, dict(zip(headers, padded))
                elem.clear()


def iter_csv_dicts(path: Path) -> Iterable[tuple[int, dict[str, str]]]:
    set_csv_field_limit()
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row_number, row in enumerate(reader, start=2):
            yield row_number, {key: (value if value is not None else "") for key, value in row.items()}


def add_unique_headers(target: list[str], new_headers: Iterable[str]) -> None:
    for header in new_headers:
        if header and header not in target:
            target.append(header)


def normalize_mambu_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    is_disbursement = "Principal Amount" in row
    dataset = "Disbursement" if is_disbursement else "Collection"
    direction = "DISBURSEMENT" if is_disbursement else "COLLECTION"
    amount_cents = amount_to_cents(row.get("Amount"))
    original_key = row.get("Identifier", "") or row.get("Transaction ID", "")
    key = normalize_reconciliation_identifier(original_key)
    phone = normalize_phone(row.get("Mobile Phone (Client)", ""))
    date_value = excel_serial_to_iso(row.get("Value Date (Entry Date)", "") or row.get("Date", ""))
    channel = row.get("Channel", "") or row.get("Type", "")
    account = (
        row.get("Account ID")
        or row.get("Account Number")
        or row.get("Loan ID")
        or row.get("Loan Number")
        or row.get("Account Holder ID", "")
    )

    values: OrderedDict[str, Any] = OrderedDict()
    values["Source_File"] = path.name
    values["Source_Row"] = source_row
    values["Source_Dataset"] = dataset
    values["Direction"] = direction
    values["Reconciliation_Key"] = key
    values["Reconciliation_Amount"] = cents_to_decimal(amount_cents)
    values["Reconciliation_DateTime"] = date_value
    values["Reconciliation_Phone"] = phone
    values["Duplicate_Source_Rows"] = 1
    values["Duplicate_Source_Files"] = path.name

    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset=dataset,
        direction=direction,
        key=key,
        amount_cents=amount_cents,
        datetime_value=date_value,
        phone=phone,
        channel_type=channel,
        account=account,
        duplicate_files={path.name},
    )


def normalize_nsano_record(path: Path, source_row: int, row: dict[str, str]) -> Record:
    transaction_type = str(row.get("Type", "")).strip().upper()
    if transaction_type == "A2W":
        direction = "DISBURSEMENT"
        raw_key = row.get("SendingHse_ID", "")
        key = normalize_key(raw_key[4:] if normalize_key(raw_key).startswith("FIDO") else raw_key)
        phone = normalize_phone(row.get("ReceivingHse_Account", ""))
        account = row.get("ReceivingHse_ID", "")
    elif transaction_type == "W2A":
        direction = "COLLECTION"
        key = normalize_reconciliation_identifier(row.get("External_Debit_Reference", ""))
        phone = normalize_phone(row.get("SendingHse_Account", ""))
        account = row.get("SendingHse_ID", "")
    else:
        direction = transaction_type or "UNKNOWN"
        key = normalize_key(row.get("Transaction_ID", ""))
        phone = normalize_phone(row.get("SendingHse_Account", "") or row.get("ReceivingHse_Account", ""))
        account = row.get("SendingHse_ID", "") or row.get("ReceivingHse_ID", "")

    amount_cents = amount_to_cents(row.get("Amount_GHC"))
    date_value = parse_nsano_datetime(row.get("DateTime", ""))
    values: OrderedDict[str, Any] = OrderedDict()
    values["Source_File"] = path.name
    values["Source_Row"] = source_row
    values["Direction"] = direction
    values["Reconciliation_Key"] = key
    values["Reconciliation_Amount"] = cents_to_decimal(amount_cents)
    values["Reconciliation_DateTime"] = date_value
    values["Reconciliation_Phone"] = phone

    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset="Nsano",
        direction=direction,
        key=key,
        amount_cents=amount_cents,
        datetime_value=date_value,
        phone=phone,
        channel_type=transaction_type,
        account=account,
        result=row.get("Result", ""),
        transaction_id=row.get("Transaction_ID", ""),
    )


@lru_cache(maxsize=1024)
def _clean_sheet_header_kind(header: str) -> str:
    normalized_header = header.strip().casefold()
    compact_header = re.sub(r"[^a-z0-9]+", "", normalized_header)
    if header in {
        "Identifier", "External_Debit_Reference", "Transaction ID", "Transactions Id", "Transactions ID",
        "Identifier (Key)", "Ext Debit Ref (Key)", "Key (Identifier)",
    } or normalized_header in {"channel_transaction_id", "receipt no.", "receipt no"} or compact_header in {
        "channeltransactionid",
        "itctransidkey",
        "receiptno",
        "receiptnumber",
    }:
        return "identifier"
    if header == "DateTime":
        return "nsano_datetime"
    if normalized_header in {"transaction_date", "created", "completion time", "initiation time"}:
        return "datetime"
    if "Date" in header or header == "Date":
        return "date_only" if ("Birth Date" in header or header == "Date") else "datetime"
    if header in {"Amount", "Amount_GHC", "Amount (GHC)", "Principal Amount", "Charge", "balanceBefore", "balanceAfter"} or normalized_header in {
        "amount",
        "net_amount",
        "fees",
        "elevy_charge",
        "paid in",
        "withdrawn",
        "balance",
    }:
        return "amount"
    return "text"


def clean_sheet_value(header: str, value: Any) -> Any:
    kind = _clean_sheet_header_kind(header)
    if kind == "identifier":
        return clean_reconciliation_identifier(value)
    if kind == "nsano_datetime":
        return parse_nsano_datetime(value)
    if kind == "datetime":
        return excel_serial_to_iso(value)
    if kind == "date_only":
        return excel_serial_to_iso(value, date_only=True)
    if kind == "amount":
        cents = amount_to_cents(value)
        return cents_to_decimal(cents) if cents is not None else value
    return value


def normalize_unidentified_key(value: Any) -> str:
    return normalize_reconciliation_identifier(value)


def normalize_coll_recon_record(
    path: Path,
    sheet_name: str,
    source_row: int,
    row: dict[str, str],
    *,
    source_dataset: str,
    key_getter: Callable[[dict[str, str]], str],
    amount_header: str,
    base_headers: list[str],
    date_getter: Callable[[dict[str, str]], str] | None = None,
    phone_getter: Callable[[dict[str, str]], str] | None = None,
    account_getter: Callable[[dict[str, str]], str] | None = None,
    channel_getter: Callable[[dict[str, str]], str] | None = None,
    result_getter: Callable[[dict[str, str]], str] | None = None,
    transaction_id_getter: Callable[[dict[str, str]], str] | None = None,
    include_helpers: bool = True,
) -> Record:
    amount_cents = amount_to_cents(row.get(amount_header, ""))
    raw_key = key_getter(row)
    key = normalize_reconciliation_identifier(raw_key)
    date_value = date_getter(row) if date_getter else ""
    phone = normalize_phone(phone_getter(row) if phone_getter else "")
    account = account_getter(row) if account_getter else ""
    channel = channel_getter(row) if channel_getter else sheet_name
    result = result_getter(row) if result_getter else ""
    transaction_id = transaction_id_getter(row) if transaction_id_getter else ""

    values: OrderedDict[str, Any] = OrderedDict()
    for header in base_headers:
        values[header] = ""

    if include_helpers:
        values["Source_File"] = path.name
        values["Source_Row"] = source_row
        if "Source_Dataset" in values:
            values["Source_Dataset"] = source_dataset
        values["Direction"] = "COLLECTION"
        values["Reconciliation_Key"] = key
        values["Reconciliation_Amount"] = cents_to_decimal(amount_cents)
        values["Reconciliation_DateTime"] = date_value
        values["Reconciliation_Phone"] = phone

    for header, value in row.items():
        values[header] = clean_sheet_value(header, value)

    return Record(
        sheet_values=values,
        raw=row,
        source_file=path.name,
        source_row=source_row,
        source_dataset=source_dataset,
        direction="COLLECTION",
        key=key,
        amount_cents=amount_cents,
        datetime_value=date_value,
        phone=phone,
        channel_type=channel,
        account=account,
        result=result,
        transaction_id=transaction_id,
    )


def load_mambu(paths: list[Path], keep_duplicates: bool) -> tuple[list[Record], list[str], int]:
    headers = list(MAMBU_BASE_HEADERS)
    records: list[Record] = []
    by_reconciliation_key: OrderedDict[tuple[str, str, int | None], Record] = OrderedDict()
    raw_count = 0

    for path in paths:
        first_row = True
        for source_row, row in iter_source_dicts(path):
            raw_count += 1
            if first_row:
                add_unique_headers(headers, row.keys())
                first_row = False
            record = normalize_mambu_record(path, source_row, row)
            if keep_duplicates:
                records.append(record)
                continue

            dedupe_key = (record.direction, record.key, record.amount_cents)
            existing = by_reconciliation_key.get(dedupe_key)
            if existing is None:
                by_reconciliation_key[dedupe_key] = record
            else:
                existing.duplicate_rows += 1
                if existing.duplicate_files is None:
                    existing.duplicate_files = {existing.source_file}
                existing.duplicate_files.add(path.name)
                existing.sheet_values["Duplicate_Source_Rows"] = existing.duplicate_rows
                existing.sheet_values["Duplicate_Source_Files"] = ", ".join(sorted(existing.duplicate_files))

    if not keep_duplicates:
        records = list(by_reconciliation_key.values())

    for record in records:
        for header in headers:
            record.sheet_values.setdefault(header, "")
    return records, headers, raw_count


def load_nsano(path: Path) -> tuple[list[Record], list[str], int]:
    headers = list(NSANO_BASE_HEADERS)
    records: list[Record] = []
    failed_rows_ignored = 0
    headers_captured = False
    for source_row, row in iter_source_dicts(path):
        if not headers_captured:
            add_unique_headers(headers, row.keys())
            headers_captured = True
        if str(row.get("Result", "")).strip().casefold() == "failed":
            failed_rows_ignored += 1
            continue
        record = normalize_nsano_record(path, source_row, row)
        records.append(record)
    return records, headers, failed_rows_ignored


def load_coll_recon_workbook(path: Path) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[Record],
    int,
    int,
]:
    mambu_headers: list[str] = []
    nsano_headers: list[str] = []
    mambu_records: list[Record] = []
    nsano_records: list[Record] = []
    write_off_records: list[Record] = []
    unidentified_records: list[Record] = []
    nsano_failed_rows_ignored = 0

    for source_row, row in iter_xlsx_dicts(path, "Mambu"):
        add_unique_headers(mambu_headers, row.keys())
        record = normalize_coll_recon_record(
            path,
            "Mambu",
            source_row,
            row,
            source_dataset="Mambu",
            key_getter=lambda r: r.get("Identifier", ""),
            amount_header="Amount",
            base_headers=[],
            date_getter=lambda r: excel_serial_to_iso(r.get("Value Date (Entry Date)", "")),
            phone_getter=lambda r: r.get("Mobile Phone (Client)", ""),
            account_getter=lambda r: r.get("Account ID", ""),
            channel_getter=lambda r: r.get("Channel", ""),
            include_helpers=False,
        )
        for header in mambu_headers:
            record.sheet_values.setdefault(header, "")
        mambu_records.append(record)

    for source_row, row in iter_xlsx_dicts(path, "Nsano"):
        add_unique_headers(nsano_headers, row.keys())
        if str(row.get("Result", "")).strip().casefold() == "failed":
            nsano_failed_rows_ignored += 1
            continue
        record = normalize_coll_recon_record(
            path,
            "Nsano",
            source_row,
            row,
            source_dataset="Nsano",
            key_getter=lambda r: r.get("External_Debit_Reference", ""),
            amount_header="Amount_GHC",
            base_headers=[],
            date_getter=lambda r: parse_nsano_datetime(r.get("DateTime", "")),
            phone_getter=lambda r: r.get("SendingHse_Account", ""),
            account_getter=lambda r: r.get("SendingHse_ID", ""),
            channel_getter=lambda r: r.get("Type", ""),
            result_getter=lambda r: r.get("Result", ""),
            transaction_id_getter=lambda r: r.get("Transaction_ID", ""),
            include_helpers=False,
        )
        for header in nsano_headers:
            record.sheet_values.setdefault(header, "")
        nsano_records.append(record)

    for source_row, row in iter_xlsx_dicts(path, "Write Off"):
        record = normalize_coll_recon_record(
            path,
            "Write Off",
            source_row,
            row,
            source_dataset="Write-Off",
            key_getter=lambda r: r.get("Transaction ID", "") or r.get("Transactions Id", "") or r.get("Transactions ID", ""),
            amount_header="Amount",
            base_headers=[],
            date_getter=lambda r: excel_serial_to_iso(r.get("Date", ""), date_only=True),
            phone_getter=lambda r: r.get("Account Number", ""),
            account_getter=lambda r: r.get("Loan ID", ""),
            channel_getter=lambda r: r.get("Channel", ""),
            transaction_id_getter=lambda r: r.get("Transaction ID", "") or r.get("Transactions Id", "") or r.get("Transactions ID", ""),
            include_helpers=False,
        )
        write_off_records.append(record)

    for source_row, row in iter_xlsx_dicts(path, "Unidentified"):
        record = normalize_coll_recon_record(
            path,
            "Unidentified",
            source_row,
            row,
            source_dataset="Unidentified",
            key_getter=lambda r: r.get("Identifier", ""),
            amount_header="Amount",
            base_headers=[],
            date_getter=lambda r: excel_serial_to_iso(r.get("Value Date (Entry Date)", ""), date_only=True),
            phone_getter=lambda r: r.get("Identifier", ""),
            account_getter=lambda r: r.get("Account ID", ""),
            channel_getter=lambda r: r.get("Type", ""),
            include_helpers=False,
        )
        unidentified_records.append(record)

    return (
        mambu_records,
        mambu_headers,
        nsano_records,
        nsano_headers,
        write_off_records,
        unidentified_records,
        len(mambu_records),
        nsano_failed_rows_ignored,
    )


def filter_to_columns(row: dict[str, str], expected: list[str]) -> dict[str, str]:
    return {col: row.get(col, "") for col in expected}


def _pick(row: dict[str, str], *keys: str) -> str:
    """Return the value of the first key found in row."""
    for key in keys:
        val = row.get(key, "")
        if val:
            return val
    return ""


def _pick_key(row: dict[str, str], *keys: str) -> str:
    """Return the first key that exists in row."""
    for key in keys:
        if key in row:
            return key
    return keys[-1]


def load_separate_sources(
    mambu_path: Path | None,
    nsano_path: Path | None,
    write_off_path: Path | None,
    unidentified_path: Path | None,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str], list[Record], list[Record], int, int]:
    mambu_headers: list[str] = []
    nsano_headers: list[str] = []
    mambu_records: list[Record] = []
    nsano_records: list[Record] = []
    write_off_records: list[Record] = []
    unidentified_records: list[Record] = []
    nsano_failed_rows_ignored = 0

    if mambu_path and mambu_path.exists():
        mambu_headers_captured = False
        for source_row, row in iter_source_dicts(mambu_path, mambu_sheet_name):
            if not mambu_headers_captured:
                add_unique_headers(mambu_headers, row.keys())
                mambu_headers_captured = True
            amount_col = _pick_key(row, "Amount (GHC)", "Amount")
            record = normalize_coll_recon_record(
                mambu_path, "Mambu", source_row, row,
                source_dataset="Mambu",
                key_getter=lambda r: _pick(r, "Identifier (Key)", "Identifier"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Date/Time", "Value Date (Entry Date)")),
                phone_getter=lambda r: _pick(r, "Mobile Phone (Client)", "Description", "Reference"),
                account_getter=lambda r: _pick(r, "Account ID"),
                channel_getter=lambda r: _pick(r, "Channel"),
                include_helpers=False,
            )
            mambu_records.append(record)

    if nsano_path and nsano_path.exists():
        nsano_sheet_name = xlsx_sheet_if_present(nsano_path, "Successful W2A")
        nsano_headers_captured = False
        for source_row, row in iter_source_dicts(nsano_path, nsano_sheet_name):
            if not nsano_headers_captured:
                add_unique_headers(nsano_headers, row.keys())
                nsano_headers_captured = True
            if str(_pick(row, "Result")).strip().casefold() == "failed":
                nsano_failed_rows_ignored += 1
                continue
            amount_col = _pick_key(row, "Amount (GHC)", "Amount_GHC")
            record = normalize_coll_recon_record(
                nsano_path, "Nsano", source_row, row,
                source_dataset="Nsano",
                key_getter=lambda r: _pick(r, "Ext Debit Ref (Key)", "External_Debit_Reference"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: parse_nsano_datetime(_pick(r, "Date/Time", "DateTime")),
                phone_getter=lambda r: _pick(r, "SendingHse_Account", "J"),
                account_getter=lambda r: _pick(r, "SendingHse_ID", "I"),
                channel_getter=lambda r: _pick(r, "Type", "C"),
                result_getter=lambda r: _pick(r, "Result"),
                transaction_id_getter=lambda r: _pick(r, "Trans ID", "Transaction_ID"),
                include_helpers=False,
            )
            nsano_records.append(record)

    if write_off_path and write_off_path.exists():
        for source_row, row in iter_source_dicts(write_off_path):
            amount_col = _pick_key(row, "Amount (GHC)", "C", "Amount")
            record = normalize_coll_recon_record(
                write_off_path, "Write Off", source_row, row,
                source_dataset="Write-Off",
                key_getter=lambda r: _pick(r, "Identifier (Key)", "Transaction ID", "Transactions Id", "Transactions ID"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Date/Time", "Date"), date_only=True),
                phone_getter=lambda r: _pick(r, "Account Number", "E"),
                account_getter=lambda r: _pick(r, "Loan ID", "G"),
                channel_getter=lambda r: _pick(r, "Channel", "B"),
                transaction_id_getter=lambda r: _pick(r, "Identifier (Key)", "Transaction ID", "Transactions Id", "Transactions ID"),
                include_helpers=False,
            )
            write_off_records.append(record)

    if unidentified_path and unidentified_path.exists():
        for source_row, row in iter_source_dicts(unidentified_path):
            amount_col = _pick_key(row, "Amount (GHC)", "Amount")
            record = normalize_coll_recon_record(
                unidentified_path, "Unidentified", source_row, row,
                source_dataset="Unidentified",
                key_getter=lambda r: _pick(r, "Key (Identifier)", "Identifier"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Date/Time", "Value Date (Entry Date)"), date_only=True),
                phone_getter=lambda r: _pick(r, "Key (Identifier)", "Identifier"),
                account_getter=lambda r: _pick(r, "Account ID", "D"),
                channel_getter=lambda r: _pick(r, "Type", "C"),
                include_helpers=False,
            )
            unidentified_records.append(record)

    return (
        mambu_records,
        mambu_headers,
        nsano_records,
        nsano_headers,
        write_off_records,
        unidentified_records,
        len(mambu_records),
        nsano_failed_rows_ignored,
    )


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


def iter_source_dicts(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, str]]]:
    if path.suffix.lower() == ".csv":
        sheet_name = None
    for row_number, items in _cached_source_rows(*_source_cache_key(path, sheet_name)):
        yield row_number, dict(items)


def clear_source_row_cache() -> None:
    _cached_source_rows.cache_clear()


def load_itc_separate_sources(
    mambu_path: Path | None,
    itc_path: Path | None,
    vodafone_collections_path: Path | None,
    write_off_path: Path | None,
    unidentified_path: Path | None,
    mambu_sheet_name: str | None = None,
) -> tuple[list[Record], list[str], list[Record], list[str], list[Record], list[str], list[Record], list[Record], int]:
    mambu_headers: list[str] = []
    itc_headers: list[str] = []
    vodafone_headers: list[str] = []
    mambu_records: list[Record] = []
    itc_records: list[Record] = []
    vodafone_records: list[Record] = []
    write_off_records: list[Record] = []
    unidentified_records: list[Record] = []

    if mambu_path and mambu_path.exists():
        mambu_headers_captured = False
        for source_row, row in iter_source_dicts(mambu_path, mambu_sheet_name):
            if not mambu_headers_captured:
                add_unique_headers(mambu_headers, row.keys())
                mambu_headers_captured = True
            amount_col = _pick_key(row, "Amount", "Amount (GHC)")
            record = normalize_coll_recon_record(
                mambu_path, "Mambu", source_row, row,
                source_dataset="Mambu",
                key_getter=lambda r: _pick(r, "Identifier", "Identifier (Key)"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Value Date (Entry Date)", "Date/Time")),
                phone_getter=lambda r: _pick(r, "Mobile Phone (Client)", "Description", "Reference"),
                account_getter=lambda r: _pick(r, "Account ID"),
                channel_getter=lambda r: _pick(r, "Channel"),
                include_helpers=False,
            )
            mambu_records.append(record)

    if itc_path and itc_path.exists():
        itc_sheet_name = xlsx_sheet_if_present(itc_path, "Inflow")
        itc_headers_captured = False
        for source_row, row in iter_source_dicts(itc_path, itc_sheet_name):
            if not itc_headers_captured:
                add_unique_headers(itc_headers, row.keys())
                itc_headers_captured = True
            amount_col = _pick_key(row, "amount", "Amount", "Amount (GHC)", "ITC Amount (GHC)")
            record = normalize_coll_recon_record(
                itc_path, "ITC", source_row, row,
                source_dataset="ITC",
                key_getter=lambda r: _pick(r, "channel_transaction_id", "Channel Transaction ID", "ITC Trans ID (Key)"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "transaction_date", "Transaction Date", "ITC Date")),
                phone_getter=lambda r: _pick(r, "payer_contact", "account_reference"),
                account_getter=lambda r: _pick(r, "account_reference"),
                channel_getter=lambda r: _pick(r, "channel"),
                transaction_id_getter=lambda r: _pick(r, "channel_transaction_id", "Channel Transaction ID", "ITC Trans ID (Key)"),
                include_helpers=False,
            )
            itc_records.append(record)

    if vodafone_collections_path and vodafone_collections_path.exists():
        vodafone_sheet_name = xlsx_sheet_if_present(
            vodafone_collections_path,
            "Cleaned Voda Coll",
        )
        vodafone_headers_captured = False
        for source_row, row in iter_source_dicts(
            vodafone_collections_path,
            vodafone_sheet_name,
        ):
            if not vodafone_headers_captured:
                add_unique_headers(vodafone_headers, row.keys())
                vodafone_headers_captured = True
            amount_col = _pick_key(row, "Paid In", "PaidIn", "Amount", "Amount (GHC)")
            record = normalize_coll_recon_record(
                vodafone_collections_path, "Vodafone Collections", source_row, row,
                source_dataset="Vodafone Collections",
                key_getter=lambda r: _pick(r, "Receipt No.", "Receipt No", "Receipt Number"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Completion Time", "Initiation Time", "Date")),
                phone_getter=lambda r: _pick(r, "Opposite Party", "Details"),
                account_getter=lambda r: _pick(r, "Opposite Party", "Details"),
                channel_getter=lambda r: _pick(r, "Reason Type", "Details"),
                transaction_id_getter=lambda r: _pick(r, "Receipt No.", "Receipt No", "Receipt Number"),
                include_helpers=False,
            )
            vodafone_records.append(record)

    if write_off_path and write_off_path.exists():
        for source_row, row in iter_source_dicts(write_off_path):
            amount_col = _pick_key(row, "Amount", "Amount (GHC)", "C")
            record = normalize_coll_recon_record(
                write_off_path, "Write Off", source_row, row,
                source_dataset="Write-Off",
                key_getter=lambda r: _pick(r, "Transaction ID", "Transactions Id", "Transactions ID", "Identifier (Key)"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Date", "Date/Time"), date_only=True),
                phone_getter=lambda r: _pick(r, "Account Number", "E"),
                account_getter=lambda r: _pick(r, "Loan ID", "G"),
                channel_getter=lambda r: _pick(r, "Channel", "B"),
                transaction_id_getter=lambda r: _pick(r, "Transaction ID", "Transactions Id", "Transactions ID", "Identifier (Key)"),
                include_helpers=False,
            )
            write_off_records.append(record)

    if unidentified_path and unidentified_path.exists():
        for source_row, row in iter_source_dicts(unidentified_path):
            amount_col = _pick_key(row, "Amount", "Amount (GHC)")
            record = normalize_coll_recon_record(
                unidentified_path, "Unidentified", source_row, row,
                source_dataset="Unidentified",
                key_getter=lambda r: _pick(r, "Identifier", "Key (Identifier)"),
                amount_header=amount_col,
                base_headers=[],
                date_getter=lambda r: excel_serial_to_iso(_pick(r, "Value Date (Entry Date)", "Date/Time"), date_only=True),
                phone_getter=lambda r: _pick(r, "Identifier", "Key (Identifier)"),
                account_getter=lambda r: _pick(r, "Account ID", "D"),
                channel_getter=lambda r: _pick(r, "Type", "C"),
                include_helpers=False,
            )
            unidentified_records.append(record)

    return (
        mambu_records,
        mambu_headers or list(MAMBU_SHEET_COLUMNS),
        itc_records,
        itc_headers or list(ITC_SHEET_COLUMNS),
        vodafone_records,
        vodafone_headers or list(VODAFONE_COLLECTIONS_SHEET_COLUMNS),
        write_off_records,
        unidentified_records,
        len(mambu_records),
    )


def is_write_off_record(record: Record) -> bool:
    text = (
        f"{record.channel_type} {record.source_dataset} {record.source_file} "
        f"{record.raw.get('Type', '')} {record.raw.get('Notes', '')} {record.raw.get('Product (Deposit)', '')}"
    ).casefold()
    compact = re.sub(r"[^a-z0-9]+", "", text)
    return "writeoff" in compact


def is_unidentified_record(record: Record) -> bool:
    text = f"{record.channel_type} {record.source_dataset} {record.source_file}".casefold()
    return "unidentified" in text or "suspense" in text


def split_mambu_sources(records: list[Record]) -> tuple[list[Record], list[Record], list[Record]]:
    normal_mambu: list[Record] = []
    write_off: list[Record] = []
    unidentified: list[Record] = []

    for record in records:
        if is_write_off_record(record):
            write_off.append(record)
        elif is_unidentified_record(record):
            unidentified.append(record)
        else:
            normal_mambu.append(record)

    return normal_mambu, write_off, unidentified


def build_index(records: list[Record]) -> dict[tuple[str, str], list[Record]]:
    index: dict[tuple[str, str], list[Record]] = defaultdict(list)
    for record in records:
        if record.key:
            index[(record.direction, record.key)].append(record)
    return index


def find_counterparty(source: Record, counterparty_index: dict[tuple[str, str], list[Record]], counterparty_name: str) -> tuple[str, str, Record | None]:
    if not source.key:
        return NOT_FOUND, "blank reconciliation key", None

    candidates = counterparty_index.get((source.direction, source.key), [])
    if not candidates:
        return NOT_FOUND, f"reference not found in {counterparty_name}", None

    same_amount = next((candidate for candidate in candidates if candidate.amount_cents == source.amount_cents), None)
    if same_amount is not None:
        return MATCHED, "reference and amount matched", same_amount

    return MATCHED, "reference matched; amount mismatch", candidates[0]


def compare_records(source_records: list[Record], counterparty_records: list[Record], counterparty_name: str) -> list[OrderedDict[str, Any]]:
    counterparty_index = build_index(counterparty_records)
    rows: list[OrderedDict[str, Any]] = []

    for source in source_records:
        status, reason, counterparty = find_counterparty(source, counterparty_index, counterparty_name)
        counterparty_amount = counterparty.amount_decimal if counterparty else ""
        amount_difference: Decimal | str = ""
        if source.amount_cents is not None and counterparty is not None and counterparty.amount_cents is not None:
            amount_difference = Decimal(source.amount_cents - counterparty.amount_cents) / Decimal(100)

        row: OrderedDict[str, Any] = OrderedDict()
        row["Source_File"] = source.source_file
        row["Source_Row"] = source.source_row
        row["Source_Dataset"] = source.source_dataset
        row["Direction"] = source.direction
        row["Source_DateTime"] = source.datetime_value
        row["Source_Key"] = source.key
        row["Source_Amount"] = source.amount_decimal
        row["Source_Channel_Type"] = source.channel_type
        row["Source_Account"] = source.account
        row["Source_Phone"] = source.phone
        row["Match_Status"] = status
        row["Match_Reason"] = reason
        row["Counterparty_File"] = counterparty.source_file if counterparty else ""
        row["Counterparty_Row"] = counterparty.source_row if counterparty else ""
        row["Counterparty_Dataset"] = counterparty.source_dataset if counterparty else ""
        row["Counterparty_Key"] = counterparty.key if counterparty else ""
        row["Counterparty_Amount"] = counterparty_amount
        row["Amount_Difference"] = amount_difference
        row["Counterparty_Result"] = counterparty.result if counterparty else ""
        row["Counterparty_Transaction_ID"] = counterparty.transaction_id if counterparty else ""
        rows.append(row)

    return rows


def compare_nsano_to_mambu_sources(
    nsano_records: list[Record],
    normal_mambu_records: list[Record],
    write_off_records: list[Record],
    unidentified_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    sources = [
        ("Mambu", MAMBU_STATUS, build_index(normal_mambu_records)),
        ("Write-Off", WRITE_OFF_STATUS, build_index(write_off_records)),
        ("Unidentified", UNIDENTIFIED_STATUS, build_index(unidentified_records)),
    ]
    rows: list[OrderedDict[str, Any]] = []

    for source in nsano_records:
        status = NSANO_NOT_FOUND_STATUS
        reason = "reference not found in Mambu, Write-Off, or Unidentified"
        counterparty: Record | None = None
        counterparty_dataset = ""
        amount_mismatch_candidate: tuple[str, Record] | None = None

        if not source.key:
            reason = "blank reconciliation key"
        else:
            for source_name, source_status, source_index in sources:
                candidates = source_index.get((source.direction, source.key), [])
                if not candidates:
                    continue
                exact_match = next((candidate for candidate in candidates if candidate.amount_cents == source.amount_cents), None)
                if exact_match is not None:
                    status = source_status
                    reason = f"reference and amount matched in {source_name}"
                    counterparty = exact_match
                    counterparty_dataset = source_name
                    break
                status = source_status
                reason = f"reference matched in {source_name}; amount mismatch"
                counterparty = candidates[0]
                counterparty_dataset = source_name
                break

            if counterparty is None and amount_mismatch_candidate is not None:
                counterparty_dataset, counterparty = amount_mismatch_candidate
                reason = "reference matched; amount mismatch"

        counterparty_amount = counterparty.amount_decimal if counterparty else ""
        amount_difference: Decimal | str = ""
        if source.amount_cents is not None and counterparty is not None and counterparty.amount_cents is not None:
            amount_difference = Decimal(source.amount_cents - counterparty.amount_cents) / Decimal(100)

        row: OrderedDict[str, Any] = OrderedDict()
        row["Source_File"] = source.source_file
        row["Source_Row"] = source.source_row
        row["Source_Dataset"] = source.source_dataset
        row["Direction"] = source.direction
        row["Source_DateTime"] = source.datetime_value
        row["Source_Key"] = source.key
        row["Source_Amount"] = source.amount_decimal
        row["Source_Channel_Type"] = source.channel_type
        row["Source_Account"] = source.account
        row["Source_Phone"] = source.phone
        row["Match_Status"] = status
        row["Match_Reason"] = reason
        row["Counterparty_File"] = counterparty.source_file if counterparty else ""
        row["Counterparty_Row"] = counterparty.source_row if counterparty else ""
        row["Counterparty_Dataset"] = counterparty_dataset
        row["Counterparty_Key"] = counterparty.key if counterparty else ""
        row["Counterparty_Amount"] = counterparty_amount
        row["Amount_Difference"] = amount_difference
        row["Counterparty_Result"] = counterparty.result if counterparty else ""
        row["Counterparty_Transaction_ID"] = counterparty.transaction_id if counterparty else ""
        rows.append(row)

    return rows


def build_key_index(records: list[Record]) -> dict[str, list[Record]]:
    index: dict[str, list[Record]] = defaultdict(list)
    for record in records:
        if record.key:
            index[record.key].append(record)
    return index


def compare_records_by_key(
    source_records: list[Record],
    match_sources: list[tuple[str, str, list[Record]]],
    not_found_status: str,
) -> list[OrderedDict[str, Any]]:
    source_indexes = [
        (source_name, status, build_key_index(records))
        for source_name, status, records in match_sources
    ]
    source_names = ", ".join(source_name for source_name, _, _ in source_indexes)
    rows: list[OrderedDict[str, Any]] = []

    for source in source_records:
        status = not_found_status
        reason = f"reference not found in {source_names}"
        counterparty: Record | None = None
        counterparty_dataset = ""

        if not source.key:
            reason = "blank reconciliation key"
        else:
            for source_name, source_status, source_index in source_indexes:
                candidates = source_index.get(source.key, [])
                if not candidates:
                    continue
                status = source_status
                reason = f"reference found in {source_name}"
                counterparty = candidates[0]
                counterparty_dataset = source_name
                break

        counterparty_amount = counterparty.amount_decimal if counterparty else ""
        amount_difference: Decimal | str = ""
        if source.amount_cents is not None and counterparty is not None and counterparty.amount_cents is not None:
            amount_difference = Decimal(source.amount_cents - counterparty.amount_cents) / Decimal(100)

        row: OrderedDict[str, Any] = OrderedDict()
        row["Source_File"] = source.source_file
        row["Source_Row"] = source.source_row
        row["Source_Dataset"] = source.source_dataset
        row["Direction"] = source.direction
        row["Source_DateTime"] = source.datetime_value
        row["Source_Key"] = source.key
        row["Source_Amount"] = source.amount_decimal
        row["Source_Channel_Type"] = source.channel_type
        row["Source_Account"] = source.account
        row["Source_Phone"] = source.phone
        row["Match_Status"] = status
        row["Match_Reason"] = reason
        row["Counterparty_File"] = counterparty.source_file if counterparty else ""
        row["Counterparty_Row"] = counterparty.source_row if counterparty else ""
        row["Counterparty_Dataset"] = counterparty_dataset
        row["Counterparty_Key"] = counterparty.key if counterparty else ""
        row["Counterparty_Amount"] = counterparty_amount
        row["Amount_Difference"] = amount_difference
        row["Counterparty_Result"] = counterparty.result if counterparty else ""
        row["Counterparty_Transaction_ID"] = counterparty.transaction_id if counterparty else ""
        rows.append(row)

    return rows


def compare_mambu_to_itc_sources(
    mambu_records: list[Record],
    itc_records: list[Record],
    vodafone_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    return compare_records_by_key(
        mambu_records,
        [
            ("ITC", ITC_STATUS, itc_records),
            ("Vodafone Collections", VODA_COLL_STATUS, vodafone_records),
        ],
        ITC_NOT_FOUND_STATUS,
    )


def compare_itc_to_mambu_sources(
    itc_records: list[Record],
    mambu_records: list[Record],
    vodafone_records: list[Record],
    unidentified_records: list[Record],
    write_off_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    rows = compare_records_by_key(
        itc_records,
        [
            ("Mambu", ITC_MAMBU_STATUS, mambu_records),
            ("Vodafone Collections", VODA_COLL_STATUS, vodafone_records),
            ("Unidentified", ITC_UNIDENTIFIED_STATUS, unidentified_records),
            ("Write Off", ITC_WRITE_OFF_STATUS, write_off_records),
        ],
        ITC_NOT_FOUND_STATUS,
    )
    for record, row in zip(itc_records, rows):
        narration = str(first_value(record, "narration", "Narration", "ITC Narration")).strip()
        if row.get("Match_Status") == ITC_NOT_FOUND_STATUS and narration.endswith("_3"):
            row["Match_Status"] = ITC_UPSALE_STATUS
            row["Match_Reason"] = "not found; narration _3 classified as upsale"
    return rows


def compare_vodafone_to_mambu_sources(
    vodafone_records: list[Record],
    mambu_records: list[Record],
    unidentified_records: list[Record],
    write_off_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    return compare_records_by_key(
        vodafone_records,
        [
            ("Mambu", ITC_MAMBU_STATUS, mambu_records),
            ("Unidentified", ITC_UNIDENTIFIED_STATUS, unidentified_records),
            ("Write Off", VODAFONE_WRITE_OFF_STATUS, write_off_records),
        ],
        ITC_NOT_FOUND_STATUS,
    )


def _load_write_off_source(path: Path | None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if not path or not path.exists():
        return records, list(WRITE_OFF_SHEET_COLUMNS)

    for source_row, row in iter_source_dicts(path):
        add_unique_headers(headers, row.keys())
        amount_col = _pick_key(row, "Amount (KEY)", "Amount", "Amount (GHC)", "C")
        record = normalize_coll_recon_record(
            path,
            "Write Off",
            source_row,
            row,
            source_dataset="Write Off",
            key_getter=lambda r: _pick(r, "Identifier (Key)", "Transaction ID", "Transactions Id", "Transactions ID"),
            amount_header=amount_col,
            base_headers=[],
            date_getter=lambda r: normalize_date_amount_match_date(
                _pick(r, "Date (KEY)", "Date/Time", "Date")
            ),
            phone_getter=lambda r: _pick(r, "Account Number", "E"),
            account_getter=lambda r: _pick(r, "Loan ID", "G", "Account Number"),
            channel_getter=lambda r: _pick(r, "Channel", "B"),
            transaction_id_getter=lambda r: _pick(r, "Identifier (Key)", "Transaction ID", "Transactions Id", "Transactions ID"),
            include_helpers=False,
        )
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(WRITE_OFF_SHEET_COLUMNS)


def _load_write_off_nsano_target(path: Path | None) -> tuple[list[Record], list[str], int]:
    headers: list[str] = []
    records: list[Record] = []
    failed_rows_ignored = 0
    if not path or not path.exists():
        return records, list(NSANO_SHEET_COLUMNS), failed_rows_ignored

    for source_row, row in iter_source_dicts(path):
        add_unique_headers(headers, row.keys())
        if str(_pick(row, "Result")).strip().casefold() == "failed":
            failed_rows_ignored += 1
            continue
        amount_col = _pick_key(row, "Amount (GHC)", "Amount_GHC")
        record = normalize_coll_recon_record(
            path,
            "Nsano",
            source_row,
            row,
            source_dataset="Nsano",
            key_getter=lambda r: _pick(r, "Ext Debit Ref (Key)", "External_Debit_Reference"),
            amount_header=amount_col,
            base_headers=[],
            date_getter=lambda r: parse_nsano_datetime(_pick(r, "Date/Time", "DateTime")),
            phone_getter=lambda r: _pick(r, "SendingHse_Account", "J"),
            account_getter=lambda r: _pick(r, "SendingHse_ID", "I"),
            channel_getter=lambda r: _pick(r, "Type", "C"),
            result_getter=lambda r: _pick(r, "Result"),
            transaction_id_getter=lambda r: _pick(r, "Trans ID", "Transaction_ID"),
            include_helpers=False,
        )
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(NSANO_SHEET_COLUMNS), failed_rows_ignored


def _load_write_off_itc_target(path: Path | None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if not path or not path.exists():
        return records, list(ITC_SHEET_COLUMNS)

    for source_row, row in iter_source_dicts(path):
        add_unique_headers(headers, row.keys())
        amount_col = _pick_key(row, "amount", "Amount", "Amount (GHC)", "ITC Amount (GHC)")
        record = normalize_coll_recon_record(
            path,
            "ITC",
            source_row,
            row,
            source_dataset="ITC",
            key_getter=lambda r: _pick(r, "channel_transaction_id", "Channel Transaction ID", "ITC Trans ID (Key)"),
            amount_header=amount_col,
            base_headers=[],
            date_getter=lambda r: excel_serial_to_iso(_pick(r, "transaction_date", "Transaction Date", "ITC Date")),
            phone_getter=lambda r: _pick(r, "payer_contact", "account_reference"),
            account_getter=lambda r: _pick(r, "account_reference"),
            channel_getter=lambda r: _pick(r, "channel"),
            transaction_id_getter=lambda r: _pick(r, "channel_transaction_id", "Channel Transaction ID", "ITC Trans ID (Key)"),
            include_helpers=False,
        )
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(ITC_SHEET_COLUMNS)


def _load_write_off_zenith_target(path: Path | None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if not path or not path.exists():
        return records, list(MAMBU_SHEET_COLUMNS)

    # Zenith bank statements often contain title/balance rows before the actual
    # transaction header. Reuse the same header detection and cleaning routine
    # as the successful Zenith collection reconciliation.
    from zenith_collection_reconciliation import _cleaned_zenith_source_rows

    for source_row, row in _cleaned_zenith_source_rows(path):
        add_unique_headers(headers, row.keys())
        # Keep the standalone write-off investigation aligned with the Zenith
        # collection reconciliation: Zenith is keyed by Create Date + Credit.
        amount_col = _pick_key(
            row,
            "Credit",
            "Credit (KEY amount)",
            "Paid In",
            "Amount",
            "Amount (GHC)",
            "Amount (KEY)",
        )
        record = normalize_coll_recon_record(
            path,
            "Zenith Collections",
            source_row,
            row,
            source_dataset="Zenith",
            key_getter=lambda r: date_amount_match_key(
                _pick(r, "Create Date (KEY date - DD/MM/YYYY)", "Create Date", "Date"),
                r.get(amount_col, ""),
            ),
            amount_header=amount_col,
            base_headers=[],
            date_getter=lambda r: normalize_date_amount_match_date(
                _pick(r, "Create Date (KEY date - DD/MM/YYYY)", "Create Date", "Date")
            ),
            phone_getter=lambda r: _pick(r, "Mobile Phone (Client)", "Description", "Reference"),
            account_getter=lambda r: _pick(r, "Account ID", "Account Number"),
            channel_getter=lambda r: _pick(r, "Channel"),
            transaction_id_getter=lambda r: date_amount_match_key(
                _pick(r, "Create Date (KEY date - DD/MM/YYYY)", "Create Date", "Date"),
                r.get(amount_col, ""),
            ),
            include_helpers=False,
        )
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(MAMBU_SHEET_COLUMNS)


def _load_write_off_vodafone_target(path: Path | None) -> tuple[list[Record], list[str]]:
    headers: list[str] = []
    records: list[Record] = []
    if not path or not path.exists():
        return records, list(VODAFONE_COLLECTIONS_SHEET_COLUMNS)

    vodafone_sheet_name = xlsx_sheet_if_present(path, "Cleaned Voda Coll")
    for source_row, row in iter_source_dicts(path, vodafone_sheet_name):
        add_unique_headers(headers, row.keys())
        amount_col = _pick_key(row, "Paid In", "PaidIn", "Amount", "Amount (GHC)")
        record = normalize_coll_recon_record(
            path,
            "Vodafone Collections",
            source_row,
            row,
            source_dataset="Vodafone Collections",
            key_getter=lambda r: _pick(r, "Receipt No.", "Receipt No", "Receipt Number"),
            amount_header=amount_col,
            base_headers=[],
            date_getter=lambda r: excel_serial_to_iso(_pick(r, "Completion Time", "Initiation Time", "Date")),
            phone_getter=lambda r: _pick(r, "Opposite Party", "Details"),
            account_getter=lambda r: _pick(r, "Opposite Party", "Details"),
            channel_getter=lambda r: _pick(r, "Reason Type", "Details"),
            transaction_id_getter=lambda r: _pick(r, "Receipt No.", "Receipt No", "Receipt Number"),
            include_helpers=False,
        )
        for header in headers:
            record.sheet_values.setdefault(header, "")
        records.append(record)
    return records, headers or list(VODAFONE_COLLECTIONS_SHEET_COLUMNS)


def load_write_off_recon_sources(
    write_off_path: Path | None,
    nsano_path: Path | None,
    itc_path: Path | None,
    zenith_path: Path | None,
    vodafone_path: Path | None,
) -> tuple[
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    list[Record],
    list[str],
    int,
]:
    write_off_records, write_off_headers = _load_write_off_source(write_off_path)
    nsano_records, nsano_headers, nsano_failed_rows_ignored = _load_write_off_nsano_target(nsano_path)
    itc_records, itc_headers = _load_write_off_itc_target(itc_path)
    zenith_records, zenith_headers = _load_write_off_zenith_target(zenith_path)
    vodafone_records, vodafone_headers = _load_write_off_vodafone_target(vodafone_path)
    return (
        write_off_records,
        write_off_headers,
        nsano_records,
        nsano_headers,
        itc_records,
        itc_headers,
        zenith_records,
        zenith_headers,
        vodafone_records,
        vodafone_headers,
        nsano_failed_rows_ignored,
    )


def compare_write_off_to_collection_sources(
    write_off_records: list[Record],
    nsano_records: list[Record],
    itc_records: list[Record],
    zenith_records: list[Record],
    vodafone_records: list[Record],
) -> list[OrderedDict[str, Any]]:
    source_indexes = {
        WRITE_OFF_NSANO_STATUS: build_key_index(nsano_records),
        WRITE_OFF_ITC_STATUS: build_key_index(itc_records),
        WRITE_OFF_ZENITH_STATUS: build_key_index(zenith_records),
        WRITE_OFF_VODAFONE_STATUS: build_key_index(vodafone_records),
    }
    preferred_vodafone_matches = {
        key: next(
            (
                candidate
                for candidate in candidates
                if candidate.amount_cents not in (None, 0)
            ),
            candidates[0],
        )
        for key, candidates in source_indexes[WRITE_OFF_VODAFONE_STATUS].items()
        if candidates
    }
    zenith_by_date: defaultdict[str, list[Record]] = defaultdict(list)
    zenith_by_amount: defaultdict[int, list[Record]] = defaultdict(list)
    for record in zenith_records:
        if record.datetime_value:
            zenith_by_date[normalize_date_amount_match_date(record.datetime_value)].append(record)
        if record.amount_cents is not None:
            zenith_by_amount[record.amount_cents].append(record)
    rows: list[OrderedDict[str, Any]] = []

    for source in write_off_records:
        routed_wallet = _write_off_routed_wallet(source.channel_type)
        matches: dict[str, Record] = {}
        match: Record | None = None
        matched_wallet = ""

        # Identifier-based wallets. Zenith-routed write-offs must never use
        # identifiers; Zenith is compared exclusively by date + amount.
        if source.key and routed_wallet != WRITE_OFF_ZENITH_STATUS:
            for wallet in WRITE_OFF_IDENTIFIER_WALLETS:
                candidates = source_indexes[wallet].get(source.key, [])
                if candidates:
                    if wallet == WRITE_OFF_VODAFONE_STATUS:
                        # A Vodafone receipt commonly has two legitimate rows:
                        # a Withdrawn charge followed by the Paid In collection.
                        # Use the collection row for the wallet amount.
                        matches[wallet] = preferred_vodafone_matches[source.key]
                    else:
                        matches[wallet] = candidates[0]

        # Zenith has no stable identifier. Only Zenith-routed write-offs may be
        # compared with Zenith, exclusively by date + amount.
        if routed_wallet == WRITE_OFF_ZENITH_STATUS:
            zenith_key = date_amount_match_key(source.datetime_value, source.amount_decimal)
            if zenith_key:
                candidates = source_indexes[WRITE_OFF_ZENITH_STATUS].get(zenith_key, [])
                if candidates:
                    matches[WRITE_OFF_ZENITH_STATUS] = candidates[0]

            source_date = normalize_date_amount_match_date(source.datetime_value)
            same_date_zenith = zenith_by_date.get(source_date, []) if source_date else []
            same_amount_zenith = (
                zenith_by_amount.get(source.amount_cents, [])
                if source.amount_cents is not None
                else []
            )
            zenith_exact = matches.get(WRITE_OFF_ZENITH_STATUS)
            if zenith_exact is not None:
                zenith_date_check = "Matched"
                zenith_amount_check = "Matched"
                zenith_mismatch_detail = "Exact date and amount match"
                zenith_candidate = zenith_exact
            elif same_date_zenith:
                zenith_date_check = "Matched"
                zenith_amount_check = "Mismatch"
                zenith_candidate = same_date_zenith[0]
                zenith_mismatch_detail = (
                    f"Zenith Amount Mismatch · write-off {source.amount_decimal} "
                    f"vs wallet {zenith_candidate.amount_decimal}"
                )
            elif same_amount_zenith:
                zenith_date_check = "Mismatch"
                zenith_amount_check = "Matched"
                zenith_candidate = same_amount_zenith[0]
                zenith_mismatch_detail = (
                    f"Zenith Date Mismatch · write-off {source_date or 'blank'} "
                    f"vs wallet {zenith_candidate.datetime_value or 'blank'}"
                )
            else:
                zenith_date_check = "No Match"
                zenith_amount_check = "No Match"
                zenith_mismatch_detail = "No Zenith row matched either date or amount"
                zenith_candidate = None
        else:
            zenith_exact = None
            zenith_date_check = "Not Applicable"
            zenith_amount_check = "Not Applicable"
            zenith_mismatch_detail = "Not a Zenith-routed write-off"
            zenith_candidate = None

        candidate_wallets = list(matches)
        if routed_wallet and routed_wallet in matches:
            matched_wallet = routed_wallet
            match = matches[routed_wallet]
            investigation = "Exact Match"
            reason = f"channel routed to {routed_wallet}; match found in routed wallet"
            if len(matches) > 1:
                investigation = "Multiple Wallet Matches — Routed by Channel"
                other_wallets = ", ".join(wallet for wallet in matches if wallet != routed_wallet)
                reason += f"; also found in {other_wallets}"
        elif len(matches) == 1:
            matched_wallet, match = next(iter(matches.items()))
            if routed_wallet and routed_wallet != matched_wallet:
                investigation = "Channel Mismatch"
                reason = f"channel routes to {routed_wallet}, but transaction was found in {matched_wallet}"
            else:
                investigation = "Exact Match"
                reason = f"transaction found in {matched_wallet}"
        elif len(matches) > 1:
            matched_wallet = WRITE_OFF_AMBIGUOUS_STATUS
            investigation = "Ambiguous — Multiple Wallet Matches"
            reason = f"transaction found in multiple wallets: {', '.join(matches)}"
        else:
            matched_wallet = WRITE_OFF_NOT_FOUND_STATUS
            investigation = "Not Found"
            reason = "transaction not found in any collection wallet"

        if routed_wallet == WRITE_OFF_ZENITH_STATUS and zenith_exact is None:
            if zenith_amount_check == "Mismatch":
                # A same-date Zenith candidate is still allocated to Zenith.
                # The differing amount is an investigation outcome, not a
                # not-found allocation.
                matched_wallet = WRITE_OFF_ZENITH_STATUS
                investigation = "Amount Mismatch"
            else:
                investigation = f"{investigation}; {zenith_mismatch_detail.split(' · ', 1)[0]}"
            reason += f"; {zenith_mismatch_detail}"

        # Retain a Zenith same-date candidate for amount-variance reporting
        # without treating it as an exact date-and-amount allocation match.
        reported_match = match
        if (
            reported_match is None
            and routed_wallet == WRITE_OFF_ZENITH_STATUS
            and zenith_amount_check == "Mismatch"
        ):
            reported_match = zenith_candidate

        wallet_amount = reported_match.amount_decimal if reported_match else None
        amount_variance: Decimal | str = ""
        if (
            reported_match is not None
            and source.amount_cents is not None
            and reported_match.amount_cents is not None
        ):
            amount_variance = Decimal(
                source.amount_cents - reported_match.amount_cents
            ) / Decimal(100)
            if amount_variance:
                investigation = "Amount Mismatch"

        row: OrderedDict[str, Any] = OrderedDict()
        row["Write Off Date"] = source.datetime_value
        row["Write Off Identifier"] = source.key
        row["Write Off Amount (GHC)"] = source.amount_decimal
        row["Write Off Channel"] = source.channel_type
        row["Routed Wallet"] = routed_wallet
        row["Write Off Account"] = source.account
        row["Match Status"] = matched_wallet
        row["Investigation Status"] = investigation
        row["Match Reason"] = reason
        row["Candidate Wallets"] = ", ".join(candidate_wallets)
        row["Zenith Date Check"] = zenith_date_check
        row["Zenith Amount Check"] = zenith_amount_check
        row["Zenith Mismatch Detail"] = zenith_mismatch_detail
        row["Zenith Candidate Date"] = zenith_candidate.datetime_value if zenith_candidate else ""
        row["Zenith Candidate Amount (GHC)"] = (
            zenith_candidate.amount_decimal if zenith_candidate else ""
        )
        row["Matched Identifier"] = reported_match.key if reported_match else ""
        row["Wallet Amount (GHC)"] = wallet_amount if wallet_amount is not None else ""
        row["Amount Variance (GHC)"] = amount_variance
        row["Matched Source Row"] = reported_match.source_row if reported_match else ""
        row["Matched Source File"] = reported_match.source_file if reported_match else ""
        row["Matched Source"] = (
            matched_wallet
            if match
            else ("Zenith amount-mismatch candidate" if reported_match else "")
        )
        rows.append(row)

    return rows


def _write_off_routed_wallet(channel: Any) -> str:
    """Map a write-off channel label to its collection wallet."""
    normalized = re.sub(r"[^a-z0-9]+", " ", str(channel or "").casefold()).strip()
    if "nsano" in normalized:
        return WRITE_OFF_NSANO_STATUS
    if "zenith" in normalized:
        return WRITE_OFF_ZENITH_STATUS
    if "vodafone" in normalized or re.search(r"\bvoda\b", normalized):
        return WRITE_OFF_VODAFONE_STATUS
    if re.search(r"\bitc\b", normalized):
        return WRITE_OFF_ITC_STATUS
    return ""


def summarize_write_off_recon(rows: list[OrderedDict[str, Any]]) -> dict[str, Any]:
    status_counts = Counter(str(row.get("Match Status", "")) for row in rows)
    investigation_counts = Counter(str(row.get("Investigation Status", "")) for row in rows)
    status_amount_cents: defaultdict[str, int] = defaultdict(int)
    status_wallet_amount_cents: defaultdict[str, int] = defaultdict(int)
    ambiguous_breakdown: Counter[str] = Counter()
    source_amount_cents = 0
    wallet_amount_cents = 0

    for row in rows:
        status = str(row.get("Match Status", ""))
        amount_cents = amount_to_cents(row.get("Write Off Amount (GHC)", "")) or 0
        matched_cents = amount_to_cents(row.get("Wallet Amount (GHC)", "")) or 0
        status_amount_cents[status] += amount_cents
        status_wallet_amount_cents[status] += matched_cents
        source_amount_cents += amount_cents
        wallet_amount_cents += matched_cents
        if status == WRITE_OFF_AMBIGUOUS_STATUS:
            channel = str(row.get("Write Off Channel", "") or "Blank / Unrecognized")
            candidates = str(row.get("Candidate Wallets", "") or "None")
            ambiguous_breakdown[f"Channel: {channel} · Candidates: {candidates}"] += 1

    total = len(rows)
    matched_statuses = {
        WRITE_OFF_NSANO_STATUS,
        WRITE_OFF_ITC_STATUS,
        WRITE_OFF_ZENITH_STATUS,
        WRITE_OFF_VODAFONE_STATUS,
    }
    matched = sum(status_counts[status] for status in matched_statuses)
    ambiguous = status_counts[WRITE_OFF_AMBIGUOUS_STATUS]
    return {
        "total": total,
        "matched": matched,
        "not_found": status_counts[WRITE_OFF_NOT_FOUND_STATUS],
        "ambiguous": ambiguous,
        "match_rate": matched / total if total else 0,
        "source_amount": cents_to_decimal(source_amount_cents),
        "wallet_amount": cents_to_decimal(wallet_amount_cents),
        "amount_variance": cents_to_decimal(source_amount_cents - wallet_amount_cents),
        "status_counts": dict(status_counts),
        "investigation_counts": dict(investigation_counts),
        "ambiguous_breakdown": dict(ambiguous_breakdown),
        "status_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_amount_cents.items()
        },
        "status_wallet_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_wallet_amount_cents.items()
        },
        "status_variances": {
            status: cents_to_decimal(
                status_amount_cents[status] - status_wallet_amount_cents[status]
            )
            for status in set(status_amount_cents) | set(status_wallet_amount_cents)
        },
    }


def build_write_off_recon_summary_rows(stats: dict[str, Any]) -> list[list[Any]]:
    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_COUNT)

    def tm(value: Any) -> Cell:
        return Cell(value, STYLE_TABLE_MONEY)

    def count(status: str) -> int:
        return int(stats["status_counts"].get(status, 0))

    def write_off_amount(status: str) -> Decimal:
        return stats["status_amounts"].get(status, Decimal("0"))

    def wallet_amount(status: str) -> Decimal:
        return stats["status_wallet_amounts"].get(status, Decimal("0"))

    def variance(status: str) -> Decimal:
        return stats["status_variances"].get(status, Decimal("0"))

    statuses = (
        (WRITE_OFF_NSANO_STATUS, STYLE_MATCHED),
        (WRITE_OFF_ITC_STATUS, STYLE_MATCHED),
        (WRITE_OFF_ZENITH_STATUS, STYLE_UNIDENTIFIED),
        (WRITE_OFF_VODAFONE_STATUS, STYLE_UNIDENTIFIED),
        (WRITE_OFF_AMBIGUOUS_STATUS, STYLE_NOT_FOUND),
        (WRITE_OFF_NOT_FOUND_STATUS, STYLE_NOT_FOUND),
    )
    rows: list[list[Any]] = [
        [Cell("WRITE-OFF RECONCILIATION — SUMMARY", STYLE_TITLE), "", "", "", ""],
        [
            Cell(
                f"Report logic {REPORT_LOGIC_VERSION} · generated {dt.datetime.now().astimezone():%Y-%m-%d %H:%M %Z}",
                STYLE_DASH_MUTED,
            ),
            "", "", "", "",
        ],
        [
            th("Allocated Wallet"),
            th("Count"),
            th("Write-off Amount (GHC)"),
            th("Wallet Amount (GHC)"),
            th("Variance (GHC)"),
        ],
    ]
    rows.extend(
        [
            Cell(status, style),
            count(status),
            Cell(write_off_amount(status), STYLE_MONEY),
            Cell(wallet_amount(status), STYLE_MONEY),
            Cell(variance(status), STYLE_MONEY),
        ]
        for status, style in statuses
    )
    rows.append([
        th("TOTAL"),
        tc(stats["total"]),
        tm(stats["source_amount"]),
        tm(stats["wallet_amount"]),
        tm(stats["amount_variance"]),
    ])
    rows.extend([
        ["", "", "", "", ""],
        [th("Investigation Status"), th("Count"), "", "", ""],
    ])
    for label, count_value in sorted(stats.get("investigation_counts", {}).items()):
        rows.append([Cell(label, STYLE_MASTER_BODY), int(count_value), "", "", ""])
    if stats.get("ambiguous_breakdown"):
        rows.extend([
            ["", "", "", "", ""],
            [th("Ambiguous Breakdown"), th("Count"), "", "", ""],
        ])
        for label, count_value in sorted(stats["ambiguous_breakdown"].items()):
            rows.append([Cell(label, STYLE_MASTER_BODY), int(count_value), "", "", ""])
    return rows


def _write_off_mismatch_analysis_rows(
    comparison_rows: list[OrderedDict[str, Any]],
) -> tuple[list[list[Any]], list[str]]:
    """Build separate amount-mismatch and not-found analyses."""
    mismatches = [
        row
        for row in comparison_rows
        if str(row.get("Match Status", "")) != WRITE_OFF_NOT_FOUND_STATUS
        if (amount_to_cents(row.get("Amount Variance (GHC)", "")) or 0) != 0
    ]
    not_found_rows = [
        row
        for row in comparison_rows
        if str(row.get("Match Status", "")) == WRITE_OFF_NOT_FOUND_STATUS
    ]
    wallet_order = {
        WRITE_OFF_NSANO_STATUS: 0,
        WRITE_OFF_ITC_STATUS: 1,
        WRITE_OFF_VODAFONE_STATUS: 2,
        WRITE_OFF_ZENITH_STATUS: 3,
    }
    mismatches.sort(
        key=lambda row: (
            wallet_order.get(str(row.get("Match Status", "")), 99),
            str(row.get("Write Off Identifier", "")),
        )
    )

    grouped: dict[str, dict[str, Any]] = {}
    for row in mismatches:
        wallet = str(
            (
                row.get("Routed Wallet")
                if row.get("Investigation Status") == "Amount Mismatch"
                else row.get("Match Status")
            )
            or row.get("Match Status")
            or "Unallocated"
        )
        bucket = grouped.setdefault(
            wallet,
            {
                "count": 0,
                "write_off_amount": Decimal("0"),
                "wallet_amount": Decimal("0"),
                "variance": Decimal("0"),
            },
        )
        bucket["count"] += 1
        bucket["write_off_amount"] += cents_to_decimal(
            amount_to_cents(row.get("Write Off Amount (GHC)", "")) or 0
        )
        bucket["wallet_amount"] += cents_to_decimal(
            amount_to_cents(row.get("Wallet Amount (GHC)", "")) or 0
        )
        bucket["variance"] += cents_to_decimal(
            amount_to_cents(row.get("Amount Variance (GHC)", "")) or 0
        )

    rows: list[list[Any]] = [
        [Cell("WRITE-OFF AMOUNT MISMATCH ANALYSIS", STYLE_TITLE), *("" for _ in range(9))],
        [
            Cell(
                "Signed variance = Write-off Amount − Wallet Amount · grouped by allocated wallet",
                STYLE_DASH_MUTED,
            ),
            *("" for _ in range(9)),
        ],
        ["" for _ in range(10)],
        [
            Cell("Allocated Wallet", STYLE_TABLE_SECTION),
            Cell("Mismatch Count", STYLE_TABLE_SECTION),
            Cell("Write-off Amount (GHC)", STYLE_TABLE_SECTION),
            Cell("Wallet Amount (GHC)", STYLE_TABLE_SECTION),
            Cell("Variance (GHC)", STYLE_TABLE_SECTION),
            *("" for _ in range(5)),
        ],
    ]
    for wallet in sorted(grouped, key=lambda value: wallet_order.get(value, 99)):
        summary = grouped[wallet]
        rows.append([
            Cell(wallet, STYLE_MASTER_BODY),
            Cell(summary["count"], STYLE_TABLE_COUNT),
            Cell(summary["write_off_amount"], STYLE_TABLE_MONEY),
            Cell(summary["wallet_amount"], STYLE_TABLE_MONEY),
            Cell(summary["variance"], STYLE_TABLE_MONEY),
            *("" for _ in range(5)),
        ])

    total_count = sum(summary["count"] for summary in grouped.values())
    total_write_off = sum(
        (summary["write_off_amount"] for summary in grouped.values()),
        Decimal("0"),
    )
    total_wallet = sum(
        (summary["wallet_amount"] for summary in grouped.values()),
        Decimal("0"),
    )
    total_variance = sum(
        (summary["variance"] for summary in grouped.values()),
        Decimal("0"),
    )
    rows.extend([
        [
            Cell("TOTAL", STYLE_TABLE_SECTION),
            Cell(total_count, STYLE_TABLE_COUNT),
            Cell(total_write_off, STYLE_TABLE_MONEY),
            Cell(total_wallet, STYLE_TABLE_MONEY),
            Cell(total_variance, STYLE_TABLE_MONEY),
            *("" for _ in range(5)),
        ],
        ["" for _ in range(10)],
        [
            Cell("Allocated Wallet", STYLE_TABLE_SECTION),
            Cell("Write-off Identifier", STYLE_TABLE_SECTION),
            Cell("Write-off Channel", STYLE_TABLE_SECTION),
            Cell("Write-off Date", STYLE_TABLE_SECTION),
            Cell("Write-off Amount", STYLE_TABLE_SECTION),
            Cell("Wallet Amount", STYLE_TABLE_SECTION),
            Cell("Variance", STYLE_TABLE_SECTION),
            Cell("Candidate Wallets", STYLE_TABLE_SECTION),
            Cell("Investigation Status", STYLE_TABLE_SECTION),
            Cell("Match Reason", STYLE_TABLE_SECTION),
        ],
    ])
    for row in mismatches:
        analysis_wallet = str(
            (
                row.get("Routed Wallet")
                if row.get("Investigation Status") == "Amount Mismatch"
                else row.get("Match Status")
            )
            or row.get("Match Status")
            or "Unallocated"
        )
        rows.append([
            Cell(analysis_wallet, STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Identifier", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Channel", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Date", "")), STYLE_MASTER_BODY),
            Cell(row.get("Write Off Amount (GHC)", ""), STYLE_TABLE_MONEY),
            Cell(row.get("Wallet Amount (GHC)", ""), STYLE_TABLE_MONEY),
            Cell(row.get("Amount Variance (GHC)", ""), STYLE_TABLE_MONEY),
            Cell(str(row.get("Candidate Wallets", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Investigation Status", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Match Reason", "")), STYLE_MASTER_BODY),
        ])
    if not mismatches:
        rows.append([
            Cell("No amount mismatches found", STYLE_DASH_MUTED),
            *("" for _ in range(9)),
        ])

    not_found_grouped: dict[str, dict[str, Any]] = {}
    for row in not_found_rows:
        wallet = str(row.get("Routed Wallet", "") or "Unallocated")
        bucket = not_found_grouped.setdefault(
            wallet,
            {"count": 0, "write_off_amount": Decimal("0")},
        )
        bucket["count"] += 1
        bucket["write_off_amount"] += cents_to_decimal(
            amount_to_cents(row.get("Write Off Amount (GHC)", "")) or 0
        )

    rows.append(["" for _ in range(10)])
    not_found_title_row = len(rows) + 1
    rows.extend([
        [Cell("WRITE-OFF NOT FOUND ANALYSIS", STYLE_TITLE), *("" for _ in range(9))],
        [
            Cell(
                "Grouped by the wallet indicated by the write-off channel · these records were not found in a collection wallet",
                STYLE_DASH_MUTED,
            ),
            *("" for _ in range(9)),
        ],
        ["" for _ in range(10)],
        [
            Cell("Routed Wallet", STYLE_TABLE_SECTION),
            Cell("Not Found Count", STYLE_TABLE_SECTION),
            Cell("Write-off Amount (GHC)", STYLE_TABLE_SECTION),
            *("" for _ in range(7)),
        ],
    ])
    for wallet in sorted(
        not_found_grouped,
        key=lambda value: wallet_order.get(value, 99),
    ):
        summary = not_found_grouped[wallet]
        rows.append([
            Cell(wallet, STYLE_MASTER_BODY),
            Cell(summary["count"], STYLE_TABLE_COUNT),
            Cell(summary["write_off_amount"], STYLE_TABLE_MONEY),
            *("" for _ in range(7)),
        ])

    not_found_total_count = sum(
        summary["count"] for summary in not_found_grouped.values()
    )
    not_found_total_amount = sum(
        (
            summary["write_off_amount"]
            for summary in not_found_grouped.values()
        ),
        Decimal("0"),
    )
    not_found_total_row = len(rows) + 1
    rows.extend([
        [
            Cell("TOTAL", STYLE_TABLE_SECTION),
            Cell(not_found_total_count, STYLE_TABLE_COUNT),
            Cell(not_found_total_amount, STYLE_TABLE_MONEY),
            *("" for _ in range(7)),
        ],
        ["" for _ in range(10)],
        [
            Cell("Routed Wallet", STYLE_TABLE_SECTION),
            Cell("Write-off Identifier", STYLE_TABLE_SECTION),
            Cell("Write-off Channel", STYLE_TABLE_SECTION),
            Cell("Write-off Date", STYLE_TABLE_SECTION),
            Cell("Write-off Amount", STYLE_TABLE_SECTION),
            Cell("Zenith Date Check", STYLE_TABLE_SECTION),
            Cell("Zenith Amount Check", STYLE_TABLE_SECTION),
            Cell("Candidate Wallets", STYLE_TABLE_SECTION),
            Cell("Investigation Status", STYLE_TABLE_SECTION),
            Cell("Match Reason", STYLE_TABLE_SECTION),
        ],
    ])
    for row in not_found_rows:
        rows.append([
            Cell(str(row.get("Routed Wallet", "") or "Unallocated"), STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Identifier", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Channel", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Write Off Date", "")), STYLE_MASTER_BODY),
            Cell(row.get("Write Off Amount (GHC)", ""), STYLE_TABLE_MONEY),
            Cell(str(row.get("Zenith Date Check", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Zenith Amount Check", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Candidate Wallets", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Investigation Status", "")), STYLE_MASTER_BODY),
            Cell(str(row.get("Match Reason", "")), STYLE_MASTER_BODY),
        ])
    if not not_found_rows:
        rows.append([
            Cell("No not-found records", STYLE_DASH_MUTED),
            *("" for _ in range(9)),
        ])

    return rows, [
        "A1:J1",
        "A2:J2",
        "E4:J4",
        f"E{5 + len(grouped)}:J{5 + len(grouped)}",
        f"A{not_found_title_row}:J{not_found_title_row}",
        f"A{not_found_title_row + 1}:J{not_found_title_row + 1}",
        f"D{not_found_title_row + 3}:J{not_found_title_row + 3}",
        f"D{not_found_total_row}:J{not_found_total_row}",
    ]


def write_write_off_recon_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    write_off_headers: list[str],
    write_off_records: list[Record],
    nsano_headers: list[str],
    nsano_records: list[Record],
    itc_headers: list[str],
    itc_records: list[Record],
    zenith_headers: list[str],
    zenith_records: list[Record],
    vodafone_headers: list[str],
    vodafone_records: list[Record],
    comparison_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Summary",
        "Write Off Recon",
        "Zenith Investigation",
        "Ambiguous Investigation",
        "Amount Mismatch Analysis",
        "Write Off",
        "Nsano",
        "ITC",
        "Zenith Collections",
        "Vodafone Collections",
    ]

    write_off_preview = ([record.sheet_values.get(header, "") for header in write_off_headers] for record in write_off_records[:5000])
    nsano_preview = ([record.sheet_values.get(header, "") for header in nsano_headers] for record in nsano_records[:5000])
    itc_preview = ([record.sheet_values.get(header, "") for header in itc_headers] for record in itc_records[:5000])
    zenith_preview = ([record.sheet_values.get(header, "") for header in zenith_headers] for record in zenith_records[:5000])
    vodafone_preview = ([record.sheet_values.get(header, "") for header in vodafone_headers] for record in vodafone_records[:5000])
    comparison_preview = ([row.get(header, "") for header in WRITE_OFF_RECON_HEADERS] for row in comparison_rows[:5000])
    zenith_investigation_rows = [
        row
        for row in comparison_rows
        if row.get("Routed Wallet") == WRITE_OFF_ZENITH_STATUS
        or row.get("Match Status") == WRITE_OFF_ZENITH_STATUS
    ]
    zenith_investigation_preview = (
        [row.get(header, "") for header in WRITE_OFF_RECON_HEADERS]
        for row in zenith_investigation_rows[:5000]
    )
    ambiguous_investigation_rows = [
        row
        for row in comparison_rows
        if row.get("Match Status") == WRITE_OFF_AMBIGUOUS_STATUS
    ]
    ambiguous_investigation_preview = (
        [row.get(header, "") for header in WRITE_OFF_RECON_HEADERS]
        for row in ambiguous_investigation_rows[:5000]
    )
    mismatch_analysis_rows, mismatch_analysis_merges = _write_off_mismatch_analysis_rows(
        comparison_rows
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
            summary_rows,
            len(summary_rows),
            5,
            [38, 14, 24, 24, 22],
            freeze_top_row=False,
            autofilter=False,
            merges=["A1:E1", "A2:E2"],
            style_func=summary_style,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            rows_from_dicts(WRITE_OFF_RECON_HEADERS, comparison_rows),
            len(comparison_rows) + 1,
            len(WRITE_OFF_RECON_HEADERS),
            compute_widths(WRITE_OFF_RECON_HEADERS, comparison_preview),
            style_func=data_style(WRITE_OFF_RECON_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            rows_from_dicts(WRITE_OFF_RECON_HEADERS, zenith_investigation_rows),
            len(zenith_investigation_rows) + 1,
            len(WRITE_OFF_RECON_HEADERS),
            compute_widths(WRITE_OFF_RECON_HEADERS, zenith_investigation_preview),
            style_func=data_style(WRITE_OFF_RECON_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            rows_from_dicts(WRITE_OFF_RECON_HEADERS, ambiguous_investigation_rows),
            len(ambiguous_investigation_rows) + 1,
            len(WRITE_OFF_RECON_HEADERS),
            compute_widths(WRITE_OFF_RECON_HEADERS, ambiguous_investigation_preview),
            style_func=data_style(WRITE_OFF_RECON_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            mismatch_analysis_rows,
            len(mismatch_analysis_rows),
            10,
            [24, 25, 20, 18, 20, 20, 18, 25, 30, 55],
            freeze_top_row=False,
            autofilter=False,
            merges=mismatch_analysis_merges,
            style_func=None,
            show_gridlines=False,
            zoom_scale=90,
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            record_rows(write_off_headers, write_off_records),
            len(write_off_records) + 1,
            len(write_off_headers),
            compute_widths(write_off_headers, write_off_preview),
            style_func=data_style(write_off_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            record_rows(nsano_headers, nsano_records),
            len(nsano_records) + 1,
            len(nsano_headers),
            compute_widths(nsano_headers, nsano_preview),
            style_func=data_style(nsano_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet8.xml",
            record_rows(itc_headers, itc_records),
            len(itc_records) + 1,
            len(itc_headers),
            compute_widths(itc_headers, itc_preview),
            style_func=data_style(itc_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet9.xml",
            record_rows(zenith_headers, zenith_records),
            len(zenith_records) + 1,
            len(zenith_headers),
            compute_widths(zenith_headers, zenith_preview),
            style_func=data_style(zenith_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet10.xml",
            record_rows(vodafone_headers, vodafone_records),
            len(vodafone_records) + 1,
            len(vodafone_headers),
            compute_widths(vodafone_headers, vodafone_preview),
            style_func=data_style(vodafone_headers),
        )

    os.replace(temp_path, output_path)


def build_write_off_reconciliation(
    output_path: Path,
    write_off_records: list[Record],
    write_off_headers: list[str],
    nsano_records: list[Record],
    nsano_headers: list[str],
    itc_records: list[Record],
    itc_headers: list[str],
    zenith_records: list[Record],
    zenith_headers: list[str],
    vodafone_records: list[Record],
    vodafone_headers: list[str],
) -> tuple[dict[str, Any], list[str]]:
    comparison_rows = compare_write_off_to_collection_sources(
        write_off_records,
        nsano_records,
        itc_records,
        zenith_records,
        vodafone_records,
    )
    stats = summarize_write_off_recon(comparison_rows)
    summary_rows = build_write_off_recon_summary_rows(stats)
    write_write_off_recon_workbook(
        output_path,
        summary_rows,
        write_off_headers,
        write_off_records,
        nsano_headers,
        nsano_records,
        itc_headers,
        itc_records,
        zenith_headers,
        zenith_records,
        vodafone_headers,
        vodafone_records,
        comparison_rows,
    )
    write_off_itc_keys = sorted({
        str(row["Write Off Identifier"])
        for row in comparison_rows
        if row.get("Match Status") == WRITE_OFF_ITC_STATUS and row.get("Write Off Identifier")
    })
    return stats, write_off_itc_keys


def summarize_compare(rows: list[OrderedDict[str, Any]]) -> dict[str, Any]:
    total = len(rows)
    matched_statuses = {MATCHED, MAMBU_STATUS, WRITE_OFF_STATUS, UNIDENTIFIED_STATUS}
    status_counts = Counter(row["Match_Status"] for row in rows)
    matched = sum(1 for row in rows if row["Match_Status"] in matched_statuses)
    not_found = total - matched
    amount_mismatch = sum(1 for row in rows if "amount mismatch" in str(row["Match_Reason"]).casefold())
    blank_key = sum(1 for row in rows if row["Match_Reason"] == "blank reconciliation key")
    source_amount_cents = sum(amount_to_cents(row["Source_Amount"]) or 0 for row in rows)
    matched_amount_cents = sum(
        (amount_to_cents(row["Source_Amount"]) or 0)
        for row in rows
        if row["Match_Status"] in matched_statuses
    )

    matched_exact_cents = 0
    amount_mismatch_cents = 0
    not_found_cents = 0
    mambu_cents = 0
    write_off_cents = 0
    unidentified_cents = 0

    for row in rows:
        status = row["Match_Status"]
        reason = str(row["Match_Reason"]).casefold()
        cents = amount_to_cents(row["Source_Amount"]) or 0
        if status == MATCHED:
            if "amount mismatch" in reason:
                amount_mismatch_cents += cents
            else:
                matched_exact_cents += cents
        elif status == MAMBU_STATUS:
            mambu_cents += cents
        elif status == WRITE_OFF_STATUS:
            write_off_cents += cents
        elif status == UNIDENTIFIED_STATUS:
            unidentified_cents += cents
        else:
            not_found_cents += cents

    return {
        "total": total,
        "matched": matched,
        "matched_exact": matched - amount_mismatch,
        "not_found": not_found,
        "mambu": status_counts[MAMBU_STATUS],
        "write_off": status_counts[WRITE_OFF_STATUS],
        "unidentified": status_counts[UNIDENTIFIED_STATUS],
        "amount_mismatch": amount_mismatch,
        "blank_key": blank_key,
        "match_rate": matched / total if total else 0,
        "not_found_rate": not_found / total if total else 0,
        "source_amount": cents_to_decimal(source_amount_cents),
        "matched_amount": cents_to_decimal(matched_amount_cents),
        "matched_exact_amount": cents_to_decimal(matched_exact_cents),
        "amount_mismatch_amount": cents_to_decimal(amount_mismatch_cents),
        "not_found_amount": cents_to_decimal(not_found_cents),
        "mambu_amount": cents_to_decimal(mambu_cents),
        "write_off_amount": cents_to_decimal(write_off_cents),
        "unidentified_amount": cents_to_decimal(unidentified_cents),
    }


def summarize_itc_compare(rows: list[OrderedDict[str, Any]]) -> dict[str, Any]:
    total = len(rows)
    status_counts = Counter(row["Match_Status"] for row in rows)
    status_amount_cents: defaultdict[str, int] = defaultdict(int)
    source_amount_cents = 0
    matched_amount_cents = 0

    for row in rows:
        cents = amount_to_cents(row["Source_Amount"]) or 0
        status = row["Match_Status"]
        source_amount_cents += cents
        status_amount_cents[status] += cents
        if status != ITC_NOT_FOUND_STATUS:
            matched_amount_cents += cents

    matched = total - status_counts[ITC_NOT_FOUND_STATUS]
    return {
        "total": total,
        "matched": matched,
        "not_found": status_counts[ITC_NOT_FOUND_STATUS],
        "match_rate": matched / total if total else 0,
        "source_amount": cents_to_decimal(source_amount_cents),
        "matched_amount": cents_to_decimal(matched_amount_cents),
        "not_found_amount": cents_to_decimal(status_amount_cents[ITC_NOT_FOUND_STATUS]),
        "status_counts": dict(status_counts),
        "status_amounts": {
            status: cents_to_decimal(cents)
            for status, cents in status_amount_cents.items()
        },
    }


def summarize_itc_narration_source_breakdowns(
    itc_records: list[Record],
    itc_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> dict[str, dict[str, dict[str, Any]]]:
    groups = ("successful", "not_found")
    keys = ("_1", "_2", "_3", "_4", "genpay", "with_narration")
    raw: dict[str, dict[str, dict[str, int]]] = {
        group: {key: {"count": 0, "amount_cents": 0} for key in keys}
        for group in groups
    }

    for record, comparison in zip(itc_records, itc_vs_mambu_rows):
        status = comparison.get("Match_Status", "")
        group = "not_found" if status == ITC_NOT_FOUND_STATUS else "successful"
        amount_cents = record.amount_cents or amount_to_cents(comparison.get("Source_Amount")) or 0
        narration = str(first_value(record, "narration", "Narration", "ITC Narration")).strip()
        source = str(first_value(record, "source", "Source", "ITC Source")).strip()

        for suffix in ("_1", "_2", "_3", "_4"):
            if narration.endswith(suffix):
                raw[group][suffix]["count"] += 1
                raw[group][suffix]["amount_cents"] += amount_cents

        if "genpay" in source.casefold():
            raw[group]["genpay"]["count"] += 1
            raw[group]["genpay"]["amount_cents"] += amount_cents

        if narration:
            raw[group]["with_narration"]["count"] += 1
            raw[group]["with_narration"]["amount_cents"] += amount_cents

    return {
        group: {
            key: {
                "count": values["count"],
                "amount": cents_to_decimal(values["amount_cents"]),
            }
            for key, values in group_values.items()
        }
        for group, group_values in raw.items()
    }


def summarize_record_amounts(
    records: list[Record],
    *amount_headers: str,
    column_index: int | None = None,
) -> dict[str, Any]:
    amount_cents = sum(
        amount_to_cents(
            first_value_or_column(record, column_index, *amount_headers)
            if column_index is not None
            else first_value(record, *amount_headers)
        ) or 0
        for record in records
    )
    return {
        "count": len(records),
        "amount": cents_to_decimal(amount_cents),
    }


def summarize_vodafone_charges(vodafone_records: list[Record]) -> dict[str, Any]:
    charge_count = 0
    charge_amount_cents = 0
    charge_detail = VODAFONE_CHARGE_DETAIL.casefold()

    for record in vodafone_records:
        details = " ".join(str(first_value(record, "Details")).casefold().split())
        if charge_detail not in details:
            continue
        amount_cents = amount_to_cents(first_value(record, "Withdrawn")) or 0
        if amount_cents == 0:
            continue
        charge_count += 1
        charge_amount_cents += abs(amount_cents)

    return {
        "count": charge_count,
        "amount": cents_to_decimal(charge_amount_cents),
    }


def summarize_itc_fee_breakdowns(itc_records: list[Record]) -> dict[str, dict[str, Any]]:
    groups: dict[str, dict[str, int]] = {
        "upsales_transaction_fees": {"count": 0, "amount_cents": 0},
        "commission_charge_itc_payment": {"count": 0, "amount_cents": 0},
    }

    for record in itc_records:
        narration = str(first_value_or_column(record, 13, "narration", "Narration", "ITC Narration")).strip()
        key = (
            "upsales_transaction_fees"
            if narration.endswith("_3")
            else "commission_charge_itc_payment"
        )
        groups[key]["count"] += 1
        groups[key]["amount_cents"] += amount_to_cents(first_value_or_column(record, 17, "fees", "Fees")) or 0

    return {
        key: {
            "count": values["count"],
            "amount": cents_to_decimal(values["amount_cents"]),
        }
        for key, values in groups.items()
    }


def sum_amounts(records: list[Record]) -> Decimal:
    return cents_to_decimal(sum(record.amount_cents or 0 for record in records))  # type: ignore[return-value]


def summarize_key_presence(source_records: list[Record], counterparty_records: list[Record]) -> dict[str, Any]:
    counterparty_index = build_index(counterparty_records)
    matched_count = 0
    unmatched_count = 0
    matched_amount_cents = 0
    unmatched_amount_cents = 0

    for record in source_records:
        is_matched = bool(record.key and counterparty_index.get((record.direction, record.key)))
        amount_cents = record.amount_cents or 0
        if is_matched:
            matched_count += 1
            matched_amount_cents += amount_cents
        else:
            unmatched_count += 1
            unmatched_amount_cents += amount_cents

    return {
        "matched_count": matched_count,
        "matched_amount": cents_to_decimal(matched_amount_cents),
        "unmatched_count": unmatched_count,
        "unmatched_amount": cents_to_decimal(unmatched_amount_cents),
    }


STYLE_DEFAULT = 0
STYLE_TITLE = 1
STYLE_SUBTITLE = 2
STYLE_SECTION = 3
STYLE_HEADER = 4
STYLE_LABEL = 5
STYLE_VALUE = 6
STYLE_MONEY = 7
STYLE_PERCENT = 8
STYLE_MATCHED = 9
STYLE_NOT_FOUND = 10
STYLE_NOTE = 11
STYLE_WARNING = 12
STYLE_WRITE_OFF = 13
STYLE_UNIDENTIFIED = 14
STYLE_TABLE_SECTION = 15
STYLE_TABLE_COUNT = 16
STYLE_TABLE_MONEY = 17
STYLE_ITC_WRITE_OFF_BROWN = 18
STYLE_WALLET_COMPANY = 19
STYLE_WALLET_TITLE = 20
STYLE_WALLET_HEADER = 21
STYLE_WALLET_SECTION = 22
STYLE_WALLET_BALANCE_LABEL = 23
STYLE_WALLET_BALANCE_MONEY = 24
STYLE_WALLET_MONEY = 25
STYLE_WALLET_FINAL_LABEL = 26
STYLE_WALLET_FINAL_MONEY = 27
STYLE_WALLET_SIGNATURE = 28
STYLE_GRAN_WORKFLOW = 29
STYLE_GRAN_SECTION = 30
STYLE_GRAN_SECTION_COL = 31
STYLE_GRAN_TOTAL = 32
STYLE_GRAN_TOTAL_NUM = 33
STYLE_GRAN_TOTAL_MONEY = 34
STYLE_GRAN_COUNT = 35
STYLE_GRAN_DATA = 36
STYLE_MASTER_TITLE = 37
STYLE_MASTER_HEADER = 38
STYLE_MASTER_SECTION = 39
STYLE_MASTER_SECTION_COUNT = 40
STYLE_MASTER_SECTION_MONEY = 41
STYLE_MASTER_BODY = 42
STYLE_MASTER_COUNT = 43
STYLE_MASTER_MONEY = 44
STYLE_MASTER_BODY_BOLD = 45

# Dashboard styles used by generated Master and Monthly summary workbooks.
STYLE_DASH_CANVAS = 46
STYLE_DASH_LOGO = 47
STYLE_DASH_EYEBROW = 48
STYLE_DASH_TITLE = 49
STYLE_DASH_META_LABEL = 50
STYLE_DASH_META_VALUE = 51
STYLE_DASH_STATUS = 52
STYLE_DASH_CARD_LABEL = 53
STYLE_DASH_CARD_VALUE = 54
STYLE_DASH_CARD_VALUE_ALERT = 55
STYLE_DASH_CARD_DETAIL = 56
STYLE_DASH_CARD_TOP_PRIMARY = 57
STYLE_DASH_CARD_TOP_DARK = 58
STYLE_DASH_PANEL = 59
STYLE_DASH_SECTION_LABEL = 60
STYLE_DASH_PANEL_TITLE = 61
STYLE_DASH_TABLE_HEADER = 62
STYLE_DASH_TABLE_LABEL = 63
STYLE_DASH_TABLE_COUNT = 64
STYLE_DASH_TABLE_MONEY = 65
STYLE_DASH_TABLE_TOTAL_LABEL = 66
STYLE_DASH_TABLE_TOTAL_COUNT = 67
STYLE_DASH_TABLE_TOTAL_MONEY = 68
STYLE_DASH_VARIANCE_LABEL = 69
STYLE_DASH_VARIANCE_COUNT = 70
STYLE_DASH_VARIANCE_MONEY = 71
STYLE_DASH_MUTED = 72
STYLE_DASH_STATUS_READY = 73
STYLE_DASH_STATUS_PENDING = 74
STYLE_DASH_BAR_PRIMARY = 75
STYLE_DASH_BAR_TRACK = 76
STYLE_DASH_FOOTER = 77
STYLE_DASH_CARD_PERCENT = 78
STYLE_DASH_CARD_COUNT = 79
STYLE_DASH_TABLE_PERCENT = 80
STYLE_DASH_CARD_LABEL_LARGE = 81
STYLE_DASH_CARD_VALUE_FULL = 82
STYLE_DASH_CARD_VALUE_ALERT_FULL = 83
STYLE_DASH_CARD_DETAIL_LARGE = 84
STYLE_DASH_TABLE_HEADER_LARGE = 85
STYLE_DASH_TABLE_LABEL_LARGE = 86
STYLE_DASH_TABLE_COUNT_LARGE = 87
STYLE_DASH_TABLE_MONEY_LARGE = 88
STYLE_DASH_TABLE_TOTAL_LABEL_LARGE = 89
STYLE_DASH_TABLE_TOTAL_COUNT_LARGE = 90
STYLE_DASH_TABLE_TOTAL_MONEY_LARGE = 91
STYLE_WALLET_LINE_ITEM = 92


def _normalize_theme_hex(value: Any, fallback: str) -> str:
    text = str(value or "").strip().lstrip("#")
    return text.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", text) else fallback


@lru_cache(maxsize=1)
def dashboard_theme_colors() -> dict[str, str]:
    """Load dashboard colors from .streamlit/config.toml with safe defaults."""
    defaults = {
        "primary": "D6086B",
        "background": "F7FAF9",
        "secondary": "F7FAF9",
        "text": "212726",
    }
    config_path = Path(__file__).resolve().parent / ".streamlit" / "config.toml"
    try:
        config = tomllib.loads(config_path.read_text(encoding="utf-8"))
        theme = config.get("theme", {})
    except (OSError, tomllib.TOMLDecodeError):
        theme = {}
    colors = {
        "primary": _normalize_theme_hex(theme.get("primaryColor"), defaults["primary"]),
        "background": _normalize_theme_hex(theme.get("backgroundColor"), defaults["background"]),
        "secondary": _normalize_theme_hex(theme.get("secondaryBackgroundColor"), defaults["secondary"]),
        "text": _normalize_theme_hex(theme.get("textColor"), defaults["text"]),
    }

    def blend(foreground: str, background: str, background_weight: float) -> str:
        fg = tuple(int(foreground[i:i + 2], 16) for i in (0, 2, 4))
        bg = tuple(int(background[i:i + 2], 16) for i in (0, 2, 4))
        rgb = tuple(round(f * (1 - background_weight) + b * background_weight) for f, b in zip(fg, bg))
        return "".join(f"{component:02X}" for component in rgb)

    colors["primary_dark"] = blend(colors["primary"], colors["text"], 0.25)
    colors["primary_soft"] = blend(colors["primary"], colors["background"], 0.88)
    colors["muted"] = blend(colors["text"], colors["background"], 0.45)
    colors["line"] = blend(colors["text"], colors["background"], 0.87)
    colors["track"] = blend(colors["text"], colors["background"], 0.93)
    colors["pale"] = blend(colors["primary"], colors["background"], 0.95)
    return colors


def dashboard_grid(row_count: int, col_count: int) -> list[list[Cell]]:
    return [
        [Cell("", STYLE_DASH_CANVAS) for _ in range(col_count)]
        for _ in range(row_count)
    ]


def dashboard_style_range(
    rows: list[list[Any]],
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    style: int,
) -> None:
    for row_number in range(start_row, end_row + 1):
        for col_number in range(start_col, end_col + 1):
            rows[row_number - 1][col_number - 1] = Cell("", style)


def dashboard_merge(
    rows: list[list[Any]],
    merges: list[str],
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    value: Any,
    style: int,
) -> None:
    dashboard_style_range(rows, start_row, start_col, end_row, end_col, style)
    rows[start_row - 1][start_col - 1] = Cell(value, style)
    merges.append(
        f"{column_letter(start_col)}{start_row}:{column_letter(end_col)}{end_row}"
    )


def dashboard_cell(rows: list[list[Any]], row: int, col: int, value: Any, style: int) -> None:
    rows[row - 1][col - 1] = Cell(value, style)


def styles_xml() -> str:
    theme = dashboard_theme_colors()
    primary = f"FF{theme['primary']}"
    primary_dark = f"FF{theme['primary_dark']}"
    primary_soft = f"FF{theme['primary_soft']}"
    background = f"FF{theme['background']}"
    text = f"FF{theme['text']}"
    muted = f"FF{theme['muted']}"
    line = f"FF{theme['line']}"
    track = f"FF{theme['track']}"
    pale = f"FF{theme['pale']}"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="6">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
    <numFmt numFmtId="166" formatCode="#,##0"/>
    <numFmt numFmtId="167" formatCode="&quot;GH¢&quot;#,##0.00;[Red]-&quot;GH¢&quot;#,##0.00;&quot;GH¢&quot;0.00"/>
    <numFmt numFmtId="168" formatCode="[&gt;=1000000]&quot;GH¢&quot;0.00,,&quot;M&quot;;[&gt;=1000]&quot;GH¢&quot;0.0,&quot;K&quot;;&quot;GH¢&quot;0.00"/>
    <numFmt numFmtId="169" formatCode="+#,##0;-#,##0;0"/>
  </numFmts>
  <fonts count="40">
    <font><sz val="11"/><color rgb="{text}"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="10"/><color rgb="{muted}"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="{text}"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="{text}"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><color rgb="{text}"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="{text}"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="{primary}"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="28"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="{primary}"/><name val="Aptos"/></font>
    <font><sz val="12"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="24"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="24"/><color rgb="{primary_dark}"/><name val="Aptos"/></font>
    <font><b/><sz val="10"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><sz val="10"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="17"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="10"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><sz val="12"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="12"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="28"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
    <font><b/><sz val="12"/><color rgb="{primary_dark}"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="{primary}"/><name val="Aptos"/></font>
    <font><b/><sz val="10"/><color rgb="{primary_dark}"/><name val="Aptos"/></font>
    <font><b/><sz val="10"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="12"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><sz val="10"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><sz val="10"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="12"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="28"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><b/><sz val="28"/><color rgb="{primary_dark}"/><name val="Aptos"/></font>
    <font><sz val="12"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><sz val="13"/><color rgb="{text}"/><name val="Aptos"/></font>
    <font><sz val="13"/><color rgb="{muted}"/><name val="Aptos"/></font>
    <font><b/><sz val="13"/><color rgb="{text}"/><name val="Aptos"/></font>
  </fonts>
  <fills count="26">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{text}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{background}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_soft}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{pale}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{text}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_dark}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{background}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_soft}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{text}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_soft}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{background}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_soft}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{text}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{track}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{pale}"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{primary_dark}"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="5">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2EC"/></left><right style="thin"><color rgb="FFD9E2EC"/></right><top style="thin"><color rgb="FFD9E2EC"/></top><bottom style="thin"><color rgb="FFD9E2EC"/></bottom><diagonal/></border>
    <border><left style="thin"><color rgb="{line}"/></left><right style="thin"><color rgb="{line}"/></right><top style="thin"><color rgb="{line}"/></top><bottom style="thin"><color rgb="{line}"/></bottom><diagonal/></border>
    <border><left/><right/><top/><bottom style="thin"><color rgb="{line}"/></bottom><diagonal/></border>
    <border><left/><right/><top style="thin"><color rgb="{line}"/></top><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="93">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="4" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="3" fillId="10" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="3" fillId="10" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="8" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="8" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="8" fillId="13" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="9" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="9" fillId="14" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="9" fillId="14" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="0" fillId="12" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="9" fillId="14" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="10" fillId="14" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="9" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="11" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="12" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="3" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="3" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="11" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="15" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="3" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="3" fillId="15" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="0" fillId="17" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="0" fillId="17" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="0" fillId="17" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="4" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="18" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="24" fillId="20" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="14" fillId="18" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="13" fillId="18" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="18" fillId="18" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="23" fillId="18" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="25" fillId="21" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="18" fillId="19" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="168" fontId="16" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="167" fontId="17" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="19" fillId="19" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="20" borderId="2" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="22" borderId="2" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="19" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="26" fillId="19" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="20" fillId="19" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="21" fillId="19" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="22" fillId="19" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="15" fillId="19" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="167" fontId="29" fillId="19" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="23" fillId="19" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="23" fillId="19" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="167" fontId="23" fillId="19" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="26" fillId="24" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="169" fontId="16" fillId="24" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="167" fontId="17" fillId="24" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="19" fillId="19" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="27" fillId="21" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="28" fillId="18" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="20" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="23" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="30" fillId="18" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="165" fontId="16" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="166" fontId="16" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="165" fontId="29" fillId="19" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="32" fillId="19" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="167" fontId="33" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="167" fontId="34" fillId="19" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="35" fillId="19" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="36" fillId="19" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="37" fillId="19" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="38" fillId="19" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="167" fontId="39" fillId="19" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="39" fillId="19" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="39" fillId="19" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="167" fontId="39" fillId="19" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="0" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


def content_types_xml(sheet_count: int) -> str:
    sheet_overrides = "\n".join(
        f'  <Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{sheet_overrides}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def workbook_xml(sheet_names: list[str]) -> str:
    sheet_xml = "\n".join(
        f'    <sheet name="{xml_text(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="false"/>
  <sheets>
{sheet_xml}
  </sheets>
  <calcPr calcId="191029" calcMode="auto"/>
</workbook>
"""


def workbook_rels_xml(sheet_count: int) -> str:
    relationships = [
        f'  <Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1)
    ]
    relationships.append(
        f'  <Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{os.linesep.join(relationships)}
</Relationships>
"""


def core_xml() -> str:
    created = (
        dt.datetime.now(dt.timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Finance reconciliation generator</dc:creator>
  <cp:lastModifiedBy>Finance reconciliation generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>
"""


def app_xml(sheet_names: list[str]) -> str:
    titles = "\n".join(f"        <vt:lpstr>{xml_text(name)}</vt:lpstr>" for name in sheet_names)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="{len(sheet_names)}" baseType="lpstr">
{titles}
    </vt:vector>
  </TitlesOfParts>
</Properties>
"""


def worksheet_cell_xml(row_number: int, col_number: int, cell: Any, style: int | None = None) -> str:
    if isinstance(cell, Cell):
        value = cell.value
        style = cell.style
    else:
        value = cell

    ref = f"{column_letter(col_number)}{row_number}"
    style_attr = f' s="{style}"' if style is not None else ""

    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>' if style is not None else ""

    if isinstance(value, bool):
        return f'<c r="{ref}"{style_attr} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, int):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    if isinstance(value, float):
        return f'<c r="{ref}"{style_attr}><v>{value:.12g}</v></c>'
    if isinstance(value, Decimal):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'

    return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'


def row_xml(
    row_number: int,
    row: list[Any],
    style_func: Callable[[int, int, Any], int | None] | None = None,
    height: float | None = None,
) -> str:
    cells: list[str] = []
    for col_number, value in enumerate(row, start=1):
        style = style_func(row_number, col_number, value) if style_func else None
        cell = worksheet_cell_xml(row_number, col_number, value, style)
        if cell:
            cells.append(cell)
    height_attr = f' ht="{height:g}" customHeight="1"' if height is not None else ""
    return f'<row r="{row_number}"{height_attr}>{"".join(cells)}</row>\n'


def widths_xml(widths: list[int | float]) -> str:
    if not widths:
        return ""
    cols = []
    for idx, width in enumerate(widths, start=1):
        width = max(2.0, min(45.0, float(width)))
        cols.append(f'<col min="{idx}" max="{idx}" width="{width:g}" customWidth="1"/>')
    return f"<cols>{''.join(cols)}</cols>"


def write_worksheet(
    zf: zipfile.ZipFile,
    path: str,
    rows: Iterable[list[Any]],
    row_count: int,
    col_count: int,
    widths: list[int | float],
    *,
    freeze_top_row: bool = True,
    autofilter: bool = True,
    merges: list[str] | None = None,
    style_func: Callable[[int, int, Any], int | None] | None = None,
    row_heights: Mapping[int, float] | None = None,
    show_gridlines: bool = True,
    zoom_scale: int | None = None,
    page_orientation: str | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> None:
    dimension = f"A1:{column_letter(max(1, col_count))}{max(1, row_count)}"
    view_attrs = ['workbookViewId="0"']
    if not show_gridlines:
        view_attrs.append('showGridLines="0"')
    if zoom_scale is not None:
        view_attrs.append(f'zoomScale="{max(10, min(400, int(zoom_scale)))}"')
    view_attr_text = " ".join(view_attrs)
    sheet_views = (
        f'<sheetViews><sheetView {view_attr_text}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft"/></sheetView></sheetViews>'
        if freeze_top_row
        else f'<sheetViews><sheetView {view_attr_text}/></sheetViews>'
    )
    merge_xml = ""
    if merges:
        merge_xml = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{ref}"/>' for ref in merges) + "</mergeCells>"
    autofilter_xml = f'<autoFilter ref="{dimension}"/>' if autofilter and row_count >= 1 and col_count >= 1 else ""

    sheet_pr_xml = ""
    page_setup_xml = ""
    if page_orientation or fit_to_width is not None or fit_to_height is not None:
        orientation = page_orientation or "portrait"
        width = 1 if fit_to_width is None else int(fit_to_width)
        height = 0 if fit_to_height is None else int(fit_to_height)
        sheet_pr_xml = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        page_setup_xml = f'<pageSetup orientation="{xml_text(orientation)}" fitToWidth="{width}" fitToHeight="{height}"/>'

    with zf.open(path, "w") as handle:
        handle.write(
            (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f"{sheet_pr_xml}"
                f'<dimension ref="{dimension}"/>'
                f"{sheet_views}"
                '<sheetFormatPr defaultRowHeight="15"/>'
                f"{widths_xml(widths)}"
                "<sheetData>\n"
            ).encode("utf-8")
        )
        batch: list[str] = []
        for row_number, row in enumerate(rows, start=1):
            batch.append(row_xml(row_number, row, style_func, (row_heights or {}).get(row_number)))
            if len(batch) >= XLSX_ROW_BATCH_SIZE:
                handle.write("".join(batch).encode("utf-8"))
                batch.clear()
        if batch:
            handle.write("".join(batch).encode("utf-8"))
        handle.write(
            (
                "</sheetData>"
                f"{autofilter_xml}"
                f"{merge_xml}"
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                f"{page_setup_xml}"
                "</worksheet>"
            ).encode("utf-8")
        )


def compute_widths(headers: list[str], data_rows: Iterable[Iterable[Any]], max_rows: int = WIDTH_SAMPLE_ROWS) -> list[int]:
    widths = [min(45, max(10, len(str(header)) + 2)) for header in headers]
    for row_idx, row in enumerate(data_rows):
        if row_idx >= max_rows:
            break
        for idx, value in enumerate(row):
            if idx >= len(widths):
                continue
            widths[idx] = min(45, max(widths[idx], len(str(value)) + 2))
    return widths


def record_rows(headers: list[str], records: list[Record]) -> Iterable[list[Any]]:
    yield headers
    for record in records:
        yield [record.sheet_values.get(header, "") for header in headers]


def first_value(record: Record, *headers: str) -> Any:
    for header in headers:
        value = record.sheet_values.get(header, "")
        if value not in ("", None):
            return value
    return ""


NOT_FOUND_DETAIL_HEADERS = [
    "Wallet",
    "Recon Side",
    "Transaction ID",
    "Raw Reconciliation ID",
    "Reconciliation Key",
    "Date",
    "Amount (GHC)",
    "Account",
    "Phone",
    "Source File",
    "Source Row",
    "Match Reason",
]


def not_found_detail_rows(
    records: list[Record],
    comparison_rows: list[OrderedDict[str, Any]],
    not_found_status: str,
    wallet: str,
    recon_side: str,
) -> list[OrderedDict[str, Any]]:
    """Return a compact, investigation-ready view of Not-Found transactions."""
    details: list[OrderedDict[str, Any]] = []
    transaction_headers = (
        "Transaction_ID",
        "Trans ID",
        "processor_transaction_id",
        "Id",
        "ID",
        "External Id",
        "External id",
        "Receipt No.",
        "Description",
        "Identifier",
        "Identifier (KEY)",
        "Identifier (Key)",
    )
    reconciliation_headers = (
        "External_Debit_Reference",
        "Ext Debit Ref (Key)",
        "thirdparty_id (KEY)",
        "thirdparty_id",
        "Identifier",
        "Identifier (KEY)",
        "Identifier (Key)",
        "Id",
        "ID",
        "Description",
    )

    for record, comparison in zip(records, comparison_rows):
        status = comparison.get("Match_Status", comparison.get("Match Status", ""))
        if str(status) != not_found_status:
            continue

        transaction_id = first_value(record, *transaction_headers) or record.transaction_id
        raw_reconciliation_id = first_value(record, *reconciliation_headers)
        if not transaction_id:
            transaction_id = raw_reconciliation_id or record.key
        if not raw_reconciliation_id:
            raw_reconciliation_id = record.key

        account = record.account or first_value(
            record,
            "Account ID",
            "Account Number",
            "Loan Number",
            "SendingHse_ID",
            "ReceivingHse_ID",
        )
        phone = record.phone or first_value(
            record,
            "Mobile Phone (Client)",
            "SendingHse_Account",
            "ReceivingHse_Account",
        )
        reason = comparison.get("Match_Reason", comparison.get("Match Reason", ""))

        detail: OrderedDict[str, Any] = OrderedDict()
        detail["Wallet"] = wallet
        detail["Recon Side"] = recon_side
        detail["Transaction ID"] = transaction_id
        detail["Raw Reconciliation ID"] = raw_reconciliation_id
        detail["Reconciliation Key"] = record.key
        detail["Date"] = record.datetime_value
        detail["Amount (GHC)"] = record.amount_decimal
        detail["Account"] = account
        detail["Phone"] = phone
        detail["Source File"] = record.source_file
        detail["Source Row"] = record.source_row
        detail["Match Reason"] = reason or f"{not_found_status} during {recon_side} reconciliation"
        details.append(detail)
    return details


def first_value_or_column(record: Record, column_index: int, *headers: str) -> Any:
    value = first_value(record, *headers)
    if value not in ("", None):
        return value

    values = list(record.sheet_values.values())
    if len(values) >= column_index:
        return values[column_index - 1]
    return ""


def write_off_sheet_row(record: Record) -> list[Any]:
    return [
        first_value(record, "Date", "Value Date (Entry Date)", "Reconciliation_DateTime"),
        first_value(record, "Channel", "Type"),
        first_value(record, "Amount", "Reconciliation_Amount"),
        first_value(record, "Fido Client Name", "Account Holder Name"),
        first_value(record, "Account Number", "Account ID"),
        first_value(record, "Transaction ID", "Transactions Id", "Transactions ID", "Identifier", "Reconciliation_Key"),
        first_value(record, "Loan ID", "Loan Number"),
        first_value(record, "Staus", "Status"),
    ]


def unidentified_sheet_row(record: Record) -> list[Any]:
    return [
        first_value(record, "Value Date (Entry Date)", "Date", "Reconciliation_DateTime"),
        first_value(record, "Identifier", "Transaction ID", "Reconciliation_Key"),
        first_value(record, "Amount", "Reconciliation_Amount"),
        first_value(record, "Type", "Channel", "Source_Dataset"),
        first_value(record, "User"),
        first_value(record, "Account ID", "Account Number", "Loan ID"),
        first_value(record, "Notes"),
        first_value(record, "Product (Deposit)", "Product"),
    ]


def category_rows(headers: list[str], records: list[Record], row_builder: Callable[[Record], list[Any]]) -> Iterable[list[Any]]:
    yield headers
    for record in records:
        yield row_builder(record)


def compare_sheet_rows(rows: list[OrderedDict[str, Any]]) -> Iterable[list[Any]]:
    yield COMPARE_HEADERS
    for row in rows:
        yield [row.get(header, "") for header in COMPARE_HEADERS]


def rows_from_dicts(headers: list[str], rows: list[OrderedDict[str, Any]]) -> Iterable[list[Any]]:
    yield headers
    for row in rows:
        yield [row.get(header, "") for header in headers]


def exact_compare_rows(
    source_headers: list[str],
    source_records: list[Record],
    comparison_rows: list[OrderedDict[str, Any]],
) -> Iterable[list[Any]]:
    yield source_headers + COMPARE_RESULT_HEADERS
    for record, comparison in zip(source_records, comparison_rows):
        yield [record.sheet_values.get(header, "") for header in source_headers] + [
            comparison.get(header, "") for header in COMPARE_RESULT_HEADERS
        ]


def data_style(headers: list[str]) -> Callable[[int, int, Any], int | None]:
    amount_columns = {
        "Amount",
        "Amount (GHC)",
        "amount",
        "net_amount",
        "fees",
        "elevy_charge",
        "Paid In",
        "Withdrawn",
        "Balance",
        "Principal Amount",
        "Amount_GHC",
        "Charge",
        "balanceBefore",
        "balanceAfter",
        "Reconciliation_Amount",
        "Source_Amount",
        "Counterparty_Amount",
        "Amount_Difference",
    }

    def style(row_number: int, col_number: int, value: Any) -> int | None:
        if row_number == 1:
            return STYLE_HEADER
        header = headers[col_number - 1] if col_number - 1 < len(headers) else ""
        if header == "Match_Status":
            if value in {MATCHED, MAMBU_STATUS}:
                return STYLE_MATCHED
            if value == WRITE_OFF_STATUS:
                return STYLE_WRITE_OFF
            if value == UNIDENTIFIED_STATUS:
                return STYLE_UNIDENTIFIED
            return STYLE_NOT_FOUND
        if header in amount_columns:
            return STYLE_MONEY
        return None

    return style


def itc_data_style(
    headers: list[str],
    *,
    unidentified_style: int = STYLE_WRITE_OFF,
    write_off_status: str = ITC_WRITE_OFF_STATUS,
    write_off_style: int = STYLE_ITC_WRITE_OFF_BROWN,
) -> Callable[[int, int, Any], int | None]:
    base_style = data_style(headers)

    def style(row_number: int, col_number: int, value: Any) -> int | None:
        if row_number == 1:
            return STYLE_HEADER
        header = headers[col_number - 1] if col_number - 1 < len(headers) else ""
        if header == "Match_Status":
            if value in {ITC_STATUS, ITC_MAMBU_STATUS}:
                return STYLE_MATCHED
            if value == VODA_COLL_STATUS:
                return STYLE_UNIDENTIFIED
            if value == ITC_UPSALE_STATUS:
                return STYLE_UNIDENTIFIED
            if value == ITC_UNIDENTIFIED_STATUS:
                return unidentified_style
            if value == write_off_status:
                return write_off_style
            return STYLE_NOT_FOUND
        return base_style(row_number, col_number, value)

    return style


def build_summary_rows(
    mambu_records: list[Record],
    nsano_records: list[Record],
    write_off_records: list[Record],
    unidentified_records: list[Record],
    mambu_raw_count: int,
    mambu_vs_nsano_stats: dict[str, Any],
    nsano_vs_mambu_stats: dict[str, Any],
    output_name: str,
) -> list[list[Any]]:
    mvn = mambu_vs_nsano_stats
    nvm = nsano_vs_mambu_stats

    mambu_writeoff_count = mvn["matched"] + nvm["write_off"]
    mambu_writeoff_amount = mvn["matched_amount"] + nvm["write_off_amount"]
    not_in_nsano_count = mvn["not_found"]
    not_in_nsano_amount = mvn["not_found_amount"]
    t3_total_count = mambu_writeoff_count + nvm["total"] + not_in_nsano_count
    t3_total_amount = mambu_writeoff_amount + nvm["source_amount"] + not_in_nsano_amount
    nsano_charge = summarize_record_amounts(nsano_records, "Charge", column_index=21)

    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(val: Any) -> Cell:
        return Cell(val, STYLE_TABLE_COUNT)

    def tm(val: Any) -> Cell:
        return Cell(val, STYLE_TABLE_MONEY)

    rows: list[list[Any]] = [
        # ── Title ────────────────────────────────────────────────────────────
        [Cell("NSANO RECONCILIATION — SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        # ── Table 1: Mambu vs NSANO ──────────────────────────────────────────
        [th("Mambu vs NSANO"), th("Count"), th("Amount (GHC)")],
        [Cell("✓ MATCHED", STYLE_MATCHED), mvn["matched_exact"], Cell(mvn["matched_exact_amount"], STYLE_MONEY)],
        ["⚠ AMT MISMATCH", mvn["amount_mismatch"], Cell(mvn["amount_mismatch_amount"], STYLE_MONEY)],
        ["⚠ NOT IN NSANO", mvn["not_found"], Cell(mvn["not_found_amount"], STYLE_MONEY)],
        [th("TOTAL"), tc(mvn["total"]), tm(mvn["source_amount"])],
        ["", "", ""],
        # ── Table 2: NSANO vs Source ─────────────────────────────────────────
        [th("NSANO vs Source"), th("Count"), th("Amount (GHC)")],
        ["Mambu", nvm["mambu"], Cell(nvm["mambu_amount"], STYLE_MONEY)],
        ["Unidentified", nvm["unidentified"], Cell(nvm["unidentified_amount"], STYLE_MONEY)],
        ["Write Off", nvm["write_off"], Cell(nvm["write_off_amount"], STYLE_MONEY)],
        ["Not Found", nvm["not_found"], Cell(nvm["not_found_amount"], STYLE_MONEY)],
        [th("TOTAL"), tc(nvm["total"]), tm(nvm["source_amount"])],
        ["", "", ""],
        # ── Second title ─────────────────────────────────────────────────────
        [Cell("NSANO RECONCILIATION — SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        # ── Table 3: Final comparison ────────────────────────────────────────
        [th("Mambu vs NSANO"), th("Count"), th("Amount (GHC)")],
        ["Mambu +Writeoff", mambu_writeoff_count, Cell(mambu_writeoff_amount, STYLE_MONEY)],
        ["Nsano", nvm["total"], Cell(nvm["source_amount"], STYLE_MONEY)],
        ["NOT IN NSANO(Difference)", not_in_nsano_count, Cell(not_in_nsano_amount, STYLE_MONEY)],
        [th("TOTAL"), tc(t3_total_count), tm(t3_total_amount)],
        ["", "", ""],
        [th("Nsano Charges"), th("Count"), th("Amount (GHC)")],
        ["Nsano Charge", nsano_charge["count"], Cell(nsano_charge["amount"], STYLE_MONEY)],
    ]
    return rows


def build_itc_summary_rows(
    mambu_vs_itc_stats: dict[str, Any],
    itc_vs_mambu_stats: dict[str, Any],
    vodafone_vs_mambu_stats: dict[str, Any],
    itc_breakdowns: dict[str, dict[str, dict[str, Any]]],
    itc_fee_breakdowns: dict[str, dict[str, Any]],
    vodafone_charge_summary: dict[str, Any] | None = None,
) -> list[list[Any]]:
    mvi = mambu_vs_itc_stats
    ivm = itc_vs_mambu_stats
    vvm = vodafone_vs_mambu_stats

    def th(label: str) -> Cell:
        return Cell(label, STYLE_TABLE_SECTION)

    def tc(val: Any) -> Cell:
        return Cell(val, STYLE_TABLE_COUNT)

    def tm(val: Any) -> Cell:
        return Cell(val, STYLE_TABLE_MONEY)

    def count(stats: dict[str, Any], status: str) -> int:
        return int(stats["status_counts"].get(status, 0))

    def amount(stats: dict[str, Any], status: str) -> Decimal | str:
        return stats["status_amounts"].get(status, Decimal("0"))

    def breakdown(group: str, key: str) -> dict[str, Any]:
        return itc_breakdowns.get(group, {}).get(key, {"count": 0, "amount": Decimal("0")})

    def bcount(group: str, key: str) -> int:
        return int(breakdown(group, key)["count"])

    def bamount(group: str, key: str) -> Decimal | str:
        return breakdown(group, key)["amount"]

    def fee_breakdown(key: str) -> dict[str, Any]:
        return itc_fee_breakdowns.get(key, {"count": 0, "amount": Decimal("0")})

    def fcount(key: str) -> int:
        return int(fee_breakdown(key)["count"])

    def famount(key: str) -> Decimal | str:
        return fee_breakdown(key)["amount"]

    fees_total_count = fcount("upsales_transaction_fees") + fcount("commission_charge_itc_payment")
    fees_total_amount = famount("upsales_transaction_fees") + famount("commission_charge_itc_payment")
    vodafone_charges = vodafone_charge_summary or {"count": 0, "amount": Decimal("0")}

    return [
        [Cell("ITC RECONCILIATION — SUMMARY", STYLE_TITLE), "", ""],
        ["", "", ""],
        [th("Mambu Vs ITC"), th("Count"), th("Amount (GHC)")],
        [Cell(ITC_STATUS, STYLE_MATCHED), count(mvi, ITC_STATUS), Cell(amount(mvi, ITC_STATUS), STYLE_MONEY)],
        [Cell(VODA_COLL_STATUS, STYLE_UNIDENTIFIED), count(mvi, VODA_COLL_STATUS), Cell(amount(mvi, VODA_COLL_STATUS), STYLE_MONEY)],
        [Cell(ITC_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(mvi, ITC_NOT_FOUND_STATUS), Cell(amount(mvi, ITC_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(mvi["total"]), tm(mvi["source_amount"])],
        ["", "", ""],
        [th("ITC Vs Mambu"), th("Count"), th("Amount (GHC)")],
        [Cell(ITC_MAMBU_STATUS, STYLE_MATCHED), count(ivm, ITC_MAMBU_STATUS), Cell(amount(ivm, ITC_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(VODA_COLL_STATUS, STYLE_UNIDENTIFIED), count(ivm, VODA_COLL_STATUS), Cell(amount(ivm, VODA_COLL_STATUS), STYLE_MONEY)],
        [Cell(ITC_UNIDENTIFIED_STATUS, STYLE_WRITE_OFF), count(ivm, ITC_UNIDENTIFIED_STATUS), Cell(amount(ivm, ITC_UNIDENTIFIED_STATUS), STYLE_MONEY)],
        [Cell(ITC_WRITE_OFF_STATUS, STYLE_ITC_WRITE_OFF_BROWN), count(ivm, ITC_WRITE_OFF_STATUS), Cell(amount(ivm, ITC_WRITE_OFF_STATUS), STYLE_MONEY)],
        [Cell(ITC_UPSALE_STATUS, STYLE_UNIDENTIFIED), count(ivm, ITC_UPSALE_STATUS), Cell(amount(ivm, ITC_UPSALE_STATUS), STYLE_MONEY)],
        [Cell(ITC_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(ivm, ITC_NOT_FOUND_STATUS), Cell(amount(ivm, ITC_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(ivm["total"]), tm(ivm["source_amount"])],
        ["", "", ""],
        [th("ITC Transaction Fees"), th("Count"), th("Amount (GHC)")],
        ["Upsales Transaction fees", fcount("upsales_transaction_fees"), Cell(famount("upsales_transaction_fees"), STYLE_MONEY)],
        ["Commission Charge ITC Payment", fcount("commission_charge_itc_payment"), Cell(famount("commission_charge_itc_payment"), STYLE_MONEY)],
        [th("TOTAL"), tc(fees_total_count), tm(fees_total_amount)],
        ["", "", ""],
        [th("ITC Successful — Breakdown by Narration"), th("Count"), th("Amount (GHC)")],
        ["Successful — _1", bcount("successful", "_1"), Cell(bamount("successful", "_1"), STYLE_MONEY)],
        ["Successful — _2", bcount("successful", "_2"), Cell(bamount("successful", "_2"), STYLE_MONEY)],
        ["Successful — _3", bcount("successful", "_3"), Cell(bamount("successful", "_3"), STYLE_MONEY)],
        ["Successful — _4", bcount("successful", "_4"), Cell(bamount("successful", "_4"), STYLE_MONEY)],
        ["Genpay", bcount("successful", "genpay"), Cell(bamount("successful", "genpay"), STYLE_MONEY)],
        ["Total Successful", bcount("successful", "with_narration"), Cell(bamount("successful", "with_narration"), STYLE_MONEY)],
        ["", "", ""],
        [th("ITC Not Found — Breakdown by Narration"), th("Count"), th("Amount (GHC)")],
        ["Not Found — _1", bcount("not_found", "_1"), Cell(bamount("not_found", "_1"), STYLE_MONEY)],
        ["Not Found — _2", bcount("not_found", "_2"), Cell(bamount("not_found", "_2"), STYLE_MONEY)],
        ["Not Found — _3", bcount("not_found", "_3"), Cell(bamount("not_found", "_3"), STYLE_MONEY)],
        ["Not Found — _4", bcount("not_found", "_4"), Cell(bamount("not_found", "_4"), STYLE_MONEY)],
        ["Genpay", bcount("not_found", "genpay"), Cell(bamount("not_found", "genpay"), STYLE_MONEY)],
        ["Totla Not Found", bcount("not_found", "with_narration"), Cell(bamount("not_found", "with_narration"), STYLE_MONEY)],
        ["", "", ""],
        [th("Vodafone Collections Vs Mambu"), th("Count"), th("Amount (GHC)")],
        [Cell(ITC_MAMBU_STATUS, STYLE_MATCHED), count(vvm, ITC_MAMBU_STATUS), Cell(amount(vvm, ITC_MAMBU_STATUS), STYLE_MONEY)],
        [Cell(ITC_UNIDENTIFIED_STATUS, STYLE_UNIDENTIFIED), count(vvm, ITC_UNIDENTIFIED_STATUS), Cell(amount(vvm, ITC_UNIDENTIFIED_STATUS), STYLE_MONEY)],
        [Cell(VODAFONE_WRITE_OFF_STATUS, STYLE_WRITE_OFF), count(vvm, VODAFONE_WRITE_OFF_STATUS), Cell(amount(vvm, VODAFONE_WRITE_OFF_STATUS), STYLE_MONEY)],
        [Cell(ITC_NOT_FOUND_STATUS, STYLE_NOT_FOUND), count(vvm, ITC_NOT_FOUND_STATUS), Cell(amount(vvm, ITC_NOT_FOUND_STATUS), STYLE_MONEY)],
        [th("TOTAL"), tc(vvm["total"]), tm(vvm["source_amount"])],
        ["", "", ""],
        [th("Vodafone Charges"), th("Count"), th("Amount (GHC)")],
        [VODAFONE_CHARGE_DETAIL, int(vodafone_charges.get("count", 0)), Cell(vodafone_charges.get("amount", Decimal("0")), STYLE_MONEY)],
    ]


def summary_style(row_number: int, col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    if isinstance(value, Decimal):
        return STYLE_MONEY
    return None


def summary_merges(rows: list[list[Any]]) -> list[str]:
    merges: list[str] = []
    for row_number, row in enumerate(rows, start=1):
        first = row[0] if row else ""
        if isinstance(first, Cell) and first.style == STYLE_TITLE:
            merges.append(f"A{row_number}:C{row_number}")
    return merges


def write_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    nsano_headers: list[str],
    nsano_records: list[Record],
    write_off_records: list[Record],
    unidentified_records: list[Record],
    mambu_vs_nsano_rows: list[OrderedDict[str, Any]],
    nsano_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Summary",
        "Mambu",
        "Nsano",
        "Write Off",
        "Unidentified",
        "Mambu vs Nsano",
        "Nsano vs Mambu",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    nsano_preview = ([record.sheet_values.get(header, "") for header in nsano_headers] for record in nsano_records[:5000])
    write_off_preview = (write_off_sheet_row(record) for record in write_off_records[:5000])
    unidentified_preview = (unidentified_sheet_row(record) for record in unidentified_records[:5000])
    mambu_compare_headers = mambu_headers + COMPARE_RESULT_HEADERS
    nsano_compare_headers = nsano_headers + COMPARE_RESULT_HEADERS
    mambu_compare_preview = (
        [record.sheet_values.get(header, "") for header in mambu_headers]
        + [row.get(header, "") for header in COMPARE_RESULT_HEADERS]
        for record, row in zip(mambu_records[:5000], mambu_vs_nsano_rows[:5000])
    )
    nsano_compare_preview = (
        [record.sheet_values.get(header, "") for header in nsano_headers]
        + [row.get(header, "") for header in COMPARE_RESULT_HEADERS]
        for record, row in zip(nsano_records[:5000], nsano_vs_mambu_rows[:5000])
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
            "xl/worksheets/sheet2.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            record_rows(nsano_headers, nsano_records),
            len(nsano_records) + 1,
            len(nsano_headers),
            compute_widths(nsano_headers, nsano_preview),
            style_func=data_style(nsano_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            category_rows(WRITE_OFF_HEADERS, write_off_records, write_off_sheet_row),
            len(write_off_records) + 1,
            len(WRITE_OFF_HEADERS),
            compute_widths(WRITE_OFF_HEADERS, write_off_preview),
            style_func=data_style(WRITE_OFF_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            category_rows(UNIDENTIFIED_HEADERS, unidentified_records, unidentified_sheet_row),
            len(unidentified_records) + 1,
            len(UNIDENTIFIED_HEADERS),
            compute_widths(UNIDENTIFIED_HEADERS, unidentified_preview),
            style_func=data_style(UNIDENTIFIED_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            exact_compare_rows(mambu_headers, mambu_records, mambu_vs_nsano_rows),
            len(mambu_vs_nsano_rows) + 1,
            len(mambu_compare_headers),
            compute_widths(mambu_compare_headers, mambu_compare_preview),
            style_func=data_style(mambu_compare_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            exact_compare_rows(nsano_headers, nsano_records, nsano_vs_mambu_rows),
            len(nsano_vs_mambu_rows) + 1,
            len(nsano_compare_headers),
            compute_widths(nsano_compare_headers, nsano_compare_preview),
            style_func=data_style(nsano_compare_headers),
        )

    os.replace(temp_path, output_path)


def write_itc_workbook(
    output_path: Path,
    summary_rows: list[list[Any]],
    mambu_headers: list[str],
    mambu_records: list[Record],
    itc_headers: list[str],
    itc_records: list[Record],
    vodafone_headers: list[str],
    vodafone_records: list[Record],
    write_off_records: list[Record],
    unidentified_records: list[Record],
    mambu_vs_itc_rows: list[OrderedDict[str, Any]],
    itc_vs_mambu_rows: list[OrderedDict[str, Any]],
    vodafone_vs_mambu_rows: list[OrderedDict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    sheet_names = [
        "Summary",
        "Mambu",
        "ITC",
        "Vodafone Collections",
        "Unidentified",
        "Write Off",
        "Mambu Vs ITC",
        "ITC Vs Mambu",
        "Vodafone Collections Vs Mambu",
    ]

    mambu_preview = ([record.sheet_values.get(header, "") for header in mambu_headers] for record in mambu_records[:5000])
    itc_preview = ([record.sheet_values.get(header, "") for header in itc_headers] for record in itc_records[:5000])
    vodafone_preview = ([record.sheet_values.get(header, "") for header in vodafone_headers] for record in vodafone_records[:5000])
    write_off_preview = (write_off_sheet_row(record) for record in write_off_records[:5000])
    unidentified_preview = (unidentified_sheet_row(record) for record in unidentified_records[:5000])
    mambu_compare_headers = mambu_headers + COMPARE_RESULT_HEADERS
    itc_compare_headers = itc_headers + COMPARE_RESULT_HEADERS
    vodafone_compare_headers = vodafone_headers + COMPARE_RESULT_HEADERS
    mambu_compare_preview = (
        [record.sheet_values.get(header, "") for header in mambu_headers]
        + [row.get(header, "") for header in COMPARE_RESULT_HEADERS]
        for record, row in zip(mambu_records[:5000], mambu_vs_itc_rows[:5000])
    )
    itc_compare_preview = (
        [record.sheet_values.get(header, "") for header in itc_headers]
        + [row.get(header, "") for header in COMPARE_RESULT_HEADERS]
        for record, row in zip(itc_records[:5000], itc_vs_mambu_rows[:5000])
    )
    vodafone_compare_preview = (
        [record.sheet_values.get(header, "") for header in vodafone_headers]
        + [row.get(header, "") for header in COMPARE_RESULT_HEADERS]
        for record, row in zip(vodafone_records[:5000], vodafone_vs_mambu_rows[:5000])
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
            "xl/worksheets/sheet2.xml",
            record_rows(mambu_headers, mambu_records),
            len(mambu_records) + 1,
            len(mambu_headers),
            compute_widths(mambu_headers, mambu_preview),
            style_func=itc_data_style(mambu_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet3.xml",
            record_rows(itc_headers, itc_records),
            len(itc_records) + 1,
            len(itc_headers),
            compute_widths(itc_headers, itc_preview),
            style_func=itc_data_style(itc_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet4.xml",
            record_rows(vodafone_headers, vodafone_records),
            len(vodafone_records) + 1,
            len(vodafone_headers),
            compute_widths(vodafone_headers, vodafone_preview),
            style_func=itc_data_style(vodafone_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet5.xml",
            category_rows(UNIDENTIFIED_HEADERS, unidentified_records, unidentified_sheet_row),
            len(unidentified_records) + 1,
            len(UNIDENTIFIED_HEADERS),
            compute_widths(UNIDENTIFIED_HEADERS, unidentified_preview),
            style_func=itc_data_style(UNIDENTIFIED_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet6.xml",
            category_rows(WRITE_OFF_HEADERS, write_off_records, write_off_sheet_row),
            len(write_off_records) + 1,
            len(WRITE_OFF_HEADERS),
            compute_widths(WRITE_OFF_HEADERS, write_off_preview),
            style_func=itc_data_style(WRITE_OFF_HEADERS),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet7.xml",
            exact_compare_rows(mambu_headers, mambu_records, mambu_vs_itc_rows),
            len(mambu_vs_itc_rows) + 1,
            len(mambu_compare_headers),
            compute_widths(mambu_compare_headers, mambu_compare_preview),
            style_func=itc_data_style(mambu_compare_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet8.xml",
            exact_compare_rows(itc_headers, itc_records, itc_vs_mambu_rows),
            len(itc_vs_mambu_rows) + 1,
            len(itc_compare_headers),
            compute_widths(itc_compare_headers, itc_compare_preview),
            style_func=itc_data_style(itc_compare_headers),
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet9.xml",
            exact_compare_rows(vodafone_headers, vodafone_records, vodafone_vs_mambu_rows),
            len(vodafone_vs_mambu_rows) + 1,
            len(vodafone_compare_headers),
            compute_widths(vodafone_compare_headers, vodafone_compare_preview),
            style_func=itc_data_style(
                vodafone_compare_headers,
                unidentified_style=STYLE_UNIDENTIFIED,
                write_off_status=VODAFONE_WRITE_OFF_STATUS,
                write_off_style=STYLE_WRITE_OFF,
            ),
        )

    os.replace(temp_path, output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a Mambu/Nsano reconciliation Excel template.")
    parser.add_argument(
        "--source",
        default="NSANO Coll Recon.xlsx",
        help="Source workbook containing Mambu, Nsano, Write Off, and Unidentified sheets.",
    )
    parser.add_argument(
        "--mambu",
        nargs="+",
        default=["Mambu-*.xlsx", "Mambu-*.csv"],
        help="Legacy Mambu xlsx/csv file path(s) or glob(s), used only if --source is unavailable.",
    )
    parser.add_argument("--nsano", default="Nsano.csv", help="Legacy Nsano xlsx/csv path, used only if --source is unavailable.")
    parser.add_argument(
        "--output",
        default=None,
        help="Output xlsx path. Defaults to overwriting the source file in place.",
    )
    parser.add_argument(
        "--keep-duplicates",
        action="store_true",
        help="Keep duplicate Mambu reconciliation rows. By default duplicate Mambu key+amount rows are collapsed.",
    )
    return parser.parse_args()


def expand_paths(patterns: list[str]) -> list[Path]:
    paths: list[Path] = []
    for pattern in patterns:
        matches = sorted(glob.glob(pattern))
        if matches:
            paths.extend(Path(match) for match in matches)
        elif not glob.has_magic(pattern):
            paths.append(Path(pattern))
    unique: OrderedDict[str, Path] = OrderedDict()
    for path in paths:
        unique[str(path)] = path
    return list(unique.values())


def main() -> None:
    args = parse_args()
    source_path = Path(args.source)
    mambu_paths = expand_paths(args.mambu)
    nsano_path = Path(args.nsano)
    output_path = source_path if args.output is None else Path(args.output)

    if source_path.exists():
        print(f"Loading source workbook: {source_path}")
        (
            mambu_records,
            mambu_headers,
            nsano_records,
            nsano_headers,
            write_off_records,
            unidentified_records,
            mambu_raw_count,
            nsano_failed_rows_ignored,
        ) = load_coll_recon_workbook(source_path)
        normal_mambu_records = mambu_records
        print(f"  Mambu rows: {len(mambu_records):,}")
        print(f"  Nsano rows included: {len(nsano_records):,}")
        print(f"  Nsano failed rows ignored: {nsano_failed_rows_ignored:,}")
        print(f"  Write-Off rows: {len(write_off_records):,}")
        print(f"  Unidentified rows: {len(unidentified_records):,}")
    else:
        missing = [str(path) for path in mambu_paths if not path.exists()]
        if missing:
            raise FileNotFoundError(f"Mambu file(s) not found: {', '.join(missing)}")
        if not nsano_path.exists():
            raise FileNotFoundError(f"Nsano file not found: {nsano_path}")

        print("Loading legacy Mambu files...")
        mambu_records, mambu_headers, mambu_raw_count = load_mambu(mambu_paths, keep_duplicates=args.keep_duplicates)
        normal_mambu_records, write_off_records, unidentified_records = split_mambu_sources(mambu_records)
        print(f"  Mambu raw rows: {mambu_raw_count:,}")
        print(f"  Mambu comparison rows: {len(mambu_records):,}")
        print(f"  Write-Off rows: {len(write_off_records):,}")
        print(f"  Unidentified rows: {len(unidentified_records):,}")

        print("Loading legacy Nsano file...")
        nsano_records, nsano_headers, nsano_failed_rows_ignored = load_nsano(nsano_path)
        print(f"  Nsano rows included: {len(nsano_records):,}")
        print(f"  Nsano failed rows ignored: {nsano_failed_rows_ignored:,}")

    print("Building comparisons...")
    mambu_vs_nsano_rows = compare_records(mambu_records, nsano_records, "Nsano")
    nsano_vs_mambu_rows = compare_nsano_to_mambu_sources(
        nsano_records,
        normal_mambu_records,
        write_off_records,
        unidentified_records,
    )
    mambu_vs_nsano_stats = summarize_compare(mambu_vs_nsano_rows)
    nsano_vs_mambu_stats = summarize_compare(nsano_vs_mambu_rows)

    summary_rows = build_summary_rows(
        mambu_records,
        nsano_records,
        write_off_records,
        unidentified_records,
        mambu_raw_count,
        mambu_vs_nsano_stats,
        nsano_vs_mambu_stats,
        output_path.name,
    )

    print("Writing workbook...")
    write_workbook(
        output_path,
        summary_rows,
        mambu_headers,
        mambu_records,
        nsano_headers,
        nsano_records,
        write_off_records,
        unidentified_records,
        mambu_vs_nsano_rows,
        nsano_vs_mambu_rows,
    )
    print(f"Done: {output_path}")
    print(
        "Summary: "
        f"Mambu Vs Nsano matched {mambu_vs_nsano_stats['matched']:,}/{mambu_vs_nsano_stats['total']:,}; "
        f"Nsano Vs Mambu matched {nsano_vs_mambu_stats['matched']:,}/{nsano_vs_mambu_stats['total']:,} "
        f"(Mambu {nsano_vs_mambu_stats['mambu']:,}, Write-Off {nsano_vs_mambu_stats['write_off']:,}, "
        f"Unidentified {nsano_vs_mambu_stats['unidentified']:,})."
    )


if __name__ == "__main__":
    main()
