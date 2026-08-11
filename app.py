#!/usr/bin/env python3
"""Streamlit interface for finance reconciliation workflows."""

from __future__ import annotations

import tempfile
import importlib
import hashlib
import io
import re
import zipfile
from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from time import perf_counter

import streamlit as st

from collection_filtering import (
    ITC_COLLECTION_FILTERS,
    MAMBU_COLLECTION_FILTERS,
    MAMBU_DISBURSEMENT_FILTERS,
    NSANO_COLLECTION_FILTERS,
    VODAFONE_COLLECTION_CLEANUP_FILTERS,
    build_itc_collection_filter_workbook,
    build_mambu_collection_filter_workbook,
    build_mambu_disbursement_filter_workbook,
    build_nsano_collection_filter_workbook,
    build_vodafone_collection_cleanup_workbook,
    build_vodafone_wallet_ledger_workbook,
)
from build_reconciliation_template import (
    ITC_MAMBU_STATUS,
    ITC_NOT_FOUND_STATUS,
    ITC_SHEET_COLUMNS,
    ITC_STATUS,
    ITC_UNIDENTIFIED_STATUS,
    ITC_UPSALE_STATUS,
    ITC_WRITE_OFF_STATUS,
    MAMBU_STATUS,
    NOT_FOUND,
    NSANO_SHEET_COLUMNS,
    NSANO_NOT_FOUND_STATUS,
    UNIDENTIFIED_SHEET_COLUMNS,
    UNIDENTIFIED_STATUS,
    VODA_COLL_STATUS,
    VODAFONE_CHARGE_DETAIL,
    WRITE_OFF_STATUS,
    VODAFONE_WRITE_OFF_STATUS,
    WRITE_OFF_ITC_STATUS,
    WRITE_OFF_NSANO_STATUS,
    WRITE_OFF_SHEET_COLUMNS,
    WRITE_OFF_VODAFONE_STATUS,
    WRITE_OFF_ZENITH_STATUS,
    build_write_off_reconciliation,
    clear_source_row_cache as clear_collection_source_row_cache,
    build_itc_summary_rows,
    build_summary_rows,
    compare_itc_to_mambu_sources,
    compare_mambu_to_itc_sources,
    compare_nsano_to_mambu_sources,
    compare_records,
    compare_vodafone_to_mambu_sources,
    iter_csv_dicts,
    iter_xlsx_dicts,
    load_itc_separate_sources,
    load_separate_sources,
    load_write_off_recon_sources,
    not_found_detail_rows,
    normalize_date_amount_match_date,
    sum_amounts,
    summarize_compare,
    worksheet_paths,
    summarize_itc_fee_breakdowns,
    summarize_itc_compare,
    summarize_itc_narration_source_breakdowns,
    summarize_vodafone_charges,
    write_itc_workbook,
    write_workbook,
)
from disbursement_reconciliation import (
    DISB_MAMBU_STATUS,
    DISB_NSANO_STATUS,
    ITC_DISB_SHEET_COLUMNS,
    ITC_DISB_STATUS,
    ITC_WALLET_CREDIT_TRANSFER_SHEET_COLUMNS,
    ITC_WALLET_DEBIT_TRANSFER_SHEET_COLUMNS,
    ITC_WALLET_STATEMENT_SHEET_COLUMNS,
    MTN_MANUAL_MAMBU_STATUS,
    MTN_MANUAL_MATCHED_STATUS,
    MTN_MANUAL_REFUND_SHEET_COLUMNS,
    MTN_MANUAL_REFUND_STATUS,
    MTN_MANUAL_SHEET_COLUMNS,
    NSANO_COLLECTION_MAMBU_SHEET_COLUMNS,
    NSANO_DISB_MAMBU_SHEET_COLUMNS,
    NSANO_DISB_SHEET_COLUMNS,
    NSANO_WALLET_TRANSFER_SHEET_COLUMNS,
    VODAFONE_MANUAL_MAMBU_STATUS,
    VODAFONE_MANUAL_MATCHED_STATUS,
    VODAFONE_MANUAL_REFUND_SHEET_COLUMNS,
    VODAFONE_MANUAL_REFUND_STATUS,
    VODAFONE_MANUAL_SHEET_COLUMNS,
    build_itc_disb_reconciliation,
    build_itc_wallet_ledger_reconciliation,
    build_mtn_manual_disb_reconciliation,
    build_nsano_collection_ledger_reconciliation,
    build_nsano_disb_reconciliation,
    build_nsano_wallet_ledger_reconciliation,
    build_vodafone_manual_disb_reconciliation,
    clear_source_row_cache as clear_disbursement_source_row_cache,
    detect_vodafone_manual_disb_headers,
    load_itc_disb_file,
    load_itc_wallet_ledger_sources,
    load_mambu_disb_file,
    load_mtn_manual_disb_file,
    load_mtn_refund_file,
    load_nsano_disb_file,
    load_nsano_collection_ledger_sources,
    load_nsano_collection_ledger_workbook,
    load_nsano_wallet_ledger_sources,
    load_nsano_wallet_ledger_workbook,
    load_vodafone_manual_disb_file,
    sum_amounts as sum_disb_amounts,
)
from zenith_collection_reconciliation import (
    ZENITH_BANK_SHEET_COLUMNS,
    ZENITH_MATCHED_STATUS,
    ZENITH_MAMBU_STATUS,
    ZENITH_NOT_FOUND_STATUS,
    ZENITH_UNIDENTIFIED_STATUS,
    ZENITH_WRITE_OFF_STATUS,
    build_zenith_collection_reconciliation,
    build_zenith_wallet_ledger_workbook,
    load_zenith_collection_sources,
    sum_zenith_amounts,
)
from monthly_summary import (
    MONTHLY_RESULT_KEYS,
    MONTHLY_SUMMARY_LAYOUT_VERSION,
    build_monthly_summary_bytes,
    monthly_summary_snapshot,
    monthly_summary_status_rows,
)
from master_ledger import (
    master_ledger_detail_rows,
    master_ledger_overview_rows,
)
from collections_master_summary import (
    MASTER_SUMMARY_COLLECTION_KEYS,
    MASTER_SUMMARY_DISBURSEMENT_KEYS,
    MASTER_SUMMARY_LAYOUT_VERSION,
    build_master_summary_bytes,
    collection_exception_summary_rows,
    master_summary_snapshot,
)
from recon_archive import (
    DriveConfig,
    GoogleDriveArchiveClient,
    archive_generated_file,
    check_drive_connection,
    default_drive_config,
    delete_archive_record,
    delete_archive_records,
    list_archive_records,
    prune_missing_drive_records,
    rebuild_drive_index,
    read_archive_file,
    _is_not_found,
)

st.set_page_config(
    page_title="Finance Reconciliation",
    page_icon="📊",
    layout="wide",
)


def _inject_app_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            --recon-bg: #F7FAF9;
            --recon-surface: #FFFFFF;
            --recon-surface-soft: rgba(33, 39, 38, 0.035);
            --recon-border: rgba(33, 39, 38, 0.12);
            --recon-border-strong: rgba(33, 39, 38, 0.2);
            --recon-text: #212726;
            --recon-muted: rgba(33, 39, 38, 0.68);
            --recon-accent: #D6086B;
            --recon-accent-dark: #A90655;
            --recon-accent-soft: rgba(214, 8, 107, 0.08);
            --recon-accent-border: rgba(214, 8, 107, 0.28);
            --recon-danger: #D6086B;
            --recon-shadow: 0 8px 24px rgba(33, 39, 38, 0.06);
        }

        .stApp {
            background: var(--recon-bg);
            color: var(--recon-text);
        }

        [data-testid="stAppViewContainer"] > .main {
            background: var(--recon-bg);
        }

        [data-testid="stHeader"] {
            background: rgba(247, 250, 249, 0.94);
            border-bottom: 1px solid rgba(33, 39, 38, 0.1);
        }

        .block-container {
            max-width: 1480px;
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        .recon-hero {
            position: relative;
            overflow: hidden;
            margin: 0.85rem 0 1.35rem;
            padding: 1.55rem 1.75rem;
            border: 1px solid var(--recon-border);
            border-radius: 14px;
            background:
                radial-gradient(circle at 94% 15%, rgba(214, 8, 107, 0.13), transparent 26%),
                linear-gradient(135deg, #FFFFFF 0%, #FDF8FB 100%);
            box-shadow: var(--recon-shadow);
        }

        .recon-eyebrow {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            margin-bottom: 0.55rem;
            color: var(--recon-accent-dark);
            font-size: 0.76rem;
            font-weight: 750;
            letter-spacing: 0.09em;
            text-transform: uppercase;
        }

        .recon-eyebrow::before {
            content: "";
            width: 0.48rem;
            height: 0.48rem;
            border-radius: 50%;
            background: var(--recon-accent);
            box-shadow: 0 0 0 4px var(--recon-accent-soft);
        }

        .recon-hero h1 {
            margin: 0;
        }

        .recon-hero p {
            max-width: 700px;
            margin: 0.55rem 0 0;
            font-size: 0.98rem;
            line-height: 1.55;
        }

        .recon-sidebar-brand {
            padding: 0.2rem 0 1rem;
            margin-bottom: 0.35rem;
            border-bottom: 1px solid var(--recon-border);
        }

        .recon-sidebar-brand strong {
            display: block;
            color: var(--recon-text);
            font-size: 1rem;
        }

        .recon-sidebar-brand span {
            color: var(--recon-muted);
            font-size: 0.78rem;
        }

        h1, h2, h3 {
            color: var(--recon-text);
            letter-spacing: 0;
        }

        h1 {
            font-size: 2.2rem;
            line-height: 1.12;
            font-weight: 760;
            margin-bottom: 0.45rem;
        }

        h2, h3 {
            font-weight: 720;
        }

        p, label, .stMarkdown, [data-testid="stCaptionContainer"] {
            color: var(--recon-muted);
        }

        hr {
            margin: 1.75rem 0;
            border-color: var(--recon-border);
        }

        section[data-testid="stSidebar"] {
            background: var(--recon-surface);
            border-right: 1px solid var(--recon-border);
        }

        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
            padding-top: 1.35rem;
        }

        [data-baseweb="tab-list"] {
            gap: 0.35rem;
            border-bottom: 1px solid var(--recon-border);
            padding-bottom: 0;
        }

        [data-baseweb="tab"] {
            height: 2.45rem;
            padding: 0 0.85rem;
            border-radius: 8px 8px 0 0;
            color: var(--recon-text);
            font-weight: 700;
            letter-spacing: 0;
        }

        [data-baseweb="tab"] p {
            font-weight: 700 !important;
        }

        [data-baseweb="tab"]:hover {
            background: var(--recon-surface-soft);
            color: var(--recon-accent);
        }

        [data-baseweb="tab"][aria-selected="true"] {
            background: var(--recon-accent-soft);
            color: var(--recon-accent);
            border-bottom: 2px solid var(--recon-accent);
        }

        [data-testid="stFileUploader"] {
            margin-bottom: 0.7rem;
        }

        [data-testid="stFileUploader"] section,
        section[data-testid="stFileUploaderDropzone"] {
            background: var(--recon-surface) !important;
            border: 1px dashed var(--recon-border-strong) !important;
            border-radius: 8px !important;
            min-height: 4.25rem;
            transition: border-color 120ms ease, background-color 120ms ease;
        }

        [data-testid="stFileUploader"] section:hover,
        section[data-testid="stFileUploaderDropzone"]:hover {
            background: var(--recon-surface-soft) !important;
            border-color: var(--recon-accent) !important;
        }

        [data-testid="stFileUploader"] button,
        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {
            border-radius: 8px !important;
            border: 1px solid var(--recon-border-strong) !important;
            font-weight: 600 !important;
            letter-spacing: 0;
        }

        [data-testid="stButton"] button[kind="primary"],
        [data-testid="stDownloadButton"] button[kind="primary"] {
            background: var(--recon-accent) !important;
            border-color: var(--recon-accent) !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 12px rgba(214, 8, 107, 0.18);
        }

        [data-testid="stButton"] button[kind="primary"] p,
        [data-testid="stDownloadButton"] button[kind="primary"] p {
            color: #FFFFFF !important;
            font-weight: 750 !important;
        }

        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {
            transition: transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease;
        }

        [data-testid="stButton"] button:not(:disabled):hover,
        [data-testid="stDownloadButton"] button:not(:disabled):hover {
            transform: translateY(-1px);
            border-color: var(--recon-accent) !important;
            box-shadow: 0 6px 16px rgba(33, 39, 38, 0.10);
        }

        [data-testid="stButton"] button:disabled,
        [data-testid="stDownloadButton"] button:disabled {
            background: rgba(33, 39, 38, 0.06) !important;
            color: #000000 !important;
            border-color: var(--recon-border) !important;
            box-shadow: none !important;
        }

        [data-testid="stButton"] button:disabled p,
        [data-testid="stDownloadButton"] button:disabled p {
            color: #000000 !important;
            font-weight: 600 !important;
        }

        div[data-testid="stMetric"] {
            background: var(--recon-surface);
            border: 1px solid var(--recon-border);
            border-radius: 10px;
            padding: 0.9rem 1rem;
            box-shadow: 0 3px 12px rgba(33, 39, 38, 0.035);
        }

        div[data-testid="stMetric"] label {
            color: var(--recon-muted) !important;
            font-weight: 600;
        }

        div[data-testid="stMetricValue"] {
            color: var(--recon-text);
            font-weight: 760;
        }

        [data-testid="stAlert"] {
            border-radius: 8px;
            border: 1px solid var(--recon-border);
        }

        [data-testid="stExpander"] {
            overflow: hidden;
            border: 1px solid var(--recon-border) !important;
            border-radius: 10px !important;
            background: var(--recon-surface);
        }

        [data-testid="stExpander"] summary:hover {
            color: var(--recon-accent);
        }

        [data-testid="stTable"],
        [data-testid="stDataFrame"] {
            border: 1px solid var(--recon-border);
            border-radius: 8px;
            overflow: hidden;
            background: var(--recon-surface);
        }

        [data-testid="stTable"] table {
            border-collapse: collapse;
        }

        [data-testid="stTable"] thead tr th {
            background: var(--recon-surface-soft);
            color: var(--recon-muted);
            font-weight: 700;
        }

        [data-testid="stTable"] tbody tr:nth-child(even) {
            background: rgba(33, 39, 38, 0.025);
        }

        [data-baseweb="select"] > div,
        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input,
        textarea {
            background: var(--recon-surface) !important;
            border-color: var(--recon-border-strong) !important;
            border-radius: 8px !important;
        }

        [data-baseweb="select"] > div:focus-within,
        [data-testid="stTextInput"] input:focus,
        [data-testid="stNumberInput"] input:focus,
        textarea:focus {
            border-color: var(--recon-accent) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


_inject_app_theme()

st.markdown(
    """
    <div class="recon-hero">
        <div class="recon-eyebrow">Finance Department</div>
        <h1>Finance Reconciliation</h1>
        <p>Run collection, disbursement, ledger, and monthly reconciliation workflows
        from one secure workspace.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

LEGACY_GOOGLE_DRIVE_FOLDER_IDS = {"1lW90g-e8T1N5On343ReXx6KkDwdSXkbR"}
TEST_GOOGLE_DRIVE_FOLDER_ID = "0AOWzc2qLJuEvUk9PVA"
TEST_GOOGLE_DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/u/0/folders/"
    f"{TEST_GOOGLE_DRIVE_FOLDER_ID}"
)


def _configured_drive_folder_id_default() -> str:
    env_config = default_drive_config()
    try:
        drive_secrets = st.secrets.get("google_drive", {})
    except Exception:
        drive_secrets = {}
    return str(
        drive_secrets.get("folder_id", "")
        or (env_config.folder_id if env_config else "")
        or TEST_GOOGLE_DRIVE_FOLDER_ID
    ).strip()


st.session_state.setdefault("archive_month", date.today().strftime("%Y-%m"))
st.sidebar.markdown(
    """
    <div class="recon-sidebar-brand">
        <strong>Reconciliation Hub</strong>
        <span>Archive &amp; workspace settings</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.caption("Archive folders use the detected month from the data.")

configured_drive_folder_id = _configured_drive_folder_id_default()
if (
    st.session_state.get("drive_folder_id_input") in LEGACY_GOOGLE_DRIVE_FOLDER_IDS
    and configured_drive_folder_id not in LEGACY_GOOGLE_DRIVE_FOLDER_IDS
):
    st.session_state["drive_folder_id_input"] = configured_drive_folder_id

drive_folder_id_input = st.sidebar.text_input(
    "Google Drive Folder ID",
    value=configured_drive_folder_id,
    key="drive_folder_id_input",
    help="Defaults to the configured Drive archive folder. Testing fallback: " + TEST_GOOGLE_DRIVE_FOLDER_URL,
)
st.session_state["drive_folder_id"] = drive_folder_id_input.strip()
st.sidebar.checkbox(
    "Save outputs to Drive archive",
    value=True,
    key="archive_to_drive_enabled",
    help="Turn this off for faster master runs; download buttons still appear after the run.",
)
if st.sidebar.button(
    "Refresh archive files",
    key="refresh_archive_files",
    help="Reload saved filtering and reconciliation files from Google Drive.",
):
    st.session_state.pop("_archive_records_cache", None)
    st.session_state.pop("_archive_lookup_error", None)

COLLECTION_MASTER_MAMBU_SHEETS = {
    "Nsano Coll Recon": "Nsano Client Collection",
    "ITC/Voda Coll Recon": "ITC Vodafone Collection",
    "Zenith Coll Recon": "Zenith Collections",
}

DISBURSEMENT_MASTER_MAMBU_SHEETS = {
    "Nsano Disb Recon": "Nsano Client Disb",
    "ITC Wallet vs Mambu": "ITC Disbursement",
    "MTN Manual Disb": "MTN Manual Disb",
    "Vodafone Manual Disb": "Vodafone Manual Disb",
}

COLLECTION_RESULT_KEYS = (
    "nsano_result",
    "itc_result",
    "zenith_result",
    "write_off_recon_result",
)

DISBURSEMENT_RESULT_KEYS = (
    "nsano_disb_result",
    "itc_disb_result",
    "mtn_manual_disb_result",
    "vodafone_manual_disb_result",
)

ARCHIVABLE_RESULTS = {
    "nsano_result": ("Collection", "Nsano Collection Recon", "NSANO Coll Recon.xlsx"),
    "itc_result": ("Collection", "ITC/Voda Collection Recon", "ITC_Voda_Coll_Recon.xlsx"),
    "zenith_result": ("Collection", "Zenith Collection Recon", "Zenith Coll Recon.xlsx"),
    "write_off_recon_result": ("Collection", "Write-off Recon", "Write_off_Recon.xlsx"),
    "nsano_disb_result": ("Disbursement", "Nsano Disbursement Recon", "Nsano_DISB_Recon.xlsx"),
    "itc_disb_result": ("Disbursement", "ITC Wallet vs Mambu", "ITC_Wallet_vs_Mambu.xlsx"),
    "nsano_collection_ledger_result": ("Ledger", "Nsano Collections vs Ledger", "Nsano_Collections_vs_Ledger.xlsx"),
    "nsano_wallet_ledger_result": ("Ledger", "Nsano Disb Wallet vs Ledger", "Nsano_Disb_Wallet_vs_Ledger.xlsx"),
    "itc_wallet_ledger_result": ("Ledger", "ITC Wallet vs Ledger", "ITC_Wallet_vs_Ledger.xlsx"),
    "mtn_manual_disb_result": ("Disbursement", "MTN Manual Disbursement Recon", "MTN_MANUAL_DISB_Recon.xlsx"),
    "vodafone_manual_disb_result": ("Disbursement", "Vodafone Manual Disbursement Recon", "VF_MANUAL_DISB_Recon.xlsx"),
    "mambu_collection_filtering_result": ("Filtering", "Mambu Collection Filtering", "Mambu_Collection_Filtered.xlsx"),
    "itc_collection_filtering_result": ("Filtering", "ITC Collection Filtering", "ITC_Collection_Filtered.xlsx"),
    "nsano_collection_filtering_result": ("Filtering", "Nsano Disb and Collections Filtering", "Nsano_Disb_Collections_Filtered.xlsx"),
    "mambu_disbursement_filtering_result": ("Filtering", "Mambu Disbursement Filtering", "Mambu_Disbursement_Filtered.xlsx"),
    "voda_collection_cleanup_result": ("Filtering", "Voda Collection Cleanup", "Voda_Collection_Cleaned.xlsx"),
    "voda_wallet_ledger_result": ("Ledger", "Voda Wallet vs Ledger", "Voda_Wallet_vs_Ledger.xlsx"),
    "zenith_wallet_ledger_result": ("Ledger", "Zenith Wallet vs Ledger", "Zenith_Wallet_vs_Ledger.xlsx"),
    "collections_summary_sheet_result": ("Summary", "Collections Recon Summary", "Collections_Recon_Summary.xlsx"),
    "disbursement_summary_sheet_result": ("Summary", "Disbursement Recon Summary", "Disbursement_Recon_Summary.xlsx"),
    "monthly_summary_result": ("Summary", "Monthly Reconciliation Summary", "Monthly_Reconciliation_Summary.xlsx"),
}

def uploaded_is_csv(uploaded_file) -> bool:
    return bool(uploaded_file and uploaded_file.name.lower().endswith(".csv"))


def detected_columns(uploaded_file, is_csv: bool = False) -> list[str]:
    """Return column names from the first row of an uploaded file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / uploaded_file.name
        path.write_bytes(uploaded_file.getvalue())
        try:
            rows = iter_csv_dicts(path) if is_csv else iter_xlsx_dicts(path)
            first = next(iter(rows), None)
            return list(first[1].keys()) if first else []
        except Exception:
            return []


def detected_vodafone_manual_columns(uploaded_file) -> list[str]:
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / uploaded_file.name
        path.write_bytes(uploaded_file.getvalue())
        try:
            return detect_vodafone_manual_disb_headers(path)
        except Exception:
            return []


def detected_workbook_sheets(uploaded_file) -> set[str]:
    """Return lower-cased worksheet names from an uploaded workbook."""
    if not uploaded_file:
        return set()
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / uploaded_file.name
        path.write_bytes(uploaded_file.getvalue())
        try:
            with zipfile.ZipFile(path) as zf:
                return set(worksheet_paths(zf))
        except Exception:
            return set()


def column_status_table(raw_cols: list[str], expected_cols: list[str]) -> None:
    """Display which expected columns were found vs missing in the raw file."""
    found = [c for c in expected_cols if c in raw_cols]
    missing = [c for c in expected_cols if c not in raw_cols]
    extra = [c for c in raw_cols if c not in expected_cols]

    c1, c2, c3 = st.columns(3)
    c1.markdown(f"**✅ Found ({len(found)})**")
    for col in found:
        c1.markdown(f"- `{col}`")

    c2.markdown(f"**❌ Missing ({len(missing)})**")
    for col in missing:
        c2.markdown(f"- `{col}`")

    c3.markdown(f"**Extra in raw file ({len(extra)})**")
    for col in extra:
        c3.markdown(f"- `{col}`")


def save_upload(tmp: Path, uploaded_file):
    if not uploaded_file:
        return None
    path = tmp / uploaded_file.name
    path.write_bytes(uploaded_file.getvalue())
    return path


def save_uploads(tmp: Path, uploaded_files) -> list[Path]:
    paths: list[Path] = []
    for index, uploaded_file in enumerate(uploaded_files or [], start=1):
        path = tmp / f"{index:02d}_{Path(uploaded_file.name).name}"
        path.write_bytes(uploaded_file.getvalue())
        paths.append(path)
    return paths


def _google_drive_config() -> DriveConfig | None:
    env_config = default_drive_config()
    try:
        drive_secrets = st.secrets.get("google_drive", {})
    except Exception:
        drive_secrets = {}

    sidebar_folder_id = str(st.session_state.get("drive_folder_id", "")).strip()
    folder_id = str(
        sidebar_folder_id
        or drive_secrets.get("folder_id", "")
        or (env_config.folder_id if env_config else "")
    ).strip()
    if not folder_id:
        return None

    service_account_json = (
        drive_secrets.get("service_account_json")
        or drive_secrets.get("service_account")
        or (env_config.service_account_json if env_config else None)
    )
    service_account_file = str(
        drive_secrets.get("service_account_file", "")
        or (env_config.service_account_file if env_config else "")
    ).strip() or None
    delegated_user = str(
        drive_secrets.get("delegated_user", "")
        or drive_secrets.get("impersonated_user", "")
        or drive_secrets.get("subject", "")
        or (env_config.delegated_user if env_config else "")
    ).strip() or None
    return DriveConfig(
        folder_id=folder_id,
        service_account_json=service_account_json,
        service_account_file=service_account_file,
        delegated_user=delegated_user,
    )


def _current_archive_month() -> str:
    return str(st.session_state.get("archive_month") or date.today().strftime("%Y-%m"))


ARCHIVE_DATE_KEYS = (
    "archive_month",
    "Value Date (Entry Date)",
    "Date/Time",
    "Reconciliation_DateTime",
    "Source_DateTime",
    "Create Date (KEY date - DD/MM/YYYY)",
    "Create Date",
    "Completion Time",
    "Initiation Time",
    "Effect Date",
    "Date (KEY)",
    "DateTime",
    "transaction_date",
    "prepaid_transaction_date",
    "Transfer Date",
    "Credit Date",
    "Debit Date",
    "Mambu Date",
    "Nsano DateTime",
    "ITC Date",
    "Zenith Date",
    "Matched Date",
    "Value Date",
    "DATE",
    "Date",
    "month_title",
    "month_short",
)
ARCHIVE_MONTH_SAMPLE_LIMIT = 1000


def _parse_archive_month(value) -> str | None:
    if value in ("", None):
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    if isinstance(value, date):
        return value.strftime("%Y-%m")

    if isinstance(value, (int, float)) and 20000 <= float(value) <= 80000:
        converted = datetime(1899, 12, 30) + timedelta(days=float(value))
        return converted.strftime("%Y-%m")

    text = str(value or "").strip()
    if not text:
        return None

    # Nsano's "DateTime" column is a JS Date.toString(), e.g. "Mon May 04
    # 2026 14:23:01 GMT+0000 (Coordinated Universal Time)" — pull this out of
    # the *original* text first: the generic cleanup below replaces every "T"
    # with a space (for ISO 8601's "T" separator), which also mangles the "T"
    # inside "GMT"/"Time" and destroys the marker this split relies on.
    nsano_style = text.split(" GMT", 1)[0].strip()

    if re.fullmatch(r"20\d{2}-(0?[1-9]|1[0-2])", text):
        year, month = text.split("-", 1)
        return f"{int(year):04d}-{int(month):02d}"

    embedded_year_month = re.search(r"(?<!\d)(20\d{2})[-_/](0?[1-9]|1[0-2])(?!\d)", text)
    if embedded_year_month:
        return f"{int(embedded_year_month.group(1)):04d}-{int(embedded_year_month.group(2)):02d}"

    try:
        numeric = float(text.replace(",", ""))
    except ValueError:
        numeric = None
    if numeric is not None and 20000 <= numeric <= 80000:
        converted = datetime(1899, 12, 30) + timedelta(days=numeric)
        return converted.strftime("%Y-%m")

    cleaned = (
        text.replace("T", " ")
        .replace("\u00a0", " ")
        .strip()
        .rstrip("Z")
    )
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(r"\s*[+-]\d{2}:?\d{2}$", "", cleaned)
    candidates = [
        cleaned,
        cleaned.split(".", 1)[0],
        cleaned.split(" ", 1)[0],
        nsano_style,
    ]
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
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
        "%B %Y",
        "%b %Y",
        "%a %b %d %Y %H:%M:%S",
    )
    for candidate in candidates:
        for fmt in formats:
            for parsed_candidate in (candidate, candidate.title()):
                try:
                    return datetime.strptime(parsed_candidate, fmt).strftime("%Y-%m")
                except ValueError:
                    continue

    month_name_text = re.sub(r"[_-]+", " ", cleaned)
    month_name_match = re.search(
        r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
        r"\s+20\d{2}\b",
        month_name_text,
        flags=re.IGNORECASE,
    )
    if month_name_match:
        return _parse_archive_month(month_name_match.group(0).title())

    return None


def _mapping_value(mapping: Mapping, wanted_key: str):
    if wanted_key in mapping:
        return mapping.get(wanted_key)
    wanted = wanted_key.casefold()
    for key, value in mapping.items():
        if str(key).casefold() == wanted:
            return value
    return None


def _month_values_from_mapping(mapping: Mapping):
    yielded_keys: set[str] = set()
    for header in ARCHIVE_DATE_KEYS:
        value = _mapping_value(mapping, header)
        if value not in ("", None):
            yielded_keys.add(header.casefold())
            yield value

    for key, value in mapping.items():
        key_text = str(key or "")
        key_fold = key_text.casefold()
        if key_fold in yielded_keys:
            continue
        if "birth date" in key_fold:
            continue
        if "date" in key_fold or key_fold in {"month", "month title", "month_title", "month short", "month_short"}:
            if value not in ("", None):
                yield value


def _month_values_from_source(source):
    if source in ("", None):
        return

    datetime_value = getattr(source, "datetime_value", None)
    if datetime_value not in ("", None):
        yield datetime_value

    if any(hasattr(source, attr) for attr in ("sheet_values", "raw")):
        source_file = getattr(source, "source_file", None)
        if source_file:
            yield source_file
        return

    for attr in ("sheet_values", "raw"):
        mapping = getattr(source, attr, None)
        if isinstance(mapping, Mapping):
            yield from _month_values_from_mapping(mapping)

    if isinstance(source, Mapping):
        yield from _month_values_from_mapping(source)
        for key, value in source.items():
            if str(key).casefold() in {"output_bytes", "archive_record"}:
                continue
            if isinstance(value, (Mapping, list, tuple, set)):
                yield from _month_values_from_source(value)
        return

    if isinstance(source, (str, bytes, bytearray)):
        yield source
        return

    if isinstance(source, (list, tuple, set)):
        for item in source:
            yield from _month_values_from_source(item)
        return

    if isinstance(source, (datetime, date, int, float)):
        yield source

    source_file = getattr(source, "source_file", None)
    if source_file:
        yield source_file


def _data_month_from_sources(*sources) -> str | None:
    month_counts: Counter[str] = Counter()
    parsed_months = 0
    for source in sources:
        for value in _month_values_from_source(source):
            month = _parse_archive_month(value)
            if month:
                month_counts[month] += 1
                parsed_months += 1
                if parsed_months >= ARCHIVE_MONTH_SAMPLE_LIMIT:
                    return month_counts.most_common(1)[0][0]
    if not month_counts:
        return None
    return month_counts.most_common(1)[0][0]


_FILTERING_MONTH_SAMPLE_ROWS = 200


def _filtering_data_month(source_paths: list[Path]) -> str | None:
    """Detect the transaction month from raw uploaded files, so a filtering
    run done in one calendar month but covering an earlier month's data gets
    archived under the data's month, not the month it happened to be run.

    Scans every sheet of an .xlsx upload (not just the first) — some raw
    exports (e.g. Nsano's) put the dated transaction rows on a sheet other
    than the first, which otherwise makes date detection silently fail and
    fall back to today's date."""
    sample_rows: list[dict] = []
    for path in source_paths:
        if len(sample_rows) >= _FILTERING_MONTH_SAMPLE_ROWS:
            break
        try:
            if path.suffix.lower() == ".csv":
                row_iterators = [iter_csv_dicts(path)]
            else:
                try:
                    with zipfile.ZipFile(path) as zf:
                        sheet_names = list(worksheet_paths(zf))
                except Exception:
                    sheet_names = []
                row_iterators = (
                    [iter_xlsx_dicts(path, sheet_name) for sheet_name in sheet_names]
                    if sheet_names
                    else [iter_xlsx_dicts(path)]
                )
            for rows in row_iterators:
                if len(sample_rows) >= _FILTERING_MONTH_SAMPLE_ROWS:
                    break
                for _row_number, row in rows:
                    sample_rows.append(row)
                    if len(sample_rows) >= _FILTERING_MONTH_SAMPLE_ROWS:
                        break
        except Exception:
            continue
    return _data_month_from_sources(sample_rows)


def _master_archive_month_for_keys(*keys: str) -> str | None:
    month_counts: Counter[str] = Counter()
    for key in keys:
        result = st.session_state.get(key)
        if not isinstance(result, dict):
            continue
        month = _parse_archive_month(result.get("archive_month"))
        if month and result.get("archive_scope") == "master":
            month_counts[month] += 1
    if not month_counts:
        return None
    return month_counts.most_common(1)[0][0]


def _mark_master_archive(result: dict, archive_month: str | None = None) -> dict:
    result["archive_scope"] = "master"
    result["archive_month"] = archive_month or _current_archive_month()
    return result


def _apply_cached_archive_scope(cached: dict, archive_month: str | None) -> None:
    if archive_month:
        _mark_master_archive(cached, archive_month)
    else:
        cached.pop("archive_scope", None)
        cached.pop("archive_month", None)


def _monthly_archive_month(snapshot: Mapping[str, object]) -> str:
    return (
        _master_archive_month_for_keys(*MONTHLY_RESULT_KEYS)
        or _data_month_from_sources(snapshot)
        or _current_archive_month()
    )


def _drive_archive_enabled() -> bool:
    return bool(st.session_state.get("archive_to_drive_enabled", True))


def _build_shared_drive_client(drive_config: DriveConfig | None) -> GoogleDriveArchiveClient | None:
    """Build one Drive client to reuse across multiple archive uploads in a run.

    Reusing the client avoids re-authenticating and re-discovering the same
    month/category subfolders and index file for every archived result.
    """
    if not drive_config:
        return None
    try:
        return GoogleDriveArchiveClient(drive_config)
    except Exception:
        return None  # fall back to per-call client construction in archive_generated_file


def _ensure_result_archived(
    result_key: str,
    drive_config: DriveConfig | None,
    client: GoogleDriveArchiveClient | None = None,
) -> bool:
    if not _drive_archive_enabled():
        return False
    spec = ARCHIVABLE_RESULTS.get(result_key)
    result = st.session_state.get(result_key)
    if not spec or not isinstance(result, dict):
        return False
    if result.get("archive_scope") != "master":
        return False
    output_bytes = result.get("output_bytes")
    if not isinstance(output_bytes, (bytes, bytearray)):
        return False

    archive_month = _parse_archive_month(result.get("archive_month")) or _current_archive_month()
    signature = hashlib.sha256(output_bytes).hexdigest()
    archive_record = result.get("archive_record") or {}
    already_saved = (
        result.get("archive_signature") == signature
        and archive_record.get("month") == archive_month
        and archive_record.get("drive_saved")
    )
    if already_saved:
        return True

    category, workflow, file_name = spec
    archive_result = archive_generated_file(
        content=bytes(output_bytes),
        file_name=file_name,
        workflow=workflow,
        category=category,
        month=archive_month,
        archive_scope="master",
        drive_config=drive_config,
        client=client,
    )
    result["archive_signature"] = signature
    result["archive_record"] = archive_result.record
    if archive_result.drive_error:
        result["archive_drive_error"] = archive_result.drive_error
        result["archive_drive_attempt_signature"] = signature
        return False
    result.pop("archive_drive_error", None)
    result.pop("archive_drive_attempt_signature", None)
    return archive_result.drive_saved


def _archive_master_results(result_keys: tuple[str, ...]) -> tuple[int, list[str]]:
    if not _drive_archive_enabled():
        return 0, []
    drive_config = _google_drive_config()
    client = _build_shared_drive_client(drive_config)
    saved = 0
    errors: list[str] = []
    for result_key in result_keys:
        result = st.session_state.get(result_key)
        if not isinstance(result, dict) or result.get("archive_scope") != "master":
            continue
        if _ensure_result_archived(result_key, drive_config, client):
            saved += 1
            continue
        workflow = ARCHIVABLE_RESULTS.get(result_key, ("", result_key, ""))[1]
        errors.append(f"{workflow}: {result.get('archive_drive_error', 'archive save failed')}")
    return saved, errors


def _archive_warning_text(errors: list[str]) -> str:
    grouped: dict[str, list[str]] = {}
    for error in errors:
        workflow, separator, message = str(error).partition(": ")
        if not separator:
            workflow = "Archive"
            message = str(error)
        grouped.setdefault(message, []).append(workflow)

    parts: list[str] = []
    for message, workflows in grouped.items():
        workflow_text = ", ".join(workflows)
        parts.append(f"{workflow_text}: {message}")
    return "Archive save issue: " + "; ".join(parts)


def _record_timing(timings: list[dict[str, object]], stage: str, started_at: float) -> None:
    timings.append({"stage": stage, "seconds": perf_counter() - started_at})


def _timing_rows(timings: list[dict[str, object]]) -> list[dict[str, str]]:
    return [
        {
            "Stage": str(timing.get("stage", "")),
            "Seconds": f"{float(timing.get('seconds') or 0):,.2f}",
        }
        for timing in timings
    ]


def _timed(timings: list[dict[str, object]], stage: str, callback, *args, **kwargs):
    started_at = perf_counter()
    try:
        return callback(*args, **kwargs)
    finally:
        _record_timing(timings, stage, started_at)


def _archive_available_outputs(drive_config: DriveConfig | None) -> None:
    client = _build_shared_drive_client(drive_config)
    for result_key in ARCHIVABLE_RESULTS:
        _ensure_result_archived(result_key, drive_config, client)


def _drive_storage_setup_blocked(message: object) -> bool:
    return "service accounts do not have personal Drive storage quota" in str(message)


def _drive_storage_setup_message() -> str:
    return (
        "Drive archive setup required: the configured folder is in My Drive. "
        "Use a Shared Drive folder ID and add the service account as a member with write access, "
        "or configure domain-wide delegation with GOOGLE_DRIVE_DELEGATED_USER."
    )


def _format_bytes(size: object) -> str:
    try:
        value = float(size or 0)
    except (TypeError, ValueError):
        value = 0.0
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:,.1f} {unit}" if unit != "B" else f"{int(value):,} B"
        value /= 1024
    return f"{value:,.1f} GB"


def _month_label(month: str) -> str:
    try:
        return date.fromisoformat(f"{month}-01").strftime("%B %Y")
    except ValueError:
        return month


def _archive_record_is_previous(record: Mapping[str, object]) -> bool:
    return bool(record.get("superseded_by")) or str(record.get("version_status", "")).casefold() == "previous"


def _archive_record_version_label(record: Mapping[str, object]) -> str:
    return "Previous" if _archive_record_is_previous(record) else "Current"


def _archive_record_label(record: dict) -> str:
    created = str(record.get("created_at", ""))[:19].replace("T", " ")
    record_id = str(record.get("id", ""))[:8]
    version = _archive_record_version_label(record)
    return f"{record.get('month', '')} | {record.get('category', '')} | {record.get('workflow', '')} | {version} | {created} | {record_id}"


def _archive_zip_bytes(records: list[dict], drive_config: DriveConfig | None) -> tuple[bytes, list[str]]:
    buffer = io.BytesIO()
    skipped: list[str] = []
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for record in records:
            try:
                content = read_archive_file(record, drive_config)
            except Exception:
                skipped.append(str(record.get("file_name") or record.get("workflow") or "Unknown file"))
                continue
            name = (
                f"{record.get('month', 'unknown')}/"
                f"{record.get('category', 'Archive')}/"
                f"{record.get('id', '')[:8]}_{record.get('file_name', 'recon.xlsx')}"
            )
            zf.writestr(name, content)
    return buffer.getvalue(), skipped


@dataclass(frozen=True)
class ArchivedFileSource:
    """Wraps bytes downloaded from the Drive archive so they satisfy the same
    `.name` / `.getvalue()` contract as a Streamlit `UploadedFile`, letting
    them flow through `save_upload`/`save_uploads` unmodified."""
    name: str
    content: bytes

    def getvalue(self) -> bytes:
        return self.content


@dataclass(frozen=True)
class _ArchivedFilePlaceholder:
    """Render-time stand-in for an archived pick. Carries only the file name
    (no bytes), so truthiness/`.name` checks work before Run is clicked
    without downloading anything from Drive on every rerender."""
    name: str


ARCHIVE_RECORD_CACHE_SECONDS = 45.0
_archive_lookup_notice_shown = False


def _list_filtering_archive_records(drive_config: DriveConfig | None) -> list[dict]:
    """Fetch the archive index once per tab render, limited to current (not
    superseded) Filtering-category records."""
    global _archive_lookup_notice_shown
    cache_key = drive_config.folder_id if drive_config else ""
    cached = st.session_state.get("_archive_records_cache", {})
    cache_age = perf_counter() - float(cached.get("fetched_at", 0))
    if cached.get("folder_id") == cache_key and cache_age < ARCHIVE_RECORD_CACHE_SECONDS:
        records = cached.get("records", [])
        error = cached.get("error")
    else:
        records, error = list_archive_records(drive_config)
        st.session_state["_archive_records_cache"] = {
            "folder_id": cache_key,
            "fetched_at": perf_counter(),
            "records": records,
            "error": error,
        }
        st.session_state["_archive_lookup_error"] = str(error or "")

    filtering_records = [
        record for record in records
        if record.get("category") == "Filtering" and not _archive_record_is_previous(record)
    ]
    if not _archive_lookup_notice_shown:
        if error:
            st.warning(
                "Saved-file auto-pick is unavailable because the Google Drive archive "
                f"could not be read: {error}"
            )
            _archive_lookup_notice_shown = True
        elif not filtering_records:
            st.info(
                "No current filtering outputs are indexed in the configured Drive archive. "
                "Run and archive the filtering workflows, or upload source files manually."
            )
            _archive_lookup_notice_shown = True
    return filtering_records


def _filtered_source_picker(
    *,
    label: str,
    field_key: str,
    workflow: str,
    records: list[dict],
    file_types: list[str],
) -> tuple[str, object]:
    """Render a source-selection control for one recon input field.

    Returns ("uploaded", value) or ("archived", record) — never touches
    Drive itself; use `_render_time_value` / `_materialize_picked_source`
    to turn the result into something usable.
    """
    matches = sorted(
        (record for record in records if record.get("workflow") == workflow),
        key=lambda record: str(record.get("created_at", "")),
        reverse=True,
    )
    current_month = _current_archive_month()
    default_record = next((record for record in matches if record.get("month") == current_month), None)

    if not matches:
        return ("uploaded", st.file_uploader(label, type=file_types, key=f"{field_key}_upload"))

    def _record_option_label(record: dict) -> str:
        created = str(record.get("created_at", ""))[:19].replace("T", " ")
        return f"{_month_label(str(record.get('month', '')))} — {record.get('file_name', '')} (saved {created})"

    mode = st.radio(
        label,
        ["Use saved filtered file", "Upload new file"],
        index=0 if default_record else 1,
        horizontal=True,
        key=f"{field_key}_mode",
    )

    if mode == "Use saved filtered file":
        default_index = matches.index(default_record) if default_record in matches else 0
        chosen = st.selectbox(
            "Saved filtered file",
            matches,
            index=default_index,
            format_func=_record_option_label,
            key=f"{field_key}_record",
            label_visibility="collapsed",
        )
        st.caption(f"Using: {_record_option_label(chosen)}")
        return ("archived", chosen)

    return ("uploaded", st.file_uploader("Upload new file", type=file_types, key=f"{field_key}_upload", label_visibility="collapsed"))


def _render_time_value(picked: tuple[str, object]):
    """Turn a picker result into None / UploadedFile / _ArchivedFilePlaceholder
    for use at render time, before Run is clicked."""
    mode, value = picked
    if mode == "archived" and value:
        return _ArchivedFilePlaceholder(name=str(value.get("file_name", "archived_file.xlsx")))
    return value


def _archive_record_missing_from_drive(record: Mapping[str, object], drive_config: DriveConfig | None) -> bool:
    """True if the record's Drive file is gone or trashed. Deleting via the
    Drive UI normally moves a file to trash rather than purging it outright,
    so a plain download can still succeed on a "deleted" file — check status
    explicitly instead of relying on a download failure to notice."""
    drive_file_id = str(record.get("drive_file_id", "")).strip()
    if not drive_config or not drive_file_id:
        return True
    try:
        client = GoogleDriveArchiveClient(drive_config)
        file_meta = client.service.files().get(
            fileId=drive_file_id, fields="id,trashed", supportsAllDrives=True,
        ).execute()
        return bool(file_meta.get("trashed"))
    except Exception as exc:
        return _is_not_found(exc)


def _materialize_picked_source(picked: tuple[str, object], drive_config: DriveConfig | None):
    """Resolve a picker result into something save_upload()-compatible.
    Call only inside `if run_clicked:` blocks — this is the one place an
    archived pick's bytes get downloaded from Drive.

    Deliberately does not validate the pick against Drive at render time —
    that would mean a Drive round trip on every rerun of the tab just to
    draw a dropdown. Staleness (someone deleted the file/folder directly in
    Drive) is only discovered here, at the point of actual use, and is
    self-healed by dropping just that one record from the index so it won't
    show up again — no per-render cost either way.
    """
    mode, value = picked
    if mode == "archived" and value:
        if _archive_record_missing_from_drive(value, drive_config):
            delete_archive_record(value, drive_config)
            raise FileNotFoundError(
                f"\"{value.get('file_name', 'The selected saved file')}\" is no longer in Google Drive "
                "(it looks like it was deleted there directly) — it's been removed from the list. "
                "Pick a different saved file or upload a new one and run again."
            )
        content = read_archive_file(value, drive_config)
        return ArchivedFileSource(name=str(value.get("file_name", "archived_file.xlsx")), content=content)
    return value


def status_count(stats: dict, status: str) -> int:
    return int(stats["status_counts"].get(status, 0))


def _arrow_safe_rows(rows: list[dict]) -> list[dict]:
    """Convert table rows to stable text values for Streamlit/PyArrow display."""
    def _safe(v):
        if v is None:
            return ""
        if isinstance(v, Decimal):
            return f"{v:,.2f}"
        if isinstance(v, bool):
            return "Yes" if v else "No"
        if isinstance(v, int):
            return f"{v:,}"
        if isinstance(v, float):
            return f"{v:,.2f}"
        return str(v)
    return [{str(k): _safe(v) for k, v in row.items()} for row in rows]


_MASTER_SUMMARY_SCOPES = {
    "collection": (MASTER_SUMMARY_COLLECTION_KEYS, "collections_summary_sheet_result", "Collections Recon Summary", "Collections_Recon_Summary.xlsx"),
    "disbursement": (MASTER_SUMMARY_DISBURSEMENT_KEYS, "disbursement_summary_sheet_result", "Disbursement Recon Summary", "Disbursement_Recon_Summary.xlsx"),
}


def _render_master_summary_download(download_key: str, *, scope: str) -> None:
    """Render a Master Summary download scoped to one domain (collection or disbursement).

    Each page only ever summarizes its own workflows, even if the other domain
    also ran earlier in the session — combining both is reserved for the
    Monthly Summary export.
    """
    keys, result_key, file_label, file_name = _MASTER_SUMMARY_SCOPES[scope]
    master_snapshot = master_summary_snapshot(st.session_state, keys)
    if not master_snapshot:
        return

    signature = f"{MASTER_SUMMARY_LAYOUT_VERSION}:{repr(master_snapshot)}"
    cached = st.session_state.get(result_key)
    archive_month = _master_archive_month_for_keys(*keys)
    if not cached or cached.get("signature") != signature:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                output_bytes = build_master_summary_bytes(
                    master_snapshot,
                    Path(tmpdir) / file_name,
                )
        except Exception as exc:
            st.error(f"{file_label} generation failed: {exc}")
            return
        cached = {
            "signature": signature,
            "output_bytes": output_bytes,
        }
        st.session_state[result_key] = cached
    _apply_cached_archive_scope(cached, archive_month)
    if archive_month:
        with st.spinner(f"Saving {file_label} to Google Drive archive..."):
            _ensure_result_archived(result_key, _google_drive_config())

    st.caption("This master report includes a Daily Summary overview and separate daily wallet sheets.")
    st.download_button(
        label=f"Download {file_label} (with Daily Summary)",
        data=cached["output_bytes"],
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key=download_key,
        on_click="ignore",
    )


def render_column_inspector(files_to_inspect: list[tuple[str, object, list[str], bool]]) -> None:
    with st.expander("🔍 Column Inspector — click to verify column names match", expanded=True):
        for label, uploaded_file, expected, is_csv in files_to_inspect:
            if not uploaded_file:
                continue
            st.markdown(f"##### {label}")
            raw_cols = (
                detected_vodafone_manual_columns(uploaded_file)
                if label == "VODAFONE MANUAL"
                else detected_columns(uploaded_file, is_csv=is_csv)
            )
            if not raw_cols and label == "VODAFONE MANUAL":
                raw_cols = detected_columns(uploaded_file, is_csv=is_csv)
            if not raw_cols:
                st.warning(f"Could not read columns from {uploaded_file.name}")
            else:
                column_status_table(raw_cols, expected)
            st.markdown("")


def render_collection_master_sheet_status(uploaded_file) -> None:
    if not uploaded_file:
        return

    found_sheets = detected_workbook_sheets(uploaded_file)
    if not found_sheets:
        st.warning("Could not read sheets from the filtered Mambu workbook.")
        return

    rows = []
    for workflow, sheet_name in COLLECTION_MASTER_MAMBU_SHEETS.items():
        rows.append({
            "Workflow": workflow,
            "Mambu Sheet": sheet_name,
            "Status": "Found" if sheet_name.casefold() in found_sheets else "Missing",
        })
    st.markdown("#### Filtered Mambu Sheets")
    st.table(rows)


def render_disbursement_master_sheet_status(uploaded_file) -> None:
    if not uploaded_file:
        return

    found_sheets = detected_workbook_sheets(uploaded_file)
    if not found_sheets:
        st.warning("Could not read sheets from the filtered Mambu Disbursement workbook.")
        return

    rows = []
    for workflow, sheet_name in DISBURSEMENT_MASTER_MAMBU_SHEETS.items():
        rows.append({
            "Workflow": workflow,
            "Mambu Sheet": sheet_name,
            "Status": "Found" if sheet_name.casefold() in found_sheets else "Missing",
        })
    st.markdown("#### Filtered Mambu Disbursement Sheets")
    st.table(rows)


def _disbursement_master_mambu_sheet_candidates(workflow: str) -> list[str | None]:
    preferred = DISBURSEMENT_MASTER_MAMBU_SHEETS[workflow]
    if workflow in {"MTN Manual Disb", "Vodafone Manual Disb"}:
        return [preferred]
    return [preferred, None]


def _load_disbursement_master_mambu_records(path: Path, workflow: str, found_sheets: set[str]):
    candidates = _disbursement_master_mambu_sheet_candidates(workflow)
    last_loaded = ([], [])
    for sheet_name in candidates:
        if sheet_name and sheet_name.casefold() not in found_sheets:
            continue
        mambu_records, mambu_headers = load_mambu_disb_file(path, sheet_name)
        last_loaded = (mambu_records, mambu_headers)
        if mambu_records or sheet_name is None:
            return mambu_records, mambu_headers, sheet_name or "First sheet"
    if workflow in {"MTN Manual Disb", "Vodafone Manual Disb"}:
        return [], [], DISBURSEMENT_MASTER_MAMBU_SHEETS[workflow]
    return last_loaded[0], last_loaded[1], "First sheet"


def _disbursement_key_overlap(source_records: list, mambu_records: list) -> dict[str, object]:
    source_keys = {record.key for record in source_records if record.key}
    mambu_keys = {record.key for record in mambu_records if record.key}
    overlap = source_keys & mambu_keys
    return {
        "source_key_count": len(source_keys),
        "mambu_key_count": len(mambu_keys),
        "overlap_count": len(overlap),
        "source_samples": sorted(source_keys)[:3],
        "mambu_samples": sorted(mambu_keys)[:3],
    }


def _warn_zero_disbursement_overlap(
    workflow: str,
    source_records: list,
    mambu_records: list,
    mambu_sheet_used: str,
) -> dict[str, object]:
    diagnostics = _disbursement_key_overlap(source_records, mambu_records)
    if mambu_records and not diagnostics["mambu_key_count"]:
        st.warning(
            f"{workflow}: no Mambu identifiers were found in `{mambu_sheet_used}`. "
            "Check that the sheet contains an Identifier column."
        )
    elif source_records and not diagnostics["source_key_count"]:
        st.warning(
            f"{workflow}: no source identifiers were found. Check that the uploaded source "
            "contains the expected transaction ID column."
        )
    elif (
        diagnostics["source_key_count"]
        and diagnostics["mambu_key_count"]
        and not diagnostics["overlap_count"]
    ):
        st.warning(
            f"{workflow}: no normalized identifiers overlap with Mambu sheet "
            f"`{mambu_sheet_used}`. Check that the uploaded files are from the same month and "
            "that their Identifier columns contain the transaction IDs."
        )
    return diagnostics


def _nsano_coll_diagnostics(
    mambu_records: list,
    nsano_records: list,
    nsano_vs_mambu_rows: list,
    mvn_stats: dict,
    nvm_stats: dict,
) -> dict:
    from collections import Counter as _Counter
    mambu_key_counts = _Counter(r.key for r in mambu_records if r.key)
    nsano_key_counts = _Counter(r.key for r in nsano_records if r.key)
    mambu_dupes = {k: v for k, v in mambu_key_counts.items() if v > 1}
    nsano_dupes = {k: v for k, v in nsano_key_counts.items() if v > 1}
    mambu_blank = sum(1 for r in mambu_records if not r.key)
    nsano_blank = sum(1 for r in nsano_records if not r.key)

    mambu_key_set = {r.key for r in mambu_records if r.key}
    diverted: list[dict] = []
    for row in nsano_vs_mambu_rows:
        status = str(row.get("Match_Status", ""))
        if status in {WRITE_OFF_STATUS, UNIDENTIFIED_STATUS}:
            key = str(row.get("Source_Key", ""))
            if key and key in mambu_key_set:
                diverted.append({
                    "Key": key,
                    "Matched To": status,
                    "Nsano Amount": row.get("Source_Amount", ""),
                    "Date": row.get("Source_DateTime", ""),
                })

    mvn_matched = int(mvn_stats.get("matched", 0))
    nvm_mambu = int((nvm_stats.get("status_counts") or {}).get(MAMBU_STATUS, 0))
    mvn_amount = _decimal_amount(mvn_stats.get("matched_amount", 0))
    nvm_amount = _decimal_amount((nvm_stats.get("status_amounts") or {}).get(MAMBU_STATUS, 0))

    return {
        "mambu_dupe_count": len(mambu_dupes),
        "mambu_dupe_examples": sorted(mambu_dupes.items(), key=lambda x: -x[1])[:5],
        "nsano_dupe_count": len(nsano_dupes),
        "nsano_dupe_examples": sorted(nsano_dupes.items(), key=lambda x: -x[1])[:5],
        "mambu_blank_count": mambu_blank,
        "nsano_blank_count": nsano_blank,
        "diverted_count": len(diverted),
        "diverted_examples": diverted[:10],
        "mvn_matched": mvn_matched,
        "nvm_mambu": nvm_mambu,
        "mvn_matched_amount": mvn_amount,
        "nvm_mambu_amount": nvm_amount,
        "count_gap": mvn_matched - nvm_mambu,
        "amount_gap": mvn_amount - nvm_amount,
    }


def _render_collection_data_diagnostics(state: dict) -> None:
    nsano = state.get("nsano_result", {})
    diag = nsano.get("collection_diagnostics") if isinstance(nsano, dict) else None
    if not diag:
        return

    st.markdown("#### Data Quality Checks")

    def _flag(count: int, label: str, ok_label: str = "") -> str:
        if count:
            return f"⚠️ {count:,} {label}"
        return f"✓ {ok_label}" if ok_label else "✓ OK"

    count_gap = diag["count_gap"]
    amount_gap = diag["amount_gap"]
    asymmetry_ok = count_gap == 0 and amount_gap == 0

    overview = [
        {"Check": "Mambu duplicate keys", "Result": _flag(diag["mambu_dupe_count"], "duplicate key(s)", "No duplicates")},
        {"Check": "Nsano duplicate keys", "Result": _flag(diag["nsano_dupe_count"], "duplicate key(s)", "No duplicates")},
        {"Check": "Mambu blank keys", "Result": _flag(diag["mambu_blank_count"], "record(s) with no key", "None")},
        {"Check": "Nsano blank keys", "Result": _flag(diag["nsano_blank_count"], "record(s) with no key", "None")},
        {"Check": "Nsano rows diverted to Write-Off / Unidentified (key also in Mambu)", "Result": _flag(diag["diverted_count"], "row(s) affected", "None")},
        {
            "Check": "Mambu vs Nsano ↔ Nsano vs Mambu (count)",
            "Result": f"✓ Balanced ({diag['mvn_matched']:,})" if count_gap == 0
            else f"⚠️ Gap of {abs(count_gap):,} — Mambu side: {diag['mvn_matched']:,} | Nsano side: {diag['nvm_mambu']:,}",
        },
        {
            "Check": "Mambu vs Nsano ↔ Nsano vs Mambu (amount GHC)",
            "Result": f"✓ Balanced" if amount_gap == 0
            else f"⚠️ Gap of {abs(amount_gap):,.2f} — Mambu side: {diag['mvn_matched_amount']:,.2f} | Nsano side: {diag['nvm_mambu_amount']:,.2f}",
        },
    ]
    st.table(overview)

    any_issue = (
        diag["mambu_dupe_count"]
        or diag["nsano_dupe_count"]
        or diag["mambu_blank_count"]
        or diag["nsano_blank_count"]
        or diag["diverted_count"]
        or not asymmetry_ok
    )

    if not any_issue:
        st.success("No data quality issues detected for Nsano collection.")
        return

    if diag["mambu_dupe_count"] or diag["nsano_dupe_count"]:
        with st.expander(f"Duplicate keys detail ({diag['mambu_dupe_count']} Mambu / {diag['nsano_dupe_count']} Nsano)"):
            if diag["mambu_dupe_examples"]:
                st.markdown("**Mambu** — keys appearing more than once:")
                st.table([{"Key": k, "Count": v} for k, v in diag["mambu_dupe_examples"]])
            if diag["nsano_dupe_examples"]:
                st.markdown("**Nsano** — keys appearing more than once:")
                st.table([{"Key": k, "Count": v} for k, v in diag["nsano_dupe_examples"]])
            st.caption(
                "Duplicate keys cause one side to count more matches than the other. "
                "Each side takes the first hit, so the side with duplicates reports higher matched counts."
            )

    if diag["diverted_count"]:
        with st.expander(f"Diverted rows — {diag['diverted_count']} Nsano row(s) sent to Write-Off / Unidentified whose key also exists in Mambu"):
            st.table(diag["diverted_examples"])
            st.caption(
                "These Nsano rows were matched to Write-Off or Unidentified first (priority order). "
                "Because their key also exists in Mambu, the Mambu side counts them as matched — "
                "but the Nsano side does not count them as Mambu matches. This directly causes mvn > nvm."
            )


def _nsano_data_issues_label(nsano: dict | None) -> str:
    if not nsano:
        return ""
    diag = nsano.get("collection_diagnostics")
    if not diag:
        return ""
    issues: list[str] = []
    if diag["mambu_dupe_count"]:
        issues.append(f"{diag['mambu_dupe_count']} Mambu dupe key(s)")
    if diag["nsano_dupe_count"]:
        issues.append(f"{diag['nsano_dupe_count']} Nsano dupe key(s)")
    if diag["mambu_blank_count"]:
        issues.append(f"{diag['mambu_blank_count']} Mambu blank key(s)")
    if diag["nsano_blank_count"]:
        issues.append(f"{diag['nsano_blank_count']} Nsano blank key(s)")
    if diag["diverted_count"]:
        issues.append(f"{diag['diverted_count']} diverted row(s)")
    if diag["count_gap"] != 0:
        issues.append(f"count gap {diag['count_gap']:+,}")
    if diag["amount_gap"] != 0:
        issues.append(f"amount gap {diag['amount_gap']:+,.2f}")
    return "; ".join(issues) if issues else "✓ OK"


def _collection_master_status_rows(state: dict) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    nsano = state.get("nsano_result")
    rows.append({
        "Workflow": "Nsano Coll Recon",
        "Status": "Ready" if nsano else "Not run",
        "Rows": f"{nsano['mambu_count']:,} Mambu / {nsano['nsano_count']:,} Nsano" if nsano else "",
        "Match Rate": f"{nsano['mvn']['match_rate']:.1%}" if nsano else "",
        "Data Issues": _nsano_data_issues_label(nsano),
    })

    itc = state.get("itc_result")
    rows.append({
        "Workflow": "ITC/Voda Coll Recon",
        "Status": "Ready" if itc else "Not run",
        "Rows": f"{itc['mambu_count']:,} Mambu / {itc['itc_count']:,} ITC / {itc['vodafone_count']:,} Voda" if itc else "",
        "Match Rate": f"{itc['mvi']['match_rate']:.1%}" if itc else "",
        "Data Issues": "",
    })

    zenith = state.get("zenith_result")
    rows.append({
        "Workflow": "Zenith Coll Recon",
        "Status": "Ready" if zenith else "Not run",
        "Rows": f"{zenith['mambu_count']:,} Mambu / {zenith['zenith_count']:,} Zenith" if zenith else "",
        "Match Rate": f"{zenith['mvz']['match_rate']:.1%}" if zenith else "",
        "Data Issues": "",
    })

    write_off = state.get("write_off_recon_result")
    rows.append({
        "Workflow": "Write-off Recon",
        "Status": "Ready" if write_off else "Not run",
        "Rows": f"{write_off['write_off_count']:,} Write-off" if write_off else "",
        "Match Rate": f"{write_off['stats']['match_rate']:.1%}" if write_off else "",
        "Data Issues": "",
    })

    return rows


def _decimal_amount(value) -> Decimal:
    text = str(value or "").strip().replace(",", "").replace("\t", "")
    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _record_value(record, *headers: str):
    for source in (getattr(record, "sheet_values", {}), getattr(record, "raw", {})):
        for header in headers:
            value = source.get(header, "")
            if value not in ("", None):
                return value
        wanted = {header.casefold() for header in headers}
        for key, value in source.items():
            if str(key).casefold() in wanted and value not in ("", None):
                return value
    return ""


def _record_text(record) -> str:
    values = list(getattr(record, "raw", {}).values()) + list(getattr(record, "sheet_values", {}).values())
    return " ".join(str(value or "") for value in values).casefold()


def _summary(count: int = 0, amount: Decimal | str | int = Decimal("0")) -> dict[str, Decimal | int]:
    return {"count": int(count or 0), "amount": _decimal_amount(amount)}


def _summary_from_stats(stats: dict, count_key: str, amount_key: str) -> dict[str, Decimal | int]:
    if not stats:
        return _summary()
    return _summary(stats.get(count_key, 0), stats.get(amount_key, Decimal("0")))


def _summary_from_status(stats: dict, status: str) -> dict[str, Decimal | int]:
    if not stats:
        return _summary()
    return _summary(
        stats.get("status_counts", {}).get(status, 0),
        stats.get("status_amounts", {}).get(status, Decimal("0")),
    )


def _sum_summaries(*summaries: dict[str, Decimal | int]) -> dict[str, Decimal | int]:
    return _summary(
        sum(int(summary.get("count", 0)) for summary in summaries),
        sum((_decimal_amount(summary.get("amount", Decimal("0"))) for summary in summaries), Decimal("0")),
    )


def _summarize_charge_values(records: list, *amount_headers: str) -> dict[str, Decimal | int]:
    count = 0
    total = Decimal("0")
    for record in records:
        amount = sum((_decimal_amount(_record_value(record, header)) for header in amount_headers), Decimal("0"))
        if amount == 0:
            continue
        count += 1
        total += abs(amount)
    return _summary(count, total)


def _summarize_settlement_records(records: list, *amount_headers: str) -> dict[str, Decimal | int]:
    count = 0
    total = Decimal("0")
    for record in records:
        if "settlement" not in _record_text(record):
            continue
        amount = Decimal("0")
        for header in amount_headers:
            amount = _decimal_amount(_record_value(record, header))
            if amount != 0:
                break
        count += 1
        total += amount
    return _summary(count, total)


def _daily_record_date(value: object) -> str:
    """Return a sortable ISO day for the date formats used by wallet exports."""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value or "").strip()
    if not text:
        return ""
    reconciled_date = normalize_date_amount_match_date(text)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", reconciled_date):
        return reconciled_date
    normalized = text.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(normalized).date().isoformat()
    except ValueError:
        pass
    date_text = re.split(r"[ T]", text, maxsplit=1)[0]
    for date_format in (
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y",
        "%m/%d/%Y", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y",
    ):
        try:
            return datetime.strptime(date_text, date_format).date().isoformat()
        except ValueError:
            continue
    return ""


def _daily_wallet_rows(
    records: list,
    *charge_headers: str,
    vodafone_collection_charges: bool = False,
) -> list[dict[str, object]]:
    """Aggregate wallet transaction amounts and charges without retaining raw rows."""
    daily: dict[str, dict[str, object]] = {}
    charge_detail = VODAFONE_CHARGE_DETAIL.casefold()
    for record in records:
        transaction_date = _daily_record_date(getattr(record, "datetime_value", ""))
        if not transaction_date:
            continue
        bucket = daily.setdefault(
            transaction_date,
            {"date": transaction_date, "amount": Decimal("0"), "charges": Decimal("0")},
        )
        amount = _decimal_amount(getattr(record, "amount_decimal", Decimal("0")))
        bucket["amount"] = _decimal_amount(bucket["amount"]) + amount

        charge = sum(
            (_decimal_amount(_record_value(record, header)) for header in charge_headers),
            Decimal("0"),
        )
        if vodafone_collection_charges:
            details = " ".join(str(_record_value(record, "Details") or "").casefold().split())
            charge = abs(amount) if charge_detail in details else Decimal("0")
        bucket["charges"] = _decimal_amount(bucket["charges"]) + abs(charge)
    return [daily[transaction_date] for transaction_date in sorted(daily)]


def _collection_source_summary_rows(state: dict) -> list[dict]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})

    return [
        {
            "Source": "Nsano",
            "Status": "Completed" if nsano else "Not run",
            "Found In Mambu": "Nsano Collections",
            "Count": _summary_from_stats(nsano.get("nvm", {}), "mambu", "mambu_amount")["count"],
            "Amount": _summary_from_stats(nsano.get("nvm", {}), "mambu", "mambu_amount")["amount"],
        },
        {
            "Source": "ITC",
            "Status": "Completed" if itc else "Not run",
            "Found In Mambu": "ITC Collections",
            "Count": _summary_from_status(itc.get("ivm", {}), ITC_MAMBU_STATUS)["count"],
            "Amount": _summary_from_status(itc.get("ivm", {}), ITC_MAMBU_STATUS)["amount"],
        },
        {
            "Source": "Vodafone",
            "Status": "Completed" if itc else "Not run",
            "Found In Mambu": "Vodafone Collections",
            "Count": _summary_from_status(itc.get("vvm", {}), ITC_MAMBU_STATUS)["count"],
            "Amount": _summary_from_status(itc.get("vvm", {}), ITC_MAMBU_STATUS)["amount"],
        },
        {
            "Source": "Zenith",
            "Status": "Completed" if zenith else "Not run",
            "Found In Mambu": "Zenith Collections",
            "Count": _summary_from_status(zenith.get("zvs", {}), ZENITH_MAMBU_STATUS)["count"],
            "Amount": _summary_from_status(zenith.get("zvs", {}), ZENITH_MAMBU_STATUS)["amount"],
        },
    ]


def _collection_mambu_source_summary_rows(state: dict) -> list[dict]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})

    return [
        {
            "Source": "Mambu",
            "Found In": "Nsano Collections",
            "Count": _summary_from_stats(nsano.get("mvn", {}), "matched", "matched_amount")["count"],
            "Amount": _summary_from_stats(nsano.get("mvn", {}), "matched", "matched_amount")["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "ITC Collections",
            "Count": _summary_from_status(itc.get("mvi", {}), ITC_STATUS)["count"],
            "Amount": _summary_from_status(itc.get("mvi", {}), ITC_STATUS)["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "Vodafone Collections",
            "Count": _summary_from_status(itc.get("mvi", {}), VODA_COLL_STATUS)["count"],
            "Amount": _summary_from_status(itc.get("mvi", {}), VODA_COLL_STATUS)["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "Zenith Collections",
            "Count": _summary_from_status(zenith.get("mvz", {}), ZENITH_MATCHED_STATUS)["count"],
            "Amount": _summary_from_status(zenith.get("mvz", {}), ZENITH_MATCHED_STATUS)["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "Not-Found · Nsano",
            "Count": _summary_from_stats(nsano.get("mvn", {}), "not_found", "not_found_amount")["count"],
            "Amount": _summary_from_stats(nsano.get("mvn", {}), "not_found", "not_found_amount")["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "Not-Found · ITC / Vodafone",
            "Count": _summary_from_status(itc.get("mvi", {}), ITC_NOT_FOUND_STATUS)["count"],
            "Amount": _summary_from_status(itc.get("mvi", {}), ITC_NOT_FOUND_STATUS)["amount"],
        },
        {
            "Source": "Mambu",
            "Found In": "Not-Found · Zenith",
            "Count": _summary_from_status(zenith.get("mvz", {}), ZENITH_NOT_FOUND_STATUS)["count"],
            "Amount": _summary_from_status(zenith.get("mvz", {}), ZENITH_NOT_FOUND_STATUS)["amount"],
        },
    ]


def _collection_exception_summary_rows(state: dict) -> list[dict]:
    return collection_exception_summary_rows(state)


def _disbursement_wallet_source_summary_rows(state: dict) -> list[dict]:
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})

    nsano_summary = _summary_from_status(nsano.get("nvm", {}), DISB_MAMBU_STATUS)
    itc_summary = _summary_from_status(itc.get("ivm", {}), DISB_MAMBU_STATUS)
    mtn_summary = _summary_from_status(mtn.get("mtv", {}), MTN_MANUAL_MAMBU_STATUS)
    vodafone_summary = _summary_from_status(vodafone.get("vtv", {}), VODAFONE_MANUAL_MAMBU_STATUS)
    return [
        {"Source": "Nsano", "Found In Mambu": "Nsano Disbursement", "Count": nsano_summary["count"], "Amount": nsano_summary["amount"]},
        {"Source": "ITC", "Found In Mambu": "ITC Disbursement", "Count": itc_summary["count"], "Amount": itc_summary["amount"]},
        {"Source": "MTN Manual", "Found In Mambu": "MTN Manual Disbursement", "Count": mtn_summary["count"], "Amount": mtn_summary["amount"]},
        {"Source": "Vodafone Manual", "Found In Mambu": "Vodafone Manual Disbursement", "Count": vodafone_summary["count"], "Amount": vodafone_summary["amount"]},
    ]


def _disbursement_mambu_source_summary_rows(state: dict) -> list[dict]:
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})

    nsano_summary = _summary_from_status(nsano.get("mvn", {}), DISB_NSANO_STATUS)
    itc_summary = _summary_from_status(itc.get("mvi", {}), ITC_DISB_STATUS)
    mtn_summary = _summary_from_status(mtn.get("mvm", {}), MTN_MANUAL_MATCHED_STATUS)
    vodafone_summary = _summary_from_status(vodafone.get("mvv", {}), VODAFONE_MANUAL_MATCHED_STATUS)
    return [
        {"Source": "Mambu", "Found In": "Nsano Disbursement", "Count": nsano_summary["count"], "Amount": nsano_summary["amount"]},
        {"Source": "Mambu", "Found In": "ITC Disbursement", "Count": itc_summary["count"], "Amount": itc_summary["amount"]},
        {"Source": "Mambu", "Found In": "MTN Manual Disbursement", "Count": mtn_summary["count"], "Amount": mtn_summary["amount"]},
        {"Source": "Mambu", "Found In": "Vodafone Manual Disbursement", "Count": vodafone_summary["count"], "Amount": vodafone_summary["amount"]},
    ]


def _summary_from_table_rows(rows: list[dict]) -> dict[str, Decimal | int]:
    return _sum_summaries(*(_summary(row.get("Count", 0), row.get("Amount", Decimal("0"))) for row in rows))


def _disbursement_analysis_rows(state: dict) -> list[dict]:
    wallet_total = _summary_from_table_rows(_disbursement_wallet_source_summary_rows(state))
    mambu_total = _summary_from_table_rows(_disbursement_mambu_source_summary_rows(state))
    return [{
        "Analysis": "Wallet - Mambu",
        "Count Difference": int(wallet_total["count"]) - int(mambu_total["count"]),
        "Amount Difference": _decimal_amount(wallet_total["amount"]) - _decimal_amount(mambu_total["amount"]),
    }]


def _disbursement_refund_summary_rows(state: dict) -> list[dict]:
    mtn = state.get("mtn_manual_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn_refund = _summary_from_status(mtn.get("mtv", {}), MTN_MANUAL_REFUND_STATUS)
    vodafone_refund = _summary_from_status(vodafone.get("vtv", {}), VODAFONE_MANUAL_REFUND_STATUS)
    total_refund = _sum_summaries(mtn_refund, vodafone_refund)
    return [
        {"Wallet": "MTN Manual", "Item": "Refund", "Count": mtn_refund["count"], "Amount": mtn_refund["amount"]},
        {"Wallet": "Vodafone Manual", "Item": "Refund", "Count": vodafone_refund["count"], "Amount": vodafone_refund["amount"]},
        {"Wallet": "Total", "Item": "Refund", "Count": total_refund["count"], "Amount": total_refund["amount"]},
    ]


def _disbursement_master_status_rows(state: dict) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    def result_status(result: dict | None) -> str:
        if not result:
            return "Not run"
        if "key_diagnostics" not in result:
            return "Ready"
        diagnostics = result.get("key_diagnostics", {})
        source_keys = int(diagnostics.get("source_key_count", 0))
        mambu_keys = int(diagnostics.get("mambu_key_count", 0))
        overlap = int(diagnostics.get("overlap_count", 0))
        if source_keys and mambu_keys and not overlap:
            return "No key overlap"
        if not source_keys or not mambu_keys:
            return "Missing identifiers"
        return "Ready"

    def key_overlap(result: dict | None) -> str:
        if not result or "key_diagnostics" not in result:
            return ""
        diagnostics = result.get("key_diagnostics", {})
        return (
            f"{int(diagnostics.get('overlap_count', 0)):,} shared / "
            f"{int(diagnostics.get('source_key_count', 0)):,} source / "
            f"{int(diagnostics.get('mambu_key_count', 0)):,} Mambu"
        )

    nsano = state.get("nsano_disb_result")
    rows.append({
        "Workflow": "Nsano Disb Recon",
        "Status": result_status(nsano),
        "Rows": f"{nsano['mambu_count']:,} Mambu / {nsano['nsano_count']:,} Nsano" if nsano else "",
        "Mambu Sheet": nsano.get("mambu_sheet_used", "") if nsano else "",
        "Key Overlap": key_overlap(nsano),
        "Match Rate": f"{nsano['mvn']['match_rate']:.1%}" if nsano else "",
    })

    itc = state.get("itc_disb_result")
    rows.append({
        "Workflow": "ITC Wallet vs Mambu",
        "Status": result_status(itc),
        "Rows": f"{itc['mambu_count']:,} Mambu / {itc['itc_count']:,} ITC" if itc else "",
        "Mambu Sheet": itc.get("mambu_sheet_used", "") if itc else "",
        "Key Overlap": key_overlap(itc),
        "Match Rate": f"{itc['mvi']['match_rate']:.1%}" if itc else "",
    })

    mtn = state.get("mtn_manual_disb_result")
    rows.append({
        "Workflow": "MTN Manual Disb",
        "Status": result_status(mtn),
        "Rows": f"{mtn['mambu_count']:,} Mambu / {mtn['mtn_count']:,} MTN / {mtn['refund_count']:,} Refund" if mtn else "",
        "Mambu Sheet": mtn.get("mambu_sheet_used", "") if mtn else "",
        "Key Overlap": key_overlap(mtn),
        "Match Rate": f"{mtn['mvm']['match_rate']:.1%}" if mtn else "",
    })

    vodafone = state.get("vodafone_manual_disb_result")
    rows.append({
        "Workflow": "Vodafone Manual Disb",
        "Status": result_status(vodafone),
        "Rows": f"{vodafone['mambu_count']:,} Mambu / {vodafone['vodafone_count']:,} Vodafone / {vodafone['refund_count']:,} Refund" if vodafone else "",
        "Mambu Sheet": vodafone.get("mambu_sheet_used", "") if vodafone else "",
        "Key Overlap": key_overlap(vodafone),
        "Match Rate": f"{vodafone['mvv']['match_rate']:.1%}" if vodafone else "",
    })
    return rows


def _master_ledger_status_rows(state: dict) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    nsano_collection = state.get("nsano_collection_ledger_result")
    nsano_collection_metrics = nsano_collection.get("metrics", {}) if isinstance(nsano_collection, dict) else {}
    rows.append({
        "Workflow": "Nsano Collections vs Ledger",
        "Status": "Ready" if nsano_collection else "Not run",
        "Rows": (
            f"{int(nsano_collection_metrics.get('mambu_collection_total', 0)):,} Mambu collections / "
            f"{int(nsano_collection_metrics.get('mambu_disb_total', 0)):,} Mambu disb / "
            f"{int(nsano_collection_metrics.get('nsano_collection_total', 0)):,} Nsano collections / "
            f"{int(nsano_collection_metrics.get('transfer_total', 0)):,} Transfers"
        ) if nsano_collection else "",
        "Ledger Balance": f"{_decimal_amount(nsano_collection_metrics.get('ledger_balance')):,.2f}" if nsano_collection else "",
        "Wallet Statement Balance": f"{_decimal_amount(nsano_collection_metrics.get('wallet_statement_balance')):,.2f}" if nsano_collection else "",
    })

    nsano = state.get("nsano_wallet_ledger_result")
    nsano_metrics = nsano.get("metrics", {}) if isinstance(nsano, dict) else {}
    rows.append({
        "Workflow": "Nsano Disb Wallet vs Ledger",
        "Status": "Ready" if nsano else "Not run",
        "Rows": (
            f"{int(nsano_metrics.get('mambu_total', 0)):,} Mambu / "
            f"{int(nsano_metrics.get('nsano_disb_total', 0)):,} Nsano disb / "
            f"{int(nsano_metrics.get('transfer_total', 0)):,} Transfers"
        ) if nsano else "",
        "Ledger Balance": f"{_decimal_amount(nsano_metrics.get('ledger_balance')):,.2f}" if nsano else "",
        "Wallet Statement Balance": f"{_decimal_amount(nsano_metrics.get('wallet_statement_balance')):,.2f}" if nsano else "",
    })

    itc = state.get("itc_wallet_ledger_result")
    itc_metrics = itc.get("metrics", {}) if isinstance(itc, dict) else {}
    rows.append({
        "Workflow": "ITC Wallet vs Ledger",
        "Status": "Ready" if itc else "Not run",
        "Rows": (
            f"{int(itc_metrics.get('statement_total', 0)):,} Statement / "
            f"{int(itc_metrics.get('debit_total', 0)):,} Debit / "
            f"{int(itc_metrics.get('credit_total', 0)):,} Credit"
        ) if itc else "",
        "Ledger Balance": f"{_decimal_amount(itc_metrics.get('ledger_balance')):,.2f}" if itc else "",
        "Wallet Statement Balance": f"{_decimal_amount(itc_metrics.get('wallet_statement_balance')):,.2f}" if itc else "",
    })
    return rows


def _render_write_off_itc_overlap() -> None:
    wo = st.session_state.get("write_off_recon_result", {})
    itc = st.session_state.get("itc_result", {})
    wo_itc_keys = wo.get("write_off_itc_keys")
    itc_mambu_keys = itc.get("itc_mambu_keys")
    if wo_itc_keys is None or itc_mambu_keys is None:
        return
    overlap = sorted(set(wo_itc_keys) & set(itc_mambu_keys))
    if not overlap:
        return
    st.markdown("---")
    st.warning(
        f"**{len(overlap)} identifier(s)** appear in both Write-off → ITC matches (Write-off Recon) "
        f"and ITC → Mambu matches (ITC Recon). These records are counted differently by each workflow — "
        f"the ITC transaction was matched to Mambu first, but the same key also exists in Write-off."
    )
    st.markdown("**Overlapping identifiers:**")
    for key in overlap:
        st.code(key)


def render_collections_master_tab() -> None:
    st.subheader("Collections Master")
    st.caption(
        "Mambu, Nsano, and ITC default to the saved filtering output for this month if one exists — "
        "upload a fresh file to override. Available workflows run together from this tab."
    )

    drive_config = _google_drive_config()
    filtering_records = _list_filtering_archive_records(drive_config)

    filtered_mambu_picked = _filtered_source_picker(
        label="Filtered Mambu Collection workbook (.xlsx)",
        field_key="collections_master_mambu",
        workflow="Mambu Collection Filtering",
        records=filtering_records,
        file_types=["xlsx"],
    )
    filtered_mambu_workbook = _render_time_value(filtered_mambu_picked)

    col1, col2, col3 = st.columns(3)
    with col1:
        nsano_picked = _filtered_source_picker(
            label="Nsano Data (.xlsx or .csv)",
            field_key="collections_master_nsano",
            workflow="Nsano Disb and Collections Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        nsano_file = _render_time_value(nsano_picked)
        itc_picked = _filtered_source_picker(
            label="ITC Data (.xlsx or .csv)",
            field_key="collections_master_itc",
            workflow="ITC Collection Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        itc_file = _render_time_value(itc_picked)
    with col2:
        zenith_file = st.file_uploader(
            "Zenith Collections Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="collections_master_zenith",
        )
        vodafone_file = st.file_uploader(
            "Raw Vodafone Collections (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="collections_master_vodafone",
        )
    with col3:
        write_off_file = st.file_uploader(
            "Write Off Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="collections_master_write_off",
        )
        unidentified_file = st.file_uploader(
            "Unidentified Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="collections_master_unidentified",
        )

    if isinstance(filtered_mambu_workbook, _ArchivedFilePlaceholder):
        st.caption("Picked from archive — sheet preview available after you run it.")
    else:
        render_collection_master_sheet_status(filtered_mambu_workbook)

    if vodafone_file:
        st.info(
            "Vodafone raw data will be cleaned automatically: rows 1–5 are removed, "
            "row 6 becomes the header, and `Withdrawn` is retained before reconciliation."
        )

    any_source_uploaded = any([
        nsano_file,
        itc_file,
        zenith_file,
        vodafone_file,
        write_off_file,
        unidentified_file,
    ])
    nsano_for_preview = None if isinstance(nsano_file, _ArchivedFilePlaceholder) else nsano_file
    itc_for_preview = None if isinstance(itc_file, _ArchivedFilePlaceholder) else itc_file
    if any_source_uploaded:
        render_column_inspector([
            ("Nsano Collection", nsano_for_preview, NSANO_SHEET_COLUMNS, uploaded_is_csv(nsano_for_preview)),
            ("ITC Collection", itc_for_preview, ITC_SHEET_COLUMNS, uploaded_is_csv(itc_for_preview)),
            ("Zenith Collections", zenith_file, ZENITH_BANK_SHEET_COLUMNS, uploaded_is_csv(zenith_file)),
            ("Write Off", write_off_file, WRITE_OFF_SHEET_COLUMNS, uploaded_is_csv(write_off_file)),
            ("Unidentified", unidentified_file, UNIDENTIFIED_SHEET_COLUMNS, uploaded_is_csv(unidentified_file)),
        ])

    st.markdown("---")
    can_nsano = bool(filtered_mambu_workbook and nsano_file)
    can_itc = bool(filtered_mambu_workbook and (itc_file or vodafone_file))
    can_zenith = bool(filtered_mambu_workbook and zenith_file)
    can_write_off = bool(write_off_file and any([nsano_file, itc_file, zenith_file, vodafone_file]))
    can_run_any = any([can_nsano, can_itc, can_zenith, can_write_off])

    readiness_rows = [
        {
            "Workflow": "Nsano Coll Recon",
            "Required Uploads": "Filtered Mambu + Nsano",
            "Status": "Ready" if can_nsano else "Waiting",
        },
        {
            "Workflow": "ITC/Voda Coll Recon",
            "Required Uploads": "Filtered Mambu + ITC or Vodafone",
            "Status": "Ready" if can_itc else "Waiting",
        },
        {
            "Workflow": "Zenith Coll Recon",
            "Required Uploads": "Filtered Mambu + Zenith",
            "Status": "Ready" if can_zenith else "Waiting",
        },
        {
            "Workflow": "Write-off Recon",
            "Required Uploads": "Write-off + at least one collection source",
            "Status": "Ready" if can_write_off else "Waiting",
        },
    ]
    st.markdown("#### Workflow Readiness")
    st.table(readiness_rows)

    if not can_run_any:
        st.info("Upload the filtered Mambu workbook and at least one matching collection source to enable a master run.")

    run_clicked = st.button(
        "Run Available Collection Reconciliations",
        type="primary",
        disabled=not can_run_any,
        key="run_collections_master",
    )

    if run_clicked:
        for result_key in COLLECTION_RESULT_KEYS:
            st.session_state.pop(result_key, None)
        st.session_state.pop("collections_master_result", None)
        timings: list[dict[str, object]] = []
        master_started = perf_counter()

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            stage_started = perf_counter()
            with st.spinner("Preparing source files (fetching any saved filtered files from the archive)..."):
                filtered_mambu_workbook = _materialize_picked_source(filtered_mambu_picked, drive_config)
                nsano_file = _materialize_picked_source(nsano_picked, drive_config)
                itc_file = _materialize_picked_source(itc_picked, drive_config)
                mambu_path = save_upload(tmp, filtered_mambu_workbook)
                nsano_path = save_upload(tmp, nsano_file)
                itc_path = save_upload(tmp, itc_file)
            zenith_path = save_upload(tmp, zenith_file)
            raw_vodafone_path = save_upload(tmp, vodafone_file)
            vodafone_path = None
            write_off_path = save_upload(tmp, write_off_file)
            unidentified_path = save_upload(tmp, unidentified_file)
            _record_timing(timings, "Save uploaded files", stage_started)
            ran: list[str] = []
            skipped = [
                row["Workflow"]
                for row in readiness_rows
                if row["Status"] != "Ready"
            ]

            try:
                if raw_vodafone_path:
                    stage_started = perf_counter()
                    vodafone_path = tmp / "Voda_Collection_Cleaned.xlsx"
                    build_vodafone_collection_cleanup_workbook(
                        vodafone_path,
                        [raw_vodafone_path],
                    )
                    _record_timing(timings, "Clean Vodafone collection", stage_started)

                if can_nsano:
                    stage_started = perf_counter()
                    with st.spinner("Running Nsano collection reconciliation..."):
                        (
                            mambu_records,
                            mambu_headers,
                            nsano_records,
                            nsano_headers,
                            write_off_records,
                            unidentified_records,
                            mambu_raw_count,
                            nsano_failed_rows_ignored,
                        ) = _timed(
                            timings,
                            "Nsano coll: load sources",
                            load_separate_sources,
                            mambu_path,
                            nsano_path,
                            write_off_path,
                            unidentified_path,
                            mambu_sheet_name=COLLECTION_MASTER_MAMBU_SHEETS["Nsano Coll Recon"],
                        )
                        mambu_vs_nsano_rows = _timed(
                            timings,
                            "Nsano coll: compare Mambu to Nsano",
                            compare_records,
                            mambu_records,
                            nsano_records,
                            "Nsano",
                        )
                        nsano_vs_mambu_rows = _timed(
                            timings,
                            "Nsano coll: compare Nsano to sources",
                            compare_nsano_to_mambu_sources,
                            nsano_records,
                            mambu_records,
                            write_off_records,
                            unidentified_records,
                        )
                        mambu_vs_nsano_stats = _timed(
                            timings,
                            "Nsano coll: summarize Mambu",
                            summarize_compare,
                            mambu_vs_nsano_rows,
                        )
                        nsano_vs_mambu_stats = _timed(
                            timings,
                            "Nsano coll: summarize Nsano",
                            summarize_compare,
                            nsano_vs_mambu_rows,
                        )
                        mambu_vs_nsano_stats["not_found_details"] = not_found_detail_rows(
                            mambu_records,
                            mambu_vs_nsano_rows,
                            NOT_FOUND,
                            "Mambu",
                            "Mambu → Wallet",
                        )
                        nsano_vs_mambu_stats["not_found_details"] = not_found_detail_rows(
                            nsano_records,
                            nsano_vs_mambu_rows,
                            NSANO_NOT_FOUND_STATUS,
                            "Nsano",
                            "Wallet → Mambu",
                        )
                        output_path = tmp / "NSANO Coll Recon.xlsx"
                        summary_rows = build_summary_rows(
                            mambu_records,
                            nsano_records,
                            write_off_records,
                            unidentified_records,
                            mambu_raw_count,
                            mambu_vs_nsano_stats,
                            nsano_vs_mambu_stats,
                            "NSANO Coll Recon.xlsx",
                        )
                        _timed(
                            timings,
                            "Nsano coll: build workbook",
                            write_workbook,
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
                        result_started = perf_counter()
                        st.session_state["nsano_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "nsano_count": len(nsano_records),
                            "write_off_count": len(write_off_records),
                            "unidentified_count": len(unidentified_records),
                            "nsano_failed": nsano_failed_rows_ignored,
                            "mambu_amount": sum_amounts(mambu_records),
                            "nsano_amount": sum_amounts(nsano_records),
                            "charge_summary": _summarize_charge_values(nsano_records, "Charge"),
                            "daily_summary": {"Nsano": _daily_wallet_rows(nsano_records, "Charge")},
                            "settlement_summary": _summarize_settlement_records(nsano_records, "Amount_GHC", "Amount"),
                            "mvn": mambu_vs_nsano_stats,
                            "nvm": nsano_vs_mambu_stats,
                            "collection_diagnostics": _nsano_coll_diagnostics(
                                mambu_records,
                                nsano_records,
                                nsano_vs_mambu_rows,
                                mambu_vs_nsano_stats,
                                nsano_vs_mambu_stats,
                            ),
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, nsano_records, write_off_records, unidentified_records))
                        ran.append("Nsano Coll Recon")
                        _record_timing(timings, "Nsano coll: store result", result_started)
                    _record_timing(timings, "Nsano collection recon", stage_started)

                if can_itc:
                    stage_started = perf_counter()
                    with st.spinner("Running ITC/Voda collection reconciliation..."):
                        (
                            mambu_records,
                            mambu_headers,
                            itc_records,
                            itc_headers,
                            vodafone_records,
                            vodafone_headers,
                            write_off_records,
                            unidentified_records,
                            mambu_raw_count,
                        ) = _timed(
                            timings,
                            "ITC/Voda coll: load sources",
                            load_itc_separate_sources,
                            mambu_path,
                            itc_path,
                            vodafone_path,
                            write_off_path,
                            unidentified_path,
                            mambu_sheet_name=COLLECTION_MASTER_MAMBU_SHEETS["ITC/Voda Coll Recon"],
                        )
                        mambu_vs_itc_rows = _timed(
                            timings,
                            "ITC/Voda coll: compare Mambu",
                            compare_mambu_to_itc_sources,
                            mambu_records,
                            itc_records,
                            vodafone_records,
                        )
                        itc_vs_mambu_rows = _timed(
                            timings,
                            "ITC/Voda coll: compare ITC",
                            compare_itc_to_mambu_sources,
                            itc_records,
                            mambu_records,
                            vodafone_records,
                            unidentified_records,
                            write_off_records,
                        )
                        vodafone_vs_mambu_rows = _timed(
                            timings,
                            "ITC/Voda coll: compare Vodafone",
                            compare_vodafone_to_mambu_sources,
                            vodafone_records,
                            mambu_records,
                            unidentified_records,
                            write_off_records,
                        )
                        mambu_vs_itc_stats = _timed(
                            timings,
                            "ITC/Voda coll: summarize Mambu",
                            summarize_itc_compare,
                            mambu_vs_itc_rows,
                        )
                        itc_vs_mambu_stats = _timed(
                            timings,
                            "ITC/Voda coll: summarize ITC",
                            summarize_itc_compare,
                            itc_vs_mambu_rows,
                        )
                        vodafone_vs_mambu_stats = _timed(
                            timings,
                            "ITC/Voda coll: summarize Vodafone",
                            summarize_itc_compare,
                            vodafone_vs_mambu_rows,
                        )
                        mambu_vs_itc_stats["not_found_details"] = not_found_detail_rows(
                            mambu_records,
                            mambu_vs_itc_rows,
                            ITC_NOT_FOUND_STATUS,
                            "Mambu",
                            "Mambu → Wallet",
                        )
                        itc_vs_mambu_stats["not_found_details"] = not_found_detail_rows(
                            itc_records,
                            itc_vs_mambu_rows,
                            ITC_NOT_FOUND_STATUS,
                            "ITC",
                            "Wallet → Mambu",
                        )
                        vodafone_vs_mambu_stats["not_found_details"] = not_found_detail_rows(
                            vodafone_records,
                            vodafone_vs_mambu_rows,
                            ITC_NOT_FOUND_STATUS,
                            "Vodafone",
                            "Wallet → Mambu",
                        )
                        itc_breakdowns = _timed(
                            timings,
                            "ITC/Voda coll: narration breakdowns",
                            summarize_itc_narration_source_breakdowns,
                            itc_records,
                            itc_vs_mambu_rows,
                        )
                        itc_fee_breakdowns = _timed(
                            timings,
                            "ITC/Voda coll: fee breakdowns",
                            summarize_itc_fee_breakdowns,
                            itc_records,
                        )
                        vodafone_charge_summary = _timed(
                            timings,
                            "ITC/Voda coll: Vodafone charges",
                            summarize_vodafone_charges,
                            vodafone_records,
                        )
                        output_path = tmp / "ITC" / "Voda Coll Recon.xlsx"
                        summary_rows = build_itc_summary_rows(
                            mambu_vs_itc_stats,
                            itc_vs_mambu_stats,
                            vodafone_vs_mambu_stats,
                            itc_breakdowns,
                            itc_fee_breakdowns,
                            vodafone_charge_summary,
                        )
                        _timed(
                            timings,
                            "ITC/Voda coll: build workbook",
                            write_itc_workbook,
                            output_path,
                            summary_rows,
                            mambu_headers,
                            mambu_records,
                            itc_headers,
                            itc_records,
                            vodafone_headers,
                            vodafone_records,
                            write_off_records,
                            unidentified_records,
                            mambu_vs_itc_rows,
                            itc_vs_mambu_rows,
                            vodafone_vs_mambu_rows,
                        )
                        result_started = perf_counter()
                        itc_mambu_keys = sorted({
                            str(row["Source_Key"])
                            for row in itc_vs_mambu_rows
                            if row.get("Match_Status") == ITC_MAMBU_STATUS and row.get("Source_Key")
                        })
                        st.session_state["itc_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "itc_count": len(itc_records),
                            "vodafone_count": len(vodafone_records),
                            "write_off_count": len(write_off_records),
                            "unidentified_count": len(unidentified_records),
                            "mambu_raw_count": mambu_raw_count,
                            "mambu_amount": sum_amounts(mambu_records),
                            "itc_amount": sum_amounts(itc_records),
                            "vodafone_amount": sum_amounts(vodafone_records),
                            "itc_charge_summary": _summarize_charge_values(itc_records, "fees", "elevy_charge"),
                            "itc_settlement_summary": _summarize_settlement_records(itc_records, "amount", "net_amount"),
                            "vodafone_charge_summary": vodafone_charge_summary,
                            "daily_summary": {
                                "ITC": _daily_wallet_rows(itc_records, "fees", "elevy_charge"),
                                "Vodafone": _daily_wallet_rows(
                                    vodafone_records,
                                    vodafone_collection_charges=True,
                                ),
                            },
                            "vodafone_settlement_summary": _summarize_settlement_records(vodafone_records, "Paid In", "Amount", "Amount (GHC)"),
                            "mvi": mambu_vs_itc_stats,
                            "ivm": itc_vs_mambu_stats,
                            "vvm": vodafone_vs_mambu_stats,
                            "itc_mambu_keys": itc_mambu_keys,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, itc_records, vodafone_records, write_off_records, unidentified_records))
                        ran.append("ITC/Voda Coll Recon")
                        _record_timing(timings, "ITC/Voda coll: store result", result_started)
                    _record_timing(timings, "ITC/Voda collection recon", stage_started)

                if can_zenith:
                    stage_started = perf_counter()
                    with st.spinner("Running Zenith collection reconciliation..."):
                        (
                            mambu_records,
                            mambu_headers,
                            zenith_records,
                            zenith_headers,
                            unidentified_records,
                            unidentified_headers,
                            write_off_records,
                            write_off_headers,
                        ) = _timed(
                            timings,
                            "Zenith coll: load sources",
                            load_zenith_collection_sources,
                            mambu_path,
                            zenith_path,
                            unidentified_path,
                            write_off_path,
                            mambu_sheet_name=COLLECTION_MASTER_MAMBU_SHEETS["Zenith Coll Recon"],
                        )
                        if not zenith_records:
                            raise ValueError(
                                "No Zenith transaction rows were found. Upload a Zenith file containing "
                                "Description, Credit/Amount, and Create Date/Date columns."
                        )
                        output_path = tmp / "Zenith Coll Recon.xlsx"
                        mambu_vs_zenith_stats, zenith_vs_sources_stats = _timed(
                            timings,
                            "Zenith coll: build workbook",
                            build_zenith_collection_reconciliation,
                            output_path,
                            mambu_records,
                            mambu_headers,
                            zenith_records,
                            zenith_headers,
                            unidentified_records,
                            unidentified_headers,
                            write_off_records,
                            write_off_headers,
                        )
                        result_started = perf_counter()
                        st.session_state["zenith_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "zenith_count": len(zenith_records),
                            "unidentified_count": len(unidentified_records),
                            "write_off_count": len(write_off_records),
                            "mambu_amount": sum_zenith_amounts(mambu_records),
                            "zenith_amount": sum_zenith_amounts(zenith_records),
                            "charge_summary": _summary(),
                            "settlement_summary": _summarize_settlement_records(zenith_records, "Credit", "Credit (KEY amount)", "Paid In", "Amount"),
                            "daily_summary": {"Zenith": _daily_wallet_rows(zenith_records)},
                            "mvz": mambu_vs_zenith_stats,
                            "zvs": zenith_vs_sources_stats,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, zenith_records, unidentified_records, write_off_records))
                        ran.append("Zenith Coll Recon")
                        _record_timing(timings, "Zenith coll: store result", result_started)
                    _record_timing(timings, "Zenith collection recon", stage_started)

                if can_write_off:
                    stage_started = perf_counter()
                    with st.spinner("Running Write-off reconciliation..."):
                        (
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
                        ) = _timed(
                            timings,
                            "Write-off: load sources",
                            load_write_off_recon_sources,
                            write_off_path,
                            nsano_path,
                            itc_path,
                            zenith_path,
                            vodafone_path,
                        )
                        output_path = tmp / "Write_off_Recon.xlsx"
                        stats, write_off_itc_keys = _timed(
                            timings,
                            "Write-off: build workbook",
                            build_write_off_reconciliation,
                            output_path,
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
                        )
                        result_started = perf_counter()
                        st.session_state["write_off_recon_result"] = _mark_master_archive({
                            "write_off_count": len(write_off_records),
                            "nsano_count": len(nsano_records),
                            "itc_count": len(itc_records),
                            "zenith_count": len(zenith_records),
                            "vodafone_count": len(vodafone_records),
                            "nsano_failed": nsano_failed_rows_ignored,
                            "stats": stats,
                            "write_off_itc_keys": write_off_itc_keys,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(write_off_records, nsano_records, itc_records, zenith_records, vodafone_records))
                        ran.append("Write-off Recon")
                        _record_timing(timings, "Write-off: store result", result_started)
                    _record_timing(timings, "Write-off recon", stage_started)

                st.session_state["collections_master_result"] = {
                    "ran": ran,
                    "skipped": skipped,
                }
                stage_started = perf_counter()
                with st.spinner("Saving outputs to Google Drive archive..."):
                    archive_saved, archive_errors = _archive_master_results(COLLECTION_RESULT_KEYS)
                _record_timing(timings, "Drive archive", stage_started)
                _record_timing(timings, "Total collection master run", master_started)
                st.session_state["collections_master_result"].update({
                    "archive_saved": archive_saved,
                    "archive_errors": archive_errors,
                    "archive_month": _master_archive_month_for_keys(*COLLECTION_RESULT_KEYS),
                    "timings": timings,
                })

            except Exception as exc:
                st.error(f"Collections master run failed: {exc}")
                st.stop()
            finally:
                clear_collection_source_row_cache()

    if not any(key in st.session_state for key in COLLECTION_RESULT_KEYS):
        return

    st.markdown("---")
    st.subheader("Master Results")

    st.markdown("#### Collections Recon - Wallet as Source")
    st.table(_arrow_safe_rows(_collection_source_summary_rows(st.session_state)))

    st.markdown("#### Collections Recon - Mambu as Source")
    st.table(_arrow_safe_rows(_collection_mambu_source_summary_rows(st.session_state)))

    st.markdown("#### Exceptions — Write-off, Charges, Unidentified, and Not Found by Wallet")
    st.table(_arrow_safe_rows(_collection_exception_summary_rows(st.session_state)))

    st.markdown("#### Run Status")
    st.table(_arrow_safe_rows(_collection_master_status_rows(st.session_state)))

    _render_collection_data_diagnostics(st.session_state)

    master = st.session_state.get("collections_master_result", {})
    if master.get("timings"):
        st.markdown("#### Run Timing")
        st.table(_timing_rows(master["timings"]))
    if master.get("archive_errors"):
        st.warning(_archive_warning_text(master["archive_errors"]))
    elif master.get("archive_saved"):
        month = master.get("archive_month") or "the detected data month"
        st.success(f"Saved {master['archive_saved']:,} collection recon file(s) to the `{month}` Drive folder.")
    elif not _drive_archive_enabled():
        st.caption("Drive archive auto-save is off; use the download buttons below for this run.")
    if master.get("skipped"):
        st.caption(f"Skipped: {', '.join(master['skipped'])}")

    _render_write_off_itc_overlap()

    st.markdown("---")
    c1, c2, c3, c4 = st.columns(4)
    if "nsano_result" in st.session_state:
        c1.download_button(
            label="Download Nsano Recon",
            data=st.session_state["nsano_result"]["output_bytes"],
            file_name="NSANO Coll Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_nsano",
        )
    if "itc_result" in st.session_state:
        c2.download_button(
            label="Download ITC/Voda Recon",
            data=st.session_state["itc_result"]["output_bytes"],
            file_name="ITC_Voda_Coll_Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_itc",
        )
    if "zenith_result" in st.session_state:
        c3.download_button(
            label="Download Zenith Recon",
            data=st.session_state["zenith_result"]["output_bytes"],
            file_name="Zenith Coll Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_zenith",
        )
    if "write_off_recon_result" in st.session_state:
        c4.download_button(
            label="Download Write-off Recon",
            data=st.session_state["write_off_recon_result"]["output_bytes"],
            file_name="Write_off_Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_write_off",
        )

    _render_master_summary_download("download_master_summary_sheet", scope="collection")


def render_disbursement_master_tab() -> None:
    st.subheader("Disbursement Master")
    st.caption(
        "Mambu, Nsano, and ITC default to the saved filtering output for this month if one exists — "
        "upload a fresh file to override. Wallet-to-ledger workflows are intentionally excluded."
    )

    drive_config = _google_drive_config()
    filtering_records = _list_filtering_archive_records(drive_config)

    filtered_mambu_picked = _filtered_source_picker(
        label="Filtered Mambu Disbursement workbook (.xlsx)",
        field_key="disbursement_master_mambu",
        workflow="Mambu Disbursement Filtering",
        records=filtering_records,
        file_types=["xlsx"],
    )
    filtered_mambu_workbook = _render_time_value(filtered_mambu_picked)

    found_sheets = (
        detected_workbook_sheets(filtered_mambu_workbook)
        if filtered_mambu_workbook and not isinstance(filtered_mambu_workbook, _ArchivedFilePlaceholder)
        else set()
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        nsano_picked = _filtered_source_picker(
            label="NSANO DISB Data (.xlsx or .csv)",
            field_key="disbursement_master_nsano",
            workflow="Nsano Disb and Collections Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        nsano_file = _render_time_value(nsano_picked)
        itc_picked = _filtered_source_picker(
            label="ITC Wallet Data (.xlsx or .csv)",
            field_key="disbursement_master_itc",
            workflow="ITC Collection Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        itc_file = _render_time_value(itc_picked)
    with col2:
        mtn_manual_file = st.file_uploader(
            "MTN MANUAL Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="disbursement_master_mtn_manual",
        )
        vodafone_manual_file = st.file_uploader(
            "VODAFONE MANUAL Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="disbursement_master_vodafone_manual",
        )
    with col3:
        refund_file = st.file_uploader(
            "REFUND Data (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="disbursement_master_refund",
        )

    if vodafone_manual_file:
        st.info(
            "Vodafone Manual raw data will be cleaned automatically: rows before the `Id` / `Date` / `Amount` "
            "header are removed before matching."
        )

    if isinstance(filtered_mambu_workbook, _ArchivedFilePlaceholder):
        st.caption("Picked from archive — sheet preview available after you run it.")
    else:
        render_disbursement_master_sheet_status(filtered_mambu_workbook)

    any_source_uploaded = any([
        nsano_file,
        itc_file,
        mtn_manual_file,
        vodafone_manual_file,
        refund_file,
    ])
    nsano_for_preview = None if isinstance(nsano_file, _ArchivedFilePlaceholder) else nsano_file
    itc_for_preview = None if isinstance(itc_file, _ArchivedFilePlaceholder) else itc_file
    if any_source_uploaded:
        render_column_inspector([
            ("NSANO DISB", nsano_for_preview, NSANO_DISB_SHEET_COLUMNS, uploaded_is_csv(nsano_for_preview)),
            ("ITC Wallet", itc_for_preview, ITC_DISB_SHEET_COLUMNS, uploaded_is_csv(itc_for_preview)),
            ("MTN MANUAL", mtn_manual_file, MTN_MANUAL_SHEET_COLUMNS, uploaded_is_csv(mtn_manual_file)),
            ("VODAFONE MANUAL", vodafone_manual_file, VODAFONE_MANUAL_SHEET_COLUMNS, uploaded_is_csv(vodafone_manual_file)),
            ("REFUND", refund_file, MTN_MANUAL_REFUND_SHEET_COLUMNS, uploaded_is_csv(refund_file)),
        ])

    st.markdown("---")
    can_nsano = bool(filtered_mambu_workbook and nsano_file)
    can_itc = bool(filtered_mambu_workbook and itc_file)
    can_mtn = bool(filtered_mambu_workbook and mtn_manual_file)
    can_vodafone = bool(filtered_mambu_workbook and vodafone_manual_file)
    can_run_any = any([can_nsano, can_itc, can_mtn, can_vodafone])

    readiness_rows = [
        {
            "Workflow": "Nsano Disb Recon",
            "Required Uploads": "Filtered Mambu + NSANO DISB",
            "Mambu Sheet": DISBURSEMENT_MASTER_MAMBU_SHEETS["Nsano Disb Recon"],
            "Status": "Ready" if can_nsano else "Waiting",
        },
        {
            "Workflow": "ITC Wallet vs Mambu",
            "Required Uploads": "Filtered Mambu + ITC Wallet",
            "Mambu Sheet": DISBURSEMENT_MASTER_MAMBU_SHEETS["ITC Wallet vs Mambu"],
            "Status": "Ready" if can_itc else "Waiting",
        },
        {
            "Workflow": "MTN Manual Disb",
            "Required Uploads": "Filtered Mambu + MTN MANUAL",
            "Mambu Sheet": DISBURSEMENT_MASTER_MAMBU_SHEETS["MTN Manual Disb"],
            "Status": "Ready" if can_mtn else "Waiting",
        },
        {
            "Workflow": "Vodafone Manual Disb",
            "Required Uploads": "Filtered Mambu + VODAFONE MANUAL",
            "Mambu Sheet": DISBURSEMENT_MASTER_MAMBU_SHEETS["Vodafone Manual Disb"],
            "Status": "Ready" if can_vodafone else "Waiting",
        },
    ]
    st.markdown("#### Workflow Readiness")
    st.table(readiness_rows)

    if not can_run_any:
        st.info("Upload the filtered Mambu Disbursement workbook and at least one matching source file to enable a master run.")

    run_clicked = st.button(
        "Run Available Disbursement Reconciliations",
        type="primary",
        disabled=not can_run_any,
        key="run_disbursement_master",
    )

    if run_clicked:
        for result_key in DISBURSEMENT_RESULT_KEYS:
            st.session_state.pop(result_key, None)
        st.session_state.pop("disbursement_master_result", None)
        timings: list[dict[str, object]] = []
        master_started = perf_counter()

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            stage_started = perf_counter()
            with st.spinner("Preparing source files (fetching any saved filtered files from the archive)..."):
                filtered_mambu_workbook = _materialize_picked_source(filtered_mambu_picked, drive_config)
                nsano_file = _materialize_picked_source(nsano_picked, drive_config)
                itc_file = _materialize_picked_source(itc_picked, drive_config)
                found_sheets = detected_workbook_sheets(filtered_mambu_workbook) if filtered_mambu_workbook else found_sheets
                mambu_path = save_upload(tmp, filtered_mambu_workbook)
                nsano_path = save_upload(tmp, nsano_file)
                itc_path = save_upload(tmp, itc_file)
            mtn_manual_path = save_upload(tmp, mtn_manual_file)
            vodafone_manual_path = save_upload(tmp, vodafone_manual_file)
            refund_path = save_upload(tmp, refund_file)
            _record_timing(timings, "Save uploaded files", stage_started)
            ran: list[str] = []
            skipped = [
                row["Workflow"]
                for row in readiness_rows
                if row["Status"] != "Ready"
            ]

            try:
                if can_nsano:
                    stage_started = perf_counter()
                    with st.spinner("Running Nsano disbursement reconciliation..."):
                        mambu_records, mambu_headers, mambu_sheet_used = _timed(
                            timings,
                            "Nsano disb: load Mambu sheet",
                            _load_disbursement_master_mambu_records,
                            mambu_path,
                            "Nsano Disb Recon",
                            found_sheets,
                        )
                        nsano_records, nsano_headers = _timed(
                            timings,
                            "Nsano disb: load Nsano file",
                            load_nsano_disb_file,
                            nsano_path,
                        )
                        key_diagnostics = _timed(
                            timings,
                            "Nsano disb: key diagnostics",
                            _warn_zero_disbursement_overlap,
                            "Nsano Disb Recon",
                            nsano_records,
                            mambu_records,
                            mambu_sheet_used,
                        )
                        output_path = tmp / "Nsano_DISB_Recon.xlsx"
                        mambu_vs_nsano_stats, nsano_vs_mambu_stats = _timed(
                            timings,
                            "Nsano disb: build workbook",
                            build_nsano_disb_reconciliation,
                            output_path,
                            mambu_records,
                            mambu_headers,
                            nsano_records,
                            nsano_headers,
                        )
                        result_started = perf_counter()
                        st.session_state["nsano_disb_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "nsano_count": len(nsano_records),
                            "mambu_amount": sum_disb_amounts(mambu_records),
                            "nsano_amount": sum_disb_amounts(nsano_records),
                            "mvn": mambu_vs_nsano_stats,
                            "nvm": nsano_vs_mambu_stats,
                            "daily_summary": {"Nsano": _daily_wallet_rows(nsano_records, "Charge", "Fees", "Fee")},
                            "mambu_sheet_used": mambu_sheet_used,
                            "key_diagnostics": key_diagnostics,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, nsano_records))
                        ran.append("Nsano Disb Recon")
                        _record_timing(timings, "Nsano disb: store result", result_started)
                    _record_timing(timings, "Nsano disbursement recon", stage_started)

                if can_itc:
                    stage_started = perf_counter()
                    with st.spinner("Running ITC Wallet vs Mambu reconciliation..."):
                        mambu_records, mambu_headers, mambu_sheet_used = _timed(
                            timings,
                            "ITC vs Mambu: load Mambu sheet",
                            _load_disbursement_master_mambu_records,
                            mambu_path,
                            "ITC Wallet vs Mambu",
                            found_sheets,
                        )
                        itc_records, itc_headers = _timed(
                            timings,
                            "ITC vs Mambu: load ITC file",
                            load_itc_disb_file,
                            itc_path,
                        )
                        key_diagnostics = _timed(
                            timings,
                            "ITC vs Mambu: key diagnostics",
                            _warn_zero_disbursement_overlap,
                            "ITC Wallet vs Mambu",
                            itc_records,
                            mambu_records,
                            mambu_sheet_used,
                        )
                        output_path = tmp / "ITC_Wallet_vs_Mambu.xlsx"
                        mambu_vs_itc_stats, itc_vs_mambu_stats = _timed(
                            timings,
                            "ITC vs Mambu: build workbook",
                            build_itc_disb_reconciliation,
                            output_path,
                            mambu_records,
                            mambu_headers,
                            itc_records,
                            itc_headers,
                        )
                        result_started = perf_counter()
                        st.session_state["itc_disb_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "itc_count": len(itc_records),
                            "mambu_amount": sum_disb_amounts(mambu_records),
                            "itc_amount": sum_disb_amounts(itc_records),
                            "mvi": mambu_vs_itc_stats,
                            "ivm": itc_vs_mambu_stats,
                            "daily_summary": {"ITC": _daily_wallet_rows(itc_records, "fees", "fee", "charge")},
                            "mambu_sheet_used": mambu_sheet_used,
                            "key_diagnostics": key_diagnostics,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, itc_records))
                        ran.append("ITC Wallet vs Mambu")
                        _record_timing(timings, "ITC vs Mambu: store result", result_started)
                    _record_timing(timings, "ITC wallet vs Mambu recon", stage_started)

                if can_mtn:
                    stage_started = perf_counter()
                    with st.spinner("Running MTN Manual disbursement reconciliation..."):
                        mambu_records, mambu_headers, mambu_sheet_used = _timed(
                            timings,
                            "MTN manual disb: load Mambu sheet",
                            _load_disbursement_master_mambu_records,
                            mambu_path,
                            "MTN Manual Disb",
                            found_sheets,
                        )
                        mtn_records, mtn_headers = _timed(
                            timings,
                            "MTN manual disb: load MTN file",
                            load_mtn_manual_disb_file,
                            mtn_manual_path,
                        )
                        if refund_path and refund_path.exists():
                            refund_records, refund_headers = _timed(
                                timings,
                                "MTN manual disb: load refund file",
                                load_mtn_refund_file,
                                refund_path,
                            )
                        else:
                            refund_records, refund_headers = [], list(MTN_MANUAL_REFUND_SHEET_COLUMNS)
                        key_diagnostics = _timed(
                            timings,
                            "MTN manual disb: key diagnostics",
                            _warn_zero_disbursement_overlap,
                            "MTN Manual Disb",
                            mtn_records,
                            mambu_records,
                            mambu_sheet_used,
                        )
                        output_path = tmp / "MTN_MANUAL_DISB_Recon.xlsx"
                        mambu_vs_mtn_stats, mtn_vs_sources_stats = _timed(
                            timings,
                            "MTN manual disb: build workbook",
                            build_mtn_manual_disb_reconciliation,
                            output_path,
                            mambu_records,
                            mambu_headers,
                            mtn_records,
                            mtn_headers,
                            refund_records,
                            refund_headers,
                        )
                        result_started = perf_counter()
                        st.session_state["mtn_manual_disb_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "mtn_count": len(mtn_records),
                            "refund_count": len(refund_records),
                            "mambu_amount": sum_disb_amounts(mambu_records),
                            "mtn_amount": sum_disb_amounts(mtn_records),
                            "refund_amount": sum_disb_amounts(refund_records),
                            "mvm": mambu_vs_mtn_stats,
                            "mtv": mtn_vs_sources_stats,
                            "daily_summary": {"MTN": _daily_wallet_rows(mtn_records, "Charge", "Fees", "Fee")},
                            "mambu_sheet_used": mambu_sheet_used,
                            "key_diagnostics": key_diagnostics,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, mtn_records, refund_records))
                        ran.append("MTN Manual Disb")
                        _record_timing(timings, "MTN manual disb: store result", result_started)
                    _record_timing(timings, "MTN manual disbursement recon", stage_started)

                if can_vodafone:
                    stage_started = perf_counter()
                    with st.spinner("Running Vodafone Manual disbursement reconciliation..."):
                        mambu_records, mambu_headers, mambu_sheet_used = _timed(
                            timings,
                            "Vodafone manual disb: load Mambu sheet",
                            _load_disbursement_master_mambu_records,
                            mambu_path,
                            "Vodafone Manual Disb",
                            found_sheets,
                        )
                        vodafone_records, vodafone_headers = _timed(
                            timings,
                            "Vodafone manual disb: load Vodafone file",
                            load_vodafone_manual_disb_file,
                            vodafone_manual_path,
                        )
                        if refund_path and refund_path.exists():
                            refund_records, refund_headers = _timed(
                                timings,
                                "Vodafone manual disb: load refund file",
                                load_mtn_refund_file,
                                refund_path,
                            )
                        else:
                            refund_records, refund_headers = [], list(VODAFONE_MANUAL_REFUND_SHEET_COLUMNS)
                        key_diagnostics = _timed(
                            timings,
                            "Vodafone manual disb: key diagnostics",
                            _warn_zero_disbursement_overlap,
                            "Vodafone Manual Disb",
                            vodafone_records,
                            mambu_records,
                            mambu_sheet_used,
                        )
                        output_path = tmp / "VF_MANUAL_DISB_Recon.xlsx"
                        mambu_vs_vodafone_stats, vodafone_vs_sources_stats = _timed(
                            timings,
                            "Vodafone manual disb: build workbook",
                            build_vodafone_manual_disb_reconciliation,
                            output_path,
                            mambu_records,
                            mambu_headers,
                            vodafone_records,
                            vodafone_headers,
                            refund_records,
                            refund_headers,
                        )
                        result_started = perf_counter()
                        st.session_state["vodafone_manual_disb_result"] = _mark_master_archive({
                            "mambu_count": len(mambu_records),
                            "vodafone_count": len(vodafone_records),
                            "refund_count": len(refund_records),
                            "mambu_amount": sum_disb_amounts(mambu_records),
                            "vodafone_amount": sum_disb_amounts(vodafone_records),
                            "refund_amount": sum_disb_amounts(refund_records),
                            "mvv": mambu_vs_vodafone_stats,
                            "vtv": vodafone_vs_sources_stats,
                            "daily_summary": {"Vodafone": _daily_wallet_rows(vodafone_records, "Charge", "Fees", "Fee")},
                            "mambu_sheet_used": mambu_sheet_used,
                            "key_diagnostics": key_diagnostics,
                            "output_bytes": output_path.read_bytes(),
                        }, _data_month_from_sources(mambu_records, vodafone_records, refund_records))
                        ran.append("Vodafone Manual Disb")
                        _record_timing(timings, "Vodafone manual disb: store result", result_started)
                    _record_timing(timings, "Vodafone manual disbursement recon", stage_started)

                st.session_state["disbursement_master_result"] = {
                    "ran": ran,
                    "skipped": skipped,
                }
                stage_started = perf_counter()
                with st.spinner("Saving outputs to Google Drive archive..."):
                    archive_saved, archive_errors = _archive_master_results(DISBURSEMENT_RESULT_KEYS)
                _record_timing(timings, "Drive archive", stage_started)
                _record_timing(timings, "Total disbursement master run", master_started)
                st.session_state["disbursement_master_result"].update({
                    "archive_saved": archive_saved,
                    "archive_errors": archive_errors,
                    "archive_month": _master_archive_month_for_keys(*DISBURSEMENT_RESULT_KEYS),
                    "timings": timings,
                })

            except Exception as exc:
                st.error(f"Disbursement master run failed: {exc}")
                st.stop()
            finally:
                clear_collection_source_row_cache()
                clear_disbursement_source_row_cache()

    if not any(key in st.session_state for key in DISBURSEMENT_RESULT_KEYS):
        return

    st.markdown("---")
    st.subheader("Master Results")

    st.markdown("#### Disbursement Recon - Wallet/Manual as Source")
    st.table(_arrow_safe_rows(_disbursement_wallet_source_summary_rows(st.session_state)))

    st.markdown("#### Disbursement Recon - Mambu as Source")
    st.table(_arrow_safe_rows(_disbursement_mambu_source_summary_rows(st.session_state)))

    st.markdown("#### Analysis")
    st.table(_arrow_safe_rows(_disbursement_analysis_rows(st.session_state)))

    st.markdown("#### Refund")
    st.table(_arrow_safe_rows(_disbursement_refund_summary_rows(st.session_state)))

    st.markdown("#### Run Status")
    st.table(_arrow_safe_rows(_disbursement_master_status_rows(st.session_state)))

    master = st.session_state.get("disbursement_master_result", {})
    if master.get("timings"):
        st.markdown("#### Run Timing")
        st.table(_timing_rows(master["timings"]))
    if master.get("archive_errors"):
        st.warning(_archive_warning_text(master["archive_errors"]))
    elif master.get("archive_saved"):
        month = master.get("archive_month") or "the detected data month"
        st.success(f"Saved {master['archive_saved']:,} disbursement recon file(s) to the `{month}` Drive folder.")
    elif not _drive_archive_enabled():
        st.caption("Drive archive auto-save is off; use the download buttons below for this run.")
    if master.get("skipped"):
        st.caption(f"Skipped: {', '.join(master['skipped'])}")

    st.markdown("---")
    c1, c2, c3, c4 = st.columns(4)
    if "nsano_disb_result" in st.session_state:
        c1.download_button(
            label="Download Nsano Disb Recon",
            data=st.session_state["nsano_disb_result"]["output_bytes"],
            file_name="Nsano_DISB_Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_nsano_disb",
        )
    if "itc_disb_result" in st.session_state:
        c2.download_button(
            label="Download ITC Wallet Recon",
            data=st.session_state["itc_disb_result"]["output_bytes"],
            file_name="ITC_Wallet_vs_Mambu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_itc_disb",
        )
    if "mtn_manual_disb_result" in st.session_state:
        c3.download_button(
            label="Download MTN Manual Recon",
            data=st.session_state["mtn_manual_disb_result"]["output_bytes"],
            file_name="MTN_MANUAL_DISB_Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_mtn_manual_disb",
        )
    if "vodafone_manual_disb_result" in st.session_state:
        c4.download_button(
            label="Download Vodafone Manual Recon",
            data=st.session_state["vodafone_manual_disb_result"]["output_bytes"],
            file_name="VF_MANUAL_DISB_Recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_master_vodafone_manual_disb",
        )

    _render_master_summary_download("download_disbursement_master_summary_sheet", scope="disbursement")


def _parse_wallet_ledger_amount(text: str, label: str) -> Decimal | None:
    try:
        return Decimal((text or "0").replace(",", "").strip())
    except InvalidOperation:
        st.warning(f"Enter a valid {label} amount.")
        return None


def _render_wallet_ledger_result(
    *,
    result_key: str,
    run_result_key: str,
    workflow: str,
    file_name: str,
    download_key: str,
) -> None:
    result = st.session_state.get(result_key)
    metrics = result.get("metrics") if isinstance(result, dict) else None
    if not isinstance(metrics, Mapping):
        return

    snapshot = {result_key: {"metrics": dict(metrics)}}
    overview_rows = master_ledger_overview_rows(snapshot)
    detail_rows = master_ledger_detail_rows(snapshot)
    status_rows = [
        row
        for row in _master_ledger_status_rows(st.session_state)
        if row.get("Workflow") == workflow
    ]

    st.markdown("---")
    st.subheader(f"{workflow} Results")
    status_col, detail_col = st.columns(2)
    status_col.metric("Workflow Status", "Completed")
    detail_col.metric("Detail Lines", f"{len(detail_rows):,}")

    st.markdown("#### Ledger Overview")
    st.table(_arrow_safe_rows(overview_rows))
    st.markdown("#### Ledger Detail")
    st.table(_arrow_safe_rows(detail_rows))
    st.markdown("#### Run Status")
    st.table(_arrow_safe_rows(status_rows))

    run_result = st.session_state.get(run_result_key, {})
    if run_result.get("archive_errors"):
        st.warning(_archive_warning_text(run_result["archive_errors"]))
    elif run_result.get("archive_saved"):
        month = result.get("archive_month") or "the detected data month"
        st.success(f"Saved {workflow} to the `{month}` Drive folder.")
    elif not _drive_archive_enabled():
        st.caption("Drive archive auto-save is off; use the download button below for this run.")

    st.download_button(
        label=f"Download {workflow}",
        data=result["output_bytes"],
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key=download_key,
    )


def render_nsano_collection_ledger_tab() -> None:
    workflow = "Nsano Collections vs Ledger"
    result_key = "nsano_collection_ledger_result"
    run_result_key = "nsano_collection_ledger_run_result"

    st.subheader(workflow)
    st.caption(
        "Enter the Nsano collections ledger inputs and upload either the combined workbook or all three separate source files. "
        "Mambu Filtered Collection and Mambu Filtered Disbursement default to the saved filtering output for this month if one exists — upload a fresh file to override."
    )

    st.markdown("#### Ledger Inputs")
    field1, field2, field3, field4 = st.columns(4)
    with field1:
        ledger_text = st.text_input(
            "Ledger balance (GHC)",
            value="0.00",
            key="master_ledger_nsano_collection_balance",
        )
    with field2:
        unidentified_text = st.text_input(
            "Delayed transactions (GHC)",
            value="0.00",
            key="master_ledger_nsano_collection_unidentified",
        )
    with field3:
        settlement_text = st.text_input(
            "Recovery from write off (GHC)",
            value="0.00",
            key="master_ledger_nsano_collection_settlement",
        )
    with field4:
        reversal_text = st.text_input(
            "Reversal (GHC)",
            value="0.00",
            key="master_ledger_nsano_collection_reversal",
        )

    ledger_balance = _parse_wallet_ledger_amount(ledger_text, "ledger balance")
    manual_unidentified = _parse_wallet_ledger_amount(unidentified_text, "delayed transactions")
    manual_settlement = _parse_wallet_ledger_amount(settlement_text, "settlement")
    manual_reversal = _parse_wallet_ledger_amount(reversal_text, "reversal")

    st.markdown("#### Source Files")
    combined_workbook = st.file_uploader(
        "Combined Nsano Collections vs Ledger workbook (.xlsx)",
        type=["xlsx"],
        key="master_ledger_nsano_collection_combined",
    )
    st.caption("Combined workbook can contain Nsano Client Collection, Nsano Client Disb, Successful W2A, and Nsano_Transfers sheets.")
    st.markdown("**Or use all three separate source files**")

    drive_config = _google_drive_config()
    filtering_records = _list_filtering_archive_records(drive_config)

    source1, source2, source3 = st.columns(3)
    with source1:
        mambu_collection_picked = _filtered_source_picker(
            label="Mambu Filtered Collection Data",
            field_key="master_ledger_nsano_collection_mambu_collection",
            workflow="Mambu Collection Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        mambu_collection_file = _render_time_value(mambu_collection_picked)
    with source2:
        mambu_disb_picked = _filtered_source_picker(
            label="Mambu Filtered Disbursement Data",
            field_key="master_ledger_nsano_collection_mambu_disb",
            workflow="Mambu Disbursement Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        mambu_disb_file = _render_time_value(mambu_disb_picked)
    with source3:
        transfer_file = st.file_uploader(
            "Nsano Transfers (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="master_ledger_nsano_collection_transfers",
        )

    st.markdown("**Optional: use saved Nsano filtered data to add Charges to Total Debit**")
    nsano_charge_picked = _filtered_source_picker(
        label="Nsano Filtered Data",
        field_key="master_ledger_nsano_collection_charges",
        workflow="Nsano Disb and Collections Filtering",
        records=filtering_records,
        file_types=["xlsx", "csv"],
    )
    nsano_charge_file = _render_time_value(nsano_charge_picked)
    st.caption("Successful W2A export. Column U (Charge) is summed and added to Total Debit as Charges.")

    mambu_collection_for_preview = None if isinstance(mambu_collection_file, _ArchivedFilePlaceholder) else mambu_collection_file
    mambu_disb_for_preview = None if isinstance(mambu_disb_file, _ArchivedFilePlaceholder) else mambu_disb_file
    nsano_charge_for_preview = None if isinstance(nsano_charge_file, _ArchivedFilePlaceholder) else nsano_charge_file
    if any([mambu_collection_for_preview, mambu_disb_for_preview, transfer_file, nsano_charge_for_preview]):
        render_column_inspector([
            ("Mambu Filtered Collection", mambu_collection_for_preview, NSANO_COLLECTION_MAMBU_SHEET_COLUMNS, uploaded_is_csv(mambu_collection_for_preview)),
            ("Mambu Filtered Disbursement", mambu_disb_for_preview, NSANO_DISB_MAMBU_SHEET_COLUMNS, uploaded_is_csv(mambu_disb_for_preview)),
            ("Nsano Transfers", transfer_file, NSANO_WALLET_TRANSFER_SHEET_COLUMNS, uploaded_is_csv(transfer_file)),
            ("Nsano Filtered Data", nsano_charge_for_preview, NSANO_SHEET_COLUMNS, uploaded_is_csv(nsano_charge_for_preview)),
        ])
    if (
        isinstance(mambu_collection_file, _ArchivedFilePlaceholder)
        or isinstance(mambu_disb_file, _ArchivedFilePlaceholder)
        or isinstance(nsano_charge_file, _ArchivedFilePlaceholder)
    ):
        st.caption("Picked from archive — sheet preview available after you run it.")

    can_run = bool(
        ledger_balance is not None
        and manual_unidentified is not None
        and manual_settlement is not None
        and manual_reversal is not None
        and (combined_workbook or (mambu_collection_file and mambu_disb_file and transfer_file))
    )
    st.markdown("#### Workflow Readiness")
    st.table([{
        "Workflow": workflow,
        "Required Uploads": "Combined Nsano workbook or Mambu Filtered Collection + Mambu Filtered Disbursement + Nsano Transfers",
        "Status": "Ready" if can_run else "Waiting",
    }])
    if not can_run:
        st.info("Enter the ledger inputs and upload a combined workbook or all three separate source files.")

    run_clicked = st.button(
        f"Run {workflow}",
        type="primary",
        disabled=not can_run,
        key="run_nsano_collection_ledger",
    )
    if run_clicked:
        st.session_state.pop(result_key, None)
        st.session_state.pop(run_result_key, None)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            combined_path = save_upload(tmp, combined_workbook)
            with st.spinner("Preparing source files (fetching any saved filtered files from the archive)..."):
                mambu_collection_source = _materialize_picked_source(mambu_collection_picked, drive_config)
                mambu_disb_source = _materialize_picked_source(mambu_disb_picked, drive_config)
                nsano_charge_source = _materialize_picked_source(nsano_charge_picked, drive_config)
                mambu_collection_path = save_upload(tmp, mambu_collection_source)
                mambu_disb_path = save_upload(tmp, mambu_disb_source)
                nsano_charge_path = save_upload(tmp, nsano_charge_source)
            transfer_path = save_upload(tmp, transfer_file)
            try:
                import disbursement_reconciliation as disb_reconciliation

                disb_reconciliation = importlib.reload(disb_reconciliation)
                with st.spinner(f"Running {workflow} reconciliation..."):
                    if combined_path:
                        (
                            mambu_collection_records,
                            mambu_collection_headers,
                            mambu_disb_records,
                            mambu_disb_headers,
                            nsano_charge_records,
                            nsano_charge_headers,
                            transfer_rows,
                            transfer_headers,
                        ) = load_nsano_collection_ledger_workbook(combined_path)
                    else:
                        (
                            mambu_collection_records,
                            mambu_collection_headers,
                            mambu_disb_records,
                            mambu_disb_headers,
                            nsano_charge_records,
                            nsano_charge_headers,
                            transfer_rows,
                            transfer_headers,
                        ) = load_nsano_collection_ledger_sources(
                            mambu_collection_path,
                            mambu_disb_path,
                            transfer_path,
                            nsano_charge_path,
                        )

                    output_path = tmp / "Nsano_Collections_vs_Ledger.xlsx"
                    metrics = disb_reconciliation.build_nsano_collection_ledger_reconciliation(
                        output_path,
                        ledger_balance,
                        manual_unidentified,
                        manual_settlement,
                        manual_reversal,
                        mambu_collection_records,
                        mambu_collection_headers,
                        mambu_disb_records,
                        mambu_disb_headers,
                        transfer_rows,
                        transfer_headers,
                        nsano_charge_records,
                        nsano_charge_headers,
                    )
                    st.session_state[result_key] = _mark_master_archive(
                        {
                            "metrics": metrics,
                            "output_bytes": output_path.read_bytes(),
                        },
                        _data_month_from_sources(mambu_collection_records, mambu_disb_records, transfer_rows),
                    )

                with st.spinner(f"Saving {workflow} to Google Drive archive..."):
                    archive_saved, archive_errors = _archive_master_results((result_key,))
                st.session_state[run_result_key] = {
                    "archive_saved": archive_saved,
                    "archive_errors": archive_errors,
                }
            except Exception as exc:
                st.error(f"{workflow} run failed: {exc}")
                return
            finally:
                clear_disbursement_source_row_cache()

    _render_wallet_ledger_result(
        result_key=result_key,
        run_result_key=run_result_key,
        workflow=workflow,
        file_name="Nsano_Collections_vs_Ledger.xlsx",
        download_key="download_nsano_collection_ledger",
    )


def render_nsano_wallet_ledger_tab() -> None:
    workflow = "Nsano Disb Wallet vs Ledger"
    result_key = "nsano_wallet_ledger_result"
    run_result_key = "nsano_wallet_ledger_run_result"

    st.subheader(workflow)
    st.caption(
        "Enter the Nsano disbursement ledger inputs and upload either the combined workbook or all three separate source files. "
        "Filtered Mambu Data and Filtered Nsano Data default to the saved filtering output for this month if one exists — upload a fresh file to override."
    )

    st.markdown("#### Ledger Inputs")
    field1, field2, field3, field4 = st.columns(4)
    with field1:
        ledger_text = st.text_input(
            "Ledger balance (GHC)",
            value="0.00",
            key="master_ledger_nsano_balance",
        )
    with field2:
        unidentified_text = st.text_input(
            "Delayed transactions (GHC)",
            value="0.00",
            key="master_ledger_nsano_unidentified",
        )
    with field3:
        settlement_text = st.text_input(
            "Recovery from write off (GHC)",
            value="0.00",
            key="master_ledger_nsano_settlement",
        )
    with field4:
        reversal_text = st.text_input(
            "Reversal (GHC)",
            value="0.00",
            key="master_ledger_nsano_reversal",
        )

    ledger_balance = _parse_wallet_ledger_amount(ledger_text, "ledger balance")
    manual_unidentified = _parse_wallet_ledger_amount(unidentified_text, "delayed transactions")
    manual_settlement = _parse_wallet_ledger_amount(settlement_text, "settlement")
    manual_reversal = _parse_wallet_ledger_amount(reversal_text, "reversal")

    st.markdown("#### Source Files")
    combined_workbook = st.file_uploader(
        "Combined Nsano Disb Wallet vs Ledger workbook (.xlsx)",
        type=["xlsx"],
        key="master_ledger_nsano_combined",
    )
    st.caption("Combined workbook can contain Nsano_Mambu or Nsano Client Disb, Successful A2W, and Nsano_Transfers sheets.")
    st.markdown("**Or use all three separate source files**")

    drive_config = _google_drive_config()
    filtering_records = _list_filtering_archive_records(drive_config)

    source1, source2, source3 = st.columns(3)
    with source1:
        mambu_picked = _filtered_source_picker(
            label="Filtered Mambu Data",
            field_key="master_ledger_nsano_mambu",
            workflow="Mambu Disbursement Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        mambu_file = _render_time_value(mambu_picked)
    with source2:
        disb_picked = _filtered_source_picker(
            label="Filtered Nsano Data",
            field_key="master_ledger_nsano_disb",
            workflow="Nsano Disb and Collections Filtering",
            records=filtering_records,
            file_types=["xlsx", "csv"],
        )
        disb_file = _render_time_value(disb_picked)
    with source3:
        transfer_file = st.file_uploader(
            "Nsano Transfers (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="master_ledger_nsano_transfers",
        )

    mambu_for_preview = None if isinstance(mambu_file, _ArchivedFilePlaceholder) else mambu_file
    disb_for_preview = None if isinstance(disb_file, _ArchivedFilePlaceholder) else disb_file
    if any([mambu_for_preview, disb_for_preview, transfer_file]):
        render_column_inspector([
            ("Filtered Mambu", mambu_for_preview, NSANO_DISB_MAMBU_SHEET_COLUMNS, uploaded_is_csv(mambu_for_preview)),
            ("Filtered Nsano", disb_for_preview, NSANO_DISB_SHEET_COLUMNS, uploaded_is_csv(disb_for_preview)),
            ("Nsano Transfers", transfer_file, NSANO_WALLET_TRANSFER_SHEET_COLUMNS, uploaded_is_csv(transfer_file)),
        ])
    if isinstance(mambu_file, _ArchivedFilePlaceholder) or isinstance(disb_file, _ArchivedFilePlaceholder):
        st.caption("Picked from archive — sheet preview available after you run it.")

    can_run = bool(
        ledger_balance is not None
        and manual_unidentified is not None
        and manual_settlement is not None
        and manual_reversal is not None
        and (combined_workbook or (mambu_file and disb_file and transfer_file))
    )
    st.markdown("#### Workflow Readiness")
    st.table([{
        "Workflow": workflow,
        "Required Uploads": "Combined Nsano workbook or Filtered Mambu + Filtered Nsano + Nsano Transfers",
        "Status": "Ready" if can_run else "Waiting",
    }])
    if not can_run:
        st.info("Enter the ledger inputs and upload a combined workbook or all three separate source files.")

    run_clicked = st.button(
        f"Run {workflow}",
        type="primary",
        disabled=not can_run,
        key="run_nsano_wallet_ledger",
    )
    if run_clicked:
        st.session_state.pop(result_key, None)
        st.session_state.pop(run_result_key, None)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            combined_path = save_upload(tmp, combined_workbook)
            with st.spinner("Preparing source files (fetching any saved filtered files from the archive)..."):
                mambu_source = _materialize_picked_source(mambu_picked, drive_config)
                disb_source = _materialize_picked_source(disb_picked, drive_config)
                mambu_path = save_upload(tmp, mambu_source)
                disb_path = save_upload(tmp, disb_source)
            transfer_path = save_upload(tmp, transfer_file)
            try:
                import disbursement_reconciliation as disb_reconciliation

                disb_reconciliation = importlib.reload(disb_reconciliation)
                with st.spinner(f"Running {workflow} reconciliation..."):
                    if combined_path:
                        (
                            mambu_records,
                            mambu_headers,
                            disb_records,
                            disb_headers,
                            transfer_rows,
                            transfer_headers,
                        ) = load_nsano_wallet_ledger_workbook(combined_path)
                    else:
                        (
                            mambu_records,
                            mambu_headers,
                            disb_records,
                            disb_headers,
                            transfer_rows,
                            transfer_headers,
                        ) = load_nsano_wallet_ledger_sources(mambu_path, disb_path, transfer_path)

                    output_path = tmp / "Nsano_Disb_Wallet_vs_Ledger.xlsx"
                    metrics = disb_reconciliation.build_nsano_wallet_ledger_reconciliation(
                        output_path,
                        ledger_balance,
                        manual_unidentified,
                        manual_settlement,
                        manual_reversal,
                        mambu_records,
                        mambu_headers,
                        disb_records,
                        disb_headers,
                        transfer_rows,
                        transfer_headers,
                    )
                    st.session_state[result_key] = _mark_master_archive(
                        {
                            "metrics": metrics,
                            "output_bytes": output_path.read_bytes(),
                        },
                        _data_month_from_sources(mambu_records, disb_records, transfer_rows),
                    )

                with st.spinner(f"Saving {workflow} to Google Drive archive..."):
                    archive_saved, archive_errors = _archive_master_results((result_key,))
                st.session_state[run_result_key] = {
                    "archive_saved": archive_saved,
                    "archive_errors": archive_errors,
                }
            except Exception as exc:
                st.error(f"{workflow} run failed: {exc}")
                return
            finally:
                clear_disbursement_source_row_cache()

    _render_wallet_ledger_result(
        result_key=result_key,
        run_result_key=run_result_key,
        workflow=workflow,
        file_name="Nsano_Disb_Wallet_vs_Ledger.xlsx",
        download_key="download_nsano_wallet_ledger",
    )


def render_itc_wallet_ledger_tab() -> None:
    workflow = "ITC Wallet vs Ledger"
    result_key = "itc_wallet_ledger_result"
    run_result_key = "itc_wallet_ledger_run_result"

    st.subheader(workflow)
    st.caption(
        "ITC Statement defaults to the saved ITC Collection Filtering output (Outflow) for this month if one "
        "exists — upload a fresh file to override. Enter the ledger inputs and upload the Debit and Credit transfers."
    )

    st.markdown("#### Ledger Inputs")
    field1, field2 = st.columns(2)
    with field1:
        ledger_text = st.text_input(
            "Ledger balance (GHC)",
            value="0.00",
            key="master_ledger_itc_balance",
        )
    with field2:
        unidentified_text = st.text_input(
            "Delayed transactions (GHC)",
            value="0.00",
            key="master_ledger_itc_unidentified",
        )

    ledger_balance = _parse_wallet_ledger_amount(ledger_text, "ledger balance")
    manual_unidentified = _parse_wallet_ledger_amount(unidentified_text, "delayed transactions")

    st.markdown("#### Source Files")
    drive_config = _google_drive_config()
    filtering_records = _list_filtering_archive_records(drive_config)

    statement_picked = _filtered_source_picker(
        label="ITC Statement (.xlsx or .csv)",
        field_key="master_ledger_itc_statement",
        workflow="ITC Collection Filtering",
        records=filtering_records,
        file_types=["xlsx", "csv"],
    )
    statement_file = _render_time_value(statement_picked)

    source2, source3 = st.columns(2)
    with source2:
        debit_file = st.file_uploader(
            "ITC Debit Transfers (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="master_ledger_itc_debit_transfers",
        )
    with source3:
        credit_file = st.file_uploader(
            "ITC Credit Transfers (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="master_ledger_itc_credit_transfers",
        )

    statement_for_preview = None if isinstance(statement_file, _ArchivedFilePlaceholder) else statement_file
    if isinstance(statement_file, _ArchivedFilePlaceholder):
        st.caption("ITC Statement picked from archive — sheet preview available after you run it.")
    if any([statement_for_preview, debit_file, credit_file]):
        render_column_inspector([
            ("ITC Statement", statement_for_preview, ITC_WALLET_STATEMENT_SHEET_COLUMNS, uploaded_is_csv(statement_for_preview)),
            ("ITC Debit Transfers", debit_file, ITC_WALLET_DEBIT_TRANSFER_SHEET_COLUMNS, uploaded_is_csv(debit_file)),
            ("ITC Credit Transfers", credit_file, ITC_WALLET_CREDIT_TRANSFER_SHEET_COLUMNS, uploaded_is_csv(credit_file)),
        ])

    can_run = bool(
        ledger_balance is not None
        and manual_unidentified is not None
        and statement_file
        and debit_file
        and credit_file
    )
    st.markdown("#### Workflow Readiness")
    st.table([{
        "Workflow": workflow,
        "Required Uploads": "ITC Statement + Debit Transfers + Credit Transfers",
        "Status": "Ready" if can_run else "Waiting",
    }])
    if not can_run:
        st.info("Enter the ledger inputs and provide the ITC Statement, Debit Transfers, and Credit Transfers.")

    run_clicked = st.button(
        f"Run {workflow}",
        type="primary",
        disabled=not can_run,
        key="run_itc_wallet_ledger",
    )
    if run_clicked:
        st.session_state.pop(result_key, None)
        st.session_state.pop(run_result_key, None)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            with st.spinner("Preparing source files (fetching any saved filtered files from the archive)..."):
                statement_source = _materialize_picked_source(statement_picked, drive_config)
                statement_path = save_upload(tmp, statement_source)
            debit_path = save_upload(tmp, debit_file)
            credit_path = save_upload(tmp, credit_file)
            try:
                import disbursement_reconciliation as disb_reconciliation

                disb_reconciliation = importlib.reload(disb_reconciliation)
                with st.spinner(f"Running {workflow} reconciliation..."):
                    (
                        statement_rows,
                        statement_headers,
                        debit_rows,
                        debit_headers,
                        credit_rows,
                        credit_headers,
                    ) = load_itc_wallet_ledger_sources(statement_path, debit_path, credit_path)

                    output_path = tmp / "ITC_Wallet_vs_Ledger.xlsx"
                    metrics = disb_reconciliation.build_itc_wallet_ledger_reconciliation(
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
                    st.session_state[result_key] = _mark_master_archive(
                        {
                            "metrics": metrics,
                            "output_bytes": output_path.read_bytes(),
                        },
                        _data_month_from_sources(statement_rows, debit_rows, credit_rows, metrics),
                    )

                with st.spinner(f"Saving {workflow} to Google Drive archive..."):
                    archive_saved, archive_errors = _archive_master_results((result_key,))
                st.session_state[run_result_key] = {
                    "archive_saved": archive_saved,
                    "archive_errors": archive_errors,
                }
            except Exception as exc:
                st.error(f"{workflow} run failed: {exc}")
                return
            finally:
                clear_disbursement_source_row_cache()

    _render_wallet_ledger_result(
        result_key=result_key,
        run_result_key=run_result_key,
        workflow=workflow,
        file_name="ITC_Wallet_vs_Ledger.xlsx",
        download_key="download_itc_wallet_ledger",
    )


def _render_filtering_automation(automation_options: list[str], radio_key: str) -> None:
    automation = st.radio(
        "Filtering Automation",
        automation_options,
        horizontal=True,
        key=radio_key,
    )
    if automation == "Nsano Disb and Collections":
        title = "Nsano Disb and Collections Filtering"
        file_label = "Nsano data file(s) (.xlsx or .csv)"
        uploader_key = "nsano_collection_filter_chunks"
        run_key = "run_nsano_collection_filtering"
        result_key = "nsano_collection_filtering_result"
        output_name = "Nsano_Disb_Collections_Filtered.xlsx"
        download_label = "⬇️ Download Filtered Nsano File"
        filters = NSANO_COLLECTION_FILTERS
        builder = build_nsano_collection_filter_workbook
    elif automation == "Mambu Disbursement":
        title = "Mambu Disbursement Filtering"
        file_label = "Mambu disbursement file(s) (.xlsx or .csv)"
        uploader_key = "mambu_disbursement_filter_chunks"
        run_key = "run_mambu_disbursement_filtering"
        result_key = "mambu_disbursement_filtering_result"
        output_name = "Mambu_Disbursement_Filtered.xlsx"
        download_label = "⬇️ Download Filtered Mambu Disbursement File"
        filters = MAMBU_DISBURSEMENT_FILTERS
        builder = build_mambu_disbursement_filter_workbook
    elif automation == "Voda Collection Cleanup":
        title = "Voda Collection Cleanup"
        file_label = "Vodafone collection file(s) (.xlsx or .csv)"
        uploader_key = "voda_collection_cleanup_chunks"
        run_key = "run_voda_collection_cleanup"
        result_key = "voda_collection_cleanup_result"
        output_name = "Voda_Collection_Cleaned.xlsx"
        download_label = "⬇️ Download Cleaned Vodafone Collection File"
        filters = VODAFONE_COLLECTION_CLEANUP_FILTERS
        builder = build_vodafone_collection_cleanup_workbook
    elif automation == "Voda vs Ledger":
        title = "Voda vs Ledger"
        file_label = "Vodafone collection file(s) (.xlsx or .csv)"
        uploader_key = "voda_wallet_ledger_chunks"
        run_key = "run_voda_wallet_ledger"
        result_key = "voda_wallet_ledger_result"
        output_name = "Voda_Wallet_vs_Ledger.xlsx"
        download_label = "⬇️ Download Voda Wallet vs Ledger File"
        filters = VODAFONE_COLLECTION_CLEANUP_FILTERS
        builder = build_vodafone_wallet_ledger_workbook
    elif automation == "Zenith Wallet vs Ledger":
        title = "Zenith Wallet vs Ledger"
        file_label = "Zenith wallet file(s) (.xlsx or .csv)"
        uploader_key = "zenith_wallet_ledger_chunks"
        run_key = "run_zenith_wallet_ledger"
        result_key = "zenith_wallet_ledger_result"
        output_name = "Zenith_Wallet_vs_Ledger.xlsx"
        download_label = "⬇️ Download Zenith Wallet vs Ledger File"
        filters = []
        builder = build_zenith_wallet_ledger_workbook
    elif automation == "ITC Collection":
        title = "ITC Collection Filtering"
        file_label = "ITC collection file(s) (.xlsx or .csv)"
        uploader_key = "itc_collection_filter_chunks"
        run_key = "run_itc_collection_filtering"
        result_key = "itc_collection_filtering_result"
        output_name = "ITC_Collection_Filtered.xlsx"
        download_label = "⬇️ Download Filtered ITC Collection File"
        filters = ITC_COLLECTION_FILTERS
        builder = build_itc_collection_filter_workbook
    else:
        title = "Mambu Collection Filtering"
        file_label = "Mambu collection file(s) (.xlsx or .csv)"
        uploader_key = "mambu_collection_filter_chunks"
        run_key = "run_mambu_collection_filtering"
        result_key = "mambu_collection_filtering_result"
        output_name = "Mambu_Collection_Filtered.xlsx"
        download_label = "⬇️ Download Filtered Mambu Collection File"
        filters = MAMBU_COLLECTION_FILTERS
        builder = build_mambu_collection_filter_workbook

    st.subheader(title)
    st.caption("Upload one or more collection files. Files can be .xlsx or .csv and will be combined before filtering.")

    is_voda_ledger = automation == "Voda vs Ledger"
    is_zenith_ledger = automation == "Zenith Wallet vs Ledger"
    is_wallet_ledger = is_voda_ledger or is_zenith_ledger
    ledger_balance = None
    delayed_transactions = None
    if is_wallet_ledger:
        balance_col, delayed_col = st.columns(2)
        with balance_col:
            balance_text = st.text_input(
                "Balance Per Ledger (GHC)" if is_zenith_ledger else "Balance (GHC)",
                value="0.00",
                key="zenith_wallet_ledger_balance" if is_zenith_ledger else "voda_wallet_ledger_balance",
            )
        with delayed_col:
            delayed_text = st.text_input(
                "Delayed Transactions ( GHC )",
                value="0.00",
                key="zenith_wallet_ledger_delayed" if is_zenith_ledger else "voda_wallet_ledger_delayed",
            )
        try:
            ledger_balance = Decimal((balance_text or "0").replace(",", "").strip())
        except InvalidOperation:
            st.warning("Enter a valid balance amount.")
        try:
            delayed_transactions = Decimal((delayed_text or "0").replace(",", "").strip())
        except InvalidOperation:
            st.warning("Enter a valid delayed transactions amount.")

    uploaded_files = st.file_uploader(
        file_label,
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=uploader_key,
    )
    upload_count = len(uploaded_files or [])

    if upload_count:
        st.caption(f"{upload_count} file{'s' if upload_count != 1 else ''} selected.")

    st.markdown("---")
    can_run = upload_count > 0 and (
        not is_wallet_ledger
        or (ledger_balance is not None and delayed_transactions is not None)
    )

    if not can_run:
        if is_wallet_ledger:
            st.info("Enter the ledger inputs and upload at least one wallet file to enable the ledger run.")
        else:
            st.info("Upload at least one collection file to enable filtering.")

    run_clicked = st.button(
        f"Run {title}",
        type="primary",
        disabled=not can_run,
        key=run_key,
    )

    if run_clicked:
        st.session_state.pop(result_key, None)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            source_paths = save_uploads(tmp, uploaded_files)

            try:
                with st.spinner("Combining chunks and filtering collection transactions..."):
                    output_path = tmp / output_name
                    if is_wallet_ledger:
                        result = builder(output_path, source_paths, ledger_balance, delayed_transactions)
                    else:
                        result = builder(output_path, source_paths)
                    output_bytes = output_path.read_bytes()
                    data_month = _filtering_data_month(source_paths)

                st.session_state[result_key] = _mark_master_archive({
                    "input_rows": result.input_rows,
                    "filtered_rows": result.filtered_rows,
                    "filter_counts": result.filter_counts,
                    "filter_amounts": result.filter_amounts,
                    "metrics": result.metrics or {},
                    "output_bytes": output_bytes,
                }, data_month)

                with st.spinner(f"Saving {title} to the {_month_label(data_month or _current_archive_month())} Drive archive folder..."):
                    _saved_count, archive_errors = _archive_master_results((result_key,))
                if archive_errors:
                    st.warning(_archive_warning_text(archive_errors))

            except Exception as exc:
                st.error(f"{title} failed: {exc}")
                st.stop()

    if result_key not in st.session_state:
        return

    r = st.session_state[result_key]

    st.markdown("---")
    st.subheader("Results")

    c1, c2 = st.columns(2)
    c1.metric("Combined Input Rows", f"{r['input_rows']:,}")
    c2.metric("All Filtered Rows", f"{r['filtered_rows']:,}")

    metrics = r.get("metrics", {})
    if is_wallet_ledger and metrics:
        st.markdown("#### Ledger Summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Balance Per Ledger" if is_zenith_ledger else "Balance", f"{metrics['balance']:,.2f}")
        m2.metric("Total Collections", f"{metrics['total_collections']:,.2f}")
        m3.metric("Available Funds Before Debit", f"{metrics['available_funds']:,.2f}")
        m4.metric("Balance as per Wallet Statement", f"{metrics['wallet_statement_balance']:,.2f}")

        if is_voda_ledger:
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Delayed Transactions", f"{metrics['delayed_transactions']:,.2f}")
            d2.metric("Total Charges", f"{metrics['total_charges']:,.2f}")
            d3.metric("Transfer to Bank", f"{metrics['transfer_to_bank']:,.2f}")
            d4.metric("Total Debit", f"{metrics['total_debit']:,.2f}")
        else:
            d1, d2, d3 = st.columns(3)
            d1.metric("Delayed Transactions", f"{metrics['delayed_transactions']:,.2f}")
            d2.metric("Total Debit", f"{metrics['total_debit']:,.2f}")
            d3.metric("Debit Rows", f"{metrics.get('debit_count', 0):,}")

    rows = []
    filter_names = [rule.name for rule in filters] or list(r["filter_counts"].keys())
    for filter_name in filter_names:
        amount = r["filter_amounts"].get(filter_name, Decimal("0"))
        rows.append({
            "Filter": filter_name,
            "Rows": f"{r['filter_counts'].get(filter_name, 0):,}",
            "Amount": f"{amount:,.2f}",
        })
    st.table(rows)

    st.markdown("---")
    st.download_button(
        label=download_label,
        data=r["output_bytes"],
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key=f"download_{result_key}",
    )


def render_collection_filtering_tab() -> None:
    _render_filtering_automation(
        [
            "Mambu Collection",
            "ITC Collection",
            "Nsano Disb and Collections",
            "Mambu Disbursement",
            "Voda Collection Cleanup",
        ],
        radio_key="collection_filtering_automation",
    )


def render_wallet_ledger_filtering_tab() -> None:
    _render_filtering_automation(
        [
            "Voda vs Ledger",
            "Zenith Wallet vs Ledger",
        ],
        radio_key="wallet_ledger_filtering_automation",
    )


def render_recon_archive_tab() -> None:
    drive_config = _google_drive_config()
    drive_connected, drive_message = check_drive_connection(drive_config)
    drive_setup_blocked = bool(drive_config and _drive_storage_setup_blocked(drive_message))
    if drive_connected:
        with st.spinner("Saving generated outputs to Google Drive archive..."):
            _archive_available_outputs(drive_config)
        records, list_error = list_archive_records(drive_config)
    else:
        records, list_error = [], None

    st.subheader("Recon Archive")
    st.caption("Saved reconciliation, wallet ledger, and monthly summary workbooks.")

    current_count = sum(1 for record in records if not _archive_record_is_previous(record))
    previous_count = len(records) - current_count
    status_col, month_col, storage_col, drive_col = st.columns(4)
    status_col.metric("Current Files", f"{current_count:,}")
    month_col.metric("Months", f"{len({record.get('month', '') for record in records if record.get('month')}):,}")
    storage_col.metric("Storage", _format_bytes(sum(int(record.get("size_bytes") or 0) for record in records)))
    drive_col.metric("Previous Versions", f"{previous_count:,}")

    status_msg_col, status_btn_col = st.columns([3, 2])
    with status_msg_col:
        if not _drive_archive_enabled():
            st.info("Drive archive auto-save is off. Existing archived files can still be viewed when Drive is connected.")
        if drive_connected:
            st.success(drive_message)
        elif drive_setup_blocked:
            st.warning(_drive_storage_setup_message())
        elif drive_config:
            st.warning(f"Google Drive archive is unavailable: {drive_message}")
        else:
            st.error("Google Drive is required for the reconciliation archive but is not configured.")
        if list_error and not _drive_storage_setup_blocked(list_error):
            st.caption(f"Drive index sync skipped: {list_error}")
    with status_btn_col:
        refresh_col, rebuild_col = st.columns(2)
        with refresh_col:
            if drive_connected and st.button("Refresh from Drive", key="archive_refresh_from_drive"):
                with st.spinner("Checking archive against Google Drive..."):
                    prune_result = prune_missing_drive_records(drive_config)
                if prune_result.removed:
                    st.success(f"Removed {prune_result.removed:,} record(s) no longer in Drive.")
                else:
                    st.success("Archive is up to date with Drive.")
                if prune_result.errors:
                    st.warning("Some records could not be checked: " + "; ".join(prune_result.errors[:3]))
                st.rerun()
        with rebuild_col:
            if drive_connected and st.button(
                "Rebuild Archive Index",
                key="archive_rebuild_drive_index",
                help="Repopulate the archive index from workbooks already stored in the Drive archive folders.",
            ):
                with st.spinner("Scanning Drive and rebuilding the archive index..."):
                    rebuild_result = rebuild_drive_index(
                        drive_config,
                        list(ARCHIVABLE_RESULTS.values()),
                    )
                if rebuild_result.recovered:
                    st.success(f"Recovered {rebuild_result.recovered:,} archived file(s) from Drive.")
                if rebuild_result.skipped:
                    st.warning(
                        f"Skipped {rebuild_result.skipped:,} file(s) whose names did not match a known app output."
                    )
                if rebuild_result.errors:
                    st.warning("; ".join(rebuild_result.errors[:3]))
                if rebuild_result.recovered:
                    st.rerun()

    if not records:
        if drive_setup_blocked:
            st.info(
                "No Drive archive can be shown until the archive folder is a Shared Drive folder "
                "or a delegated user is configured."
            )
        else:
            st.info("No archived files yet. Run a reconciliation, wallet ledger, or monthly summary to populate this dashboard.")
        return

    st.markdown("---")
    filter_col1, filter_col2, filter_col3, filter_col4, filter_col5 = st.columns([1.1, 1.1, 1.4, 1.1, 1.4])
    months = ["All"] + sorted({str(record.get("month", "")) for record in records if record.get("month")}, reverse=True)
    categories = ["All"] + sorted({str(record.get("category", "")) for record in records if record.get("category")})
    workflows = ["All"] + sorted({str(record.get("workflow", "")) for record in records if record.get("workflow")})
    with filter_col1:
        selected_month = st.selectbox(
            "Month",
            months,
            format_func=lambda value: "All Months" if value == "All" else _month_label(value),
            key="archive_filter_month",
        )
    with filter_col2:
        selected_category = st.selectbox("Type", categories, key="archive_filter_category")
    with filter_col3:
        selected_workflow = st.selectbox("Workflow", workflows, key="archive_filter_workflow")
    with filter_col4:
        show_previous_versions = st.checkbox("Show previous", value=False, key="archive_show_previous_versions")
    with filter_col5:
        search_text = st.text_input("Search", key="archive_filter_search").strip().lower()

    filtered_records = []
    for record in records:
        haystack = " ".join(
            str(record.get(field, ""))
            for field in ("month", "category", "workflow", "file_name", "created_at")
        ).lower()
        if not show_previous_versions and _archive_record_is_previous(record):
            continue
        if selected_month != "All" and record.get("month") != selected_month:
            continue
        if selected_category != "All" and record.get("category") != selected_category:
            continue
        if selected_workflow != "All" and record.get("workflow") != selected_workflow:
            continue
        if search_text and search_text not in haystack:
            continue
        filtered_records.append(record)

    st.markdown("#### Archived Recons")
    st.dataframe(
        _arrow_safe_rows([
            {
                "Version": _archive_record_version_label(record),
                "Created": str(record.get("created_at", ""))[:19].replace("T", " "),
                "Month": _month_label(str(record.get("month", ""))),
                "Type": record.get("category", ""),
                "Workflow": record.get("workflow", ""),
                "File": record.get("file_name", ""),
                "Size": _format_bytes(record.get("size_bytes", 0)),
                "Drive": "Saved" if record.get("drive_saved") else "Local",
            }
            for record in filtered_records
        ]),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### Month Actions")
    if selected_month == "All":
        st.caption("Choose a specific month above to download or delete all archive files for that month.")
    else:
        month_records = [record for record in records if record.get("month") == selected_month]
        month_file_count = len(month_records)
        month_storage = _format_bytes(sum(int(record.get("size_bytes") or 0) for record in month_records))
        st.caption(
            f"Download or delete all current and previous archive files for {_month_label(selected_month)} "
            f"({month_file_count:,} file{'s' if month_file_count != 1 else ''}, {month_storage})."
        )
        month_signature = repr([record.get("id") for record in month_records])
        month_download_col, month_delete_col = st.columns([1, 1])
        with month_download_col:
            if st.button(
                "Prepare Selected Month ZIP",
                disabled=not month_records,
                key="archive_prepare_month_zip",
            ):
                with st.spinner("Preparing month archive ZIP..."):
                    zip_bytes, skipped = _archive_zip_bytes(month_records, drive_config)
                st.session_state["archive_month_zip_result"] = {
                    "month": selected_month,
                    "signature": month_signature,
                    "zip_bytes": zip_bytes,
                    "skipped": skipped,
                    "file_name": f"Recon_Archive_{selected_month}",
                }
            month_zip_result = st.session_state.get("archive_month_zip_result")
            if (
                month_zip_result
                and month_zip_result.get("month") == selected_month
                and month_zip_result.get("signature") == month_signature
            ):
                if month_zip_result.get("skipped"):
                    st.caption(f"Skipped unavailable files: {', '.join(month_zip_result['skipped'])}")
                st.download_button(
                    label="Download Selected Month ZIP",
                    data=month_zip_result["zip_bytes"],
                    file_name=f"{month_zip_result['file_name']}.zip",
                    mime="application/zip",
                    key="archive_download_month_zip",
                )
        with month_delete_col:
            confirm_delete_month = st.checkbox(
                f"Confirm delete {_month_label(selected_month)} archive data",
                key="archive_confirm_delete_month",
            )
            if st.button(
                "Delete Selected Month",
                disabled=not confirm_delete_month or not month_records,
                key="archive_delete_month",
            ):
                delete_result = delete_archive_records(month_records, drive_config)
                if delete_result.errors:
                    st.warning(
                        f"Deleted {delete_result.deleted:,} file(s), but some cleanup failed: "
                        + "; ".join(delete_result.errors[:3])
                    )
                else:
                    st.success(f"Deleted {delete_result.deleted:,} archive file(s) for {_month_label(selected_month)}.")
                st.rerun()

    if not filtered_records:
        st.info("No archive records match the current filters.")
        return

    st.markdown("---")
    action_col1, action_col2 = st.columns([1.5, 1])
    selected_record = None
    with action_col1:
        selected_label = st.selectbox(
            "Selected file",
            [_archive_record_label(record) for record in filtered_records],
            key="archive_selected_record",
        )
        selected_record = next(
            record for record in filtered_records if _archive_record_label(record) == selected_label
        )

    with action_col2:
        st.metric("Visible Files", f"{len(filtered_records):,}")
        st.metric("Visible Storage", _format_bytes(sum(int(record.get("size_bytes") or 0) for record in filtered_records)))

    download_col, delete_col = st.columns([3, 1])
    with download_col:
        try:
            selected_bytes = read_archive_file(selected_record, drive_config)
        except Exception as exc:
            st.error(f"Selected file is not available: {exc}")
        else:
            st.download_button(
                label="Download Selected Recon",
                data=selected_bytes,
                file_name=str(selected_record.get("file_name") or "Recon.xlsx"),
                mime=str(selected_record.get("mime_type") or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                type="primary",
                key="archive_download_selected",
            )
    with delete_col:
        if st.button("Delete Selected", key="archive_delete_selected"):
            delete_error = delete_archive_record(selected_record, drive_config)
            if delete_error:
                st.warning(f"Delete failed or only partially completed: {delete_error}")
            else:
                st.success("Deleted.")
            st.rerun()

    prepare_zip = st.button(
        "Prepare ZIP for Visible Files",
        disabled=not filtered_records,
        key="archive_prepare_zip",
    )
    if prepare_zip:
        with st.spinner("Preparing archive ZIP..."):
            zip_bytes, skipped = _archive_zip_bytes(filtered_records, drive_config)
        st.session_state["archive_zip_result"] = {
            "signature": repr([record.get("id") for record in filtered_records]),
            "zip_bytes": zip_bytes,
            "skipped": skipped,
            "file_name": f"Recon_Archive_{selected_month if selected_month != 'All' else 'All'}",
        }

    zip_result = st.session_state.get("archive_zip_result")
    visible_signature = repr([record.get("id") for record in filtered_records])
    if zip_result and zip_result.get("signature") == visible_signature:
        if zip_result.get("skipped"):
            st.caption(f"Skipped unavailable files: {', '.join(zip_result['skipped'])}")
        st.download_button(
            label="Download Visible Files ZIP",
            data=zip_result["zip_bytes"],
            file_name=f"{zip_result['file_name']}.zip",
            mime="application/zip",
            key="archive_download_zip",
        )


def render_monthly_summary_tab() -> None:
    st.subheader("Monthly Summary")
    st.caption("Run the collection and disbursement recons first, then generate one combined monthly summary.")

    snapshot = monthly_summary_snapshot(st.session_state)
    # Allow generating a disbursement-only summary when desired. Some users run
    # only disbursement workflows and expect the monthly summary to reflect
    # those runs without showing older collection results still present in
    # session state. Provide a simple checkbox to opt-out of including
    # collection workflows in the generated workbook.
    include_collections = st.checkbox("Include collection workflows in summary", value=True, key="monthly_include_collections")

    DISBURSEMENT_ONLY_KEYS = {
        "nsano_disb_result",
        "itc_disb_result",
        "mtn_manual_disb_result",
        "vodafone_manual_disb_result",
        "nsano_wallet_ledger_result",
        "itc_wallet_ledger_result",
    }

    use_snapshot = snapshot if include_collections else {k: v for k, v in snapshot.items() if k in DISBURSEMENT_ONLY_KEYS}

    status_rows = monthly_summary_status_rows(use_snapshot)
    ready_count = sum(1 for row in status_rows if row["Status"] == "Ready")
    current_keys = tuple(key for key in MONTHLY_RESULT_KEYS if key in use_snapshot)
    current_signature = f"{MONTHLY_SUMMARY_LAYOUT_VERSION}:{repr(use_snapshot)}"

    c1, c2 = st.columns(2)
    c1.metric("Completed Workflows", f"{ready_count:,}")
    c2.metric("Pending Workflows", f"{len(status_rows) - ready_count:,}")

    st.markdown("#### Included Runs")
    st.table(_arrow_safe_rows(status_rows))

    st.markdown("---")
    can_generate = bool(use_snapshot)
    if not can_generate:
        st.info("Run at least one reconciliation workflow before generating the monthly summary.")

    existing = st.session_state.get("monthly_summary_result")
    if existing and existing.get("snapshot_signature") != current_signature:
        st.warning("A reconciliation has changed since the last monthly summary was generated. Generate it again before downloading.")

    generate_clicked = st.button(
        "Generate Monthly Summary",
        type="primary",
        disabled=not can_generate,
        key="generate_monthly_summary",
    )

    if generate_clicked:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "Monthly_Reconciliation_Summary.xlsx"
            try:
                with st.spinner("Building monthly summary workbook..."):
                    output_bytes = build_monthly_summary_bytes(use_snapshot, output_path)
                st.session_state["monthly_summary_result"] = {
                    "snapshot_keys": current_keys,
                    "snapshot_signature": current_signature,
                    "archive_scope": "master",
                    "archive_month": _monthly_archive_month(use_snapshot),
                    "output_bytes": output_bytes,
                }
            except Exception as exc:
                st.error(f"Monthly summary generation failed: {exc}")
                st.stop()

    result = st.session_state.get("monthly_summary_result")
    if result and result.get("snapshot_signature") == current_signature:
        _mark_master_archive(result, _monthly_archive_month(use_snapshot))
        with st.spinner("Saving Monthly Summary to Google Drive archive..."):
            _ensure_result_archived("monthly_summary_result", _google_drive_config())
    if not result:
        return

    st.download_button(
        label="Download Monthly Summary",
        data=result["output_bytes"],
        file_name="Monthly_Reconciliation_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="download_monthly_summary",
    )


filtering_tab, collection_tab, disbursement_tab, ledger_tab, monthly_summary_tab, archive_tab = st.tabs([
    "Filtering",
    "Collection",
    "Disbursement",
    "Ledger",
    "Monthly Summary",
    "Recon Archive",
])

with filtering_tab:
    render_collection_filtering_tab()

with collection_tab:
    render_collections_master_tab()

with disbursement_tab:
    render_disbursement_master_tab()

with ledger_tab:
    nsano_collection_ledger_tab, nsano_disb_ledger_tab, itc_ledger_tab, wallet_ledger_filtering_tab = st.tabs([
        "Nsano Collections vs Ledger",
        "Nsano Disb Wallet vs Ledger",
        "ITC Wallet vs Ledger",
        "Voda & Zenith Ledger",
    ])

    with nsano_collection_ledger_tab:
        render_nsano_collection_ledger_tab()

    with nsano_disb_ledger_tab:
        render_nsano_wallet_ledger_tab()

    with itc_ledger_tab:
        render_itc_wallet_ledger_tab()

    with wallet_ledger_filtering_tab:
        render_wallet_ledger_filtering_tab()

with monthly_summary_tab:
    render_monthly_summary_tab()

with archive_tab:
    render_recon_archive_tab()
