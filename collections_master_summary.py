#!/usr/bin/env python3
"""Master summary workbook for collection and disbursement reconciliation results."""

from __future__ import annotations

import os
import zipfile
import datetime as dt
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Mapping

from build_reconciliation_template import (
    FAST_XLSX_COMPRESSLEVEL,
    NOT_FOUND_DETAIL_HEADERS,
    ITC_MAMBU_STATUS,
    ITC_NOT_FOUND_STATUS,
    ITC_STATUS,
    ITC_UNIDENTIFIED_STATUS,
    ITC_UPSALE_STATUS,
    ITC_WRITE_OFF_STATUS,
    STYLE_MASTER_BODY,
    STYLE_MASTER_BODY_BOLD,
    STYLE_MASTER_COUNT,
    STYLE_MASTER_HEADER,
    STYLE_MASTER_MONEY,
    STYLE_MASTER_SECTION,
    STYLE_MASTER_SECTION_COUNT,
    STYLE_MASTER_SECTION_MONEY,
    STYLE_MASTER_TITLE,
    STYLE_DASH_CARD_COUNT,
    STYLE_DASH_CARD_DETAIL,
    STYLE_DASH_CARD_DETAIL_LARGE,
    STYLE_DASH_CARD_LABEL,
    STYLE_DASH_CARD_LABEL_LARGE,
    STYLE_DASH_CARD_TOP_DARK,
    STYLE_DASH_CARD_TOP_PRIMARY,
    STYLE_DASH_CARD_VALUE,
    STYLE_DASH_CARD_VALUE_ALERT,
    STYLE_DASH_CARD_VALUE_ALERT_FULL,
    STYLE_DASH_CARD_VALUE_FULL,
    STYLE_DASH_EYEBROW,
    STYLE_DASH_FOOTER,
    STYLE_DASH_LOGO,
    STYLE_DASH_META_LABEL,
    STYLE_DASH_META_VALUE,
    STYLE_DASH_MUTED,
    STYLE_DASH_PANEL,
    STYLE_DASH_PANEL_TITLE,
    STYLE_DASH_SECTION_LABEL,
    STYLE_DASH_STATUS,
    STYLE_DASH_TABLE_COUNT,
    STYLE_DASH_TABLE_COUNT_LARGE,
    STYLE_DASH_TABLE_HEADER,
    STYLE_DASH_TABLE_HEADER_LARGE,
    STYLE_DASH_TABLE_LABEL,
    STYLE_DASH_TABLE_LABEL_LARGE,
    STYLE_DASH_TABLE_MONEY,
    STYLE_DASH_TABLE_MONEY_LARGE,
    STYLE_DASH_TABLE_TOTAL_COUNT,
    STYLE_DASH_TABLE_TOTAL_COUNT_LARGE,
    STYLE_DASH_TABLE_TOTAL_LABEL,
    STYLE_DASH_TABLE_TOTAL_LABEL_LARGE,
    STYLE_DASH_TABLE_TOTAL_MONEY,
    STYLE_DASH_TABLE_TOTAL_MONEY_LARGE,
    STYLE_DASH_TITLE,
    STYLE_DASH_VARIANCE_COUNT,
    STYLE_DASH_VARIANCE_LABEL,
    STYLE_DASH_VARIANCE_MONEY,
    VODA_COLL_STATUS,
    VODAFONE_WRITE_OFF_STATUS,
    REPORT_LOGIC_VERSION,
    WRITE_OFF_ITC_STATUS,
    WRITE_OFF_NSANO_STATUS,
    WRITE_OFF_VODAFONE_STATUS,
    WRITE_OFF_ZENITH_STATUS,
    Cell,
    dashboard_grid,
    dashboard_merge,
    dashboard_style_range,
    app_xml,
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
    ITC_DISB_STATUS,
    MTN_MANUAL_MAMBU_STATUS,
    MTN_MANUAL_MATCHED_STATUS,
    MTN_MANUAL_NOT_FOUND_STATUS,
    MTN_MANUAL_REFUND_STATUS,
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


MASTER_SUMMARY_RESULT_KEYS = (
    "nsano_result",
    "itc_result",
    "zenith_result",
    "write_off_recon_result",
    "nsano_disb_result",
    "itc_disb_result",
    "mtn_manual_disb_result",
    "vodafone_manual_disb_result",
)

MASTER_SUMMARY_LAYOUT_VERSION = "dashboard-v14-ambiguous-source-breakdown"

MASTER_SUMMARY_COLLECTION_KEYS = MASTER_SUMMARY_RESULT_KEYS[:4]
MASTER_SUMMARY_DISBURSEMENT_KEYS = MASTER_SUMMARY_RESULT_KEYS[4:]

MASTER_SUMMARY_FIELDS = {
    "nsano_result": ("charge_summary", "settlement_summary", "daily_summary", "mvn", "nvm"),
    "itc_result": (
        "itc_charge_summary",
        "itc_settlement_summary",
        "vodafone_charge_summary",
        "vodafone_settlement_summary",
        "daily_summary",
        "mvi",
        "ivm",
        "vvm",
    ),
    "zenith_result": ("charge_summary", "settlement_summary", "daily_summary", "mvz", "zvs"),
    "write_off_recon_result": ("stats",),
    "nsano_disb_result": ("daily_summary", "mvn", "nvm"),
    "itc_disb_result": ("daily_summary", "mvi", "ivm"),
    "mtn_manual_disb_result": ("daily_summary", "mvm", "mtv"),
    "vodafone_manual_disb_result": ("daily_summary", "mvv", "vtv"),
}


def master_summary_snapshot(
    state: Mapping[str, Any],
    keys: tuple[str, ...] = MASTER_SUMMARY_RESULT_KEYS,
) -> dict[str, Any]:
    snapshot: dict[str, Any] = {}
    for key in keys:
        result = state.get(key)
        if not isinstance(result, Mapping):
            continue
        snapshot[key] = {
            field: result[field]
            for field in MASTER_SUMMARY_FIELDS[key]
            if field in result
        }
    return snapshot


def build_master_summary_bytes(snapshot: Mapping[str, Any], output_path: Path) -> bytes:
    build_master_summary_workbook(output_path, snapshot)
    return output_path.read_bytes()


def build_master_summary_workbook(output_path: Path, snapshot: Mapping[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    not_found_sheets = master_not_found_sheet_specs(snapshot)
    daily_sheets = master_daily_sheet_specs(snapshot)
    write_off_analysis = _write_off_analysis_rows(snapshot)
    sheet_names = [
        "Master Summary",
        "Daily Summary",
        *(["Write-off Analysis"] if write_off_analysis else []),
        *(name for name, _groups in not_found_sheets),
        *(name for name, _rows in daily_sheets),
    ]
    rows, merges, row_heights = master_summary_dashboard_rows(snapshot)

    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=FAST_XLSX_COMPRESSLEVEL, allowZip64=True) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheet_names)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml(sheet_names))
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheet_names)))
        zf.writestr("xl/styles.xml", styles_xml())
        has_collection = any(key in snapshot for key in MASTER_SUMMARY_COLLECTION_KEYS)
        has_disbursement = any(key in snapshot for key in MASTER_SUMMARY_DISBURSEMENT_KEYS)
        column_widths = (
            [2.2, *([10.5] * 9), 3.0, *([10.5] * 9), 2.2]
            if has_collection and not has_disbursement
            else [2.2, 3.0, 11.5, 11.5, 11.5, 2.2, 11.5, 11.5, 11.5, 11.5, 2.2, 3.0, 11.5, 11.5, 11.5, 2.2, 11.5, 11.5, 11.5, 11.5, 2.2]
        )
        write_worksheet(
            zf,
            "xl/worksheets/sheet1.xml",
            rows,
            len(rows),
            21,
            column_widths,
            freeze_top_row=False,
            autofilter=False,
            merges=merges,
            style_func=None,
            row_heights=row_heights,
            show_gridlines=False,
            zoom_scale=85 if has_collection and not has_disbursement else 75,
            page_orientation="landscape",
            fit_to_width=1,
            fit_to_height=2,
        )
        daily_overview_rows, daily_overview_merges = _daily_summary_overview_rows(daily_sheets)
        write_worksheet(
            zf,
            "xl/worksheets/sheet2.xml",
            daily_overview_rows,
            len(daily_overview_rows),
            4,
            [20, 18, 24, 20],
            freeze_top_row=False,
            autofilter=False,
            merges=daily_overview_merges,
            style_func=None,
            show_gridlines=False,
            zoom_scale=95,
        )
        next_sheet_index = 3
        if write_off_analysis:
            analysis_rows, analysis_merges = write_off_analysis
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{next_sheet_index}.xml",
                analysis_rows,
                len(analysis_rows),
                5,
                [28, 16, 24, 24, 22],
                freeze_top_row=False,
                autofilter=False,
                merges=analysis_merges,
                style_func=None,
                show_gridlines=False,
                zoom_scale=95,
            )
            next_sheet_index += 1

        for sheet_index, (sheet_name, groups) in enumerate(not_found_sheets, start=next_sheet_index):
            detail_rows, detail_merges, detail_heights = _grouped_not_found_sheet_rows(
                sheet_name,
                groups,
            )
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{sheet_index}.xml",
                detail_rows,
                len(detail_rows),
                len(NOT_FOUND_DETAIL_HEADERS),
                [20, 22, 22, 26, 22, 21, 18, 20, 18, 24, 12, 48],
                freeze_top_row=False,
                autofilter=False,
                merges=detail_merges,
                style_func=None,
                row_heights=detail_heights,
                show_gridlines=False,
                zoom_scale=90,
                page_orientation="landscape",
                fit_to_width=1,
                fit_to_height=0,
            )

        daily_start = next_sheet_index + len(not_found_sheets)
        for sheet_index, (sheet_name, daily_rows) in enumerate(daily_sheets, start=daily_start):
            rows = _daily_wallet_sheet_rows(sheet_name, daily_rows)
            write_worksheet(
                zf,
                f"xl/worksheets/sheet{sheet_index}.xml",
                rows,
                len(rows),
                3,
                [18, 24, 20],
                freeze_top_row=False,
                autofilter=False,
                merges=["A1:C1"],
                style_func=None,
                show_gridlines=False,
                zoom_scale=95,
            )

    os.replace(temp_path, output_path)


def _write_off_analysis_rows(
    snapshot: Mapping[str, Any],
) -> tuple[list[list[Any]], list[str]] | None:
    """Build the master report's write-off control and investigation breakdown."""
    result = snapshot.get("write_off_recon_result")
    stats = result.get("stats", {}) if isinstance(result, Mapping) else {}
    if not isinstance(stats, Mapping) or not stats:
        return None

    statuses = (
        ("Nsano", WRITE_OFF_NSANO_STATUS),
        ("Itc", WRITE_OFF_ITC_STATUS),
        ("Vodafone", WRITE_OFF_VODAFONE_STATUS),
        ("Zenith", WRITE_OFF_ZENITH_STATUS),
        ("Ambiguous", "Ambiguous"),
        ("Not Found", "Not Found"),
    )
    status_counts = stats.get("status_counts", {})
    write_off_amounts = stats.get("status_amounts", {})
    wallet_amounts = stats.get("status_wallet_amounts", {})
    variances = stats.get("status_variances", {})

    rows: list[list[Any]] = [
        [Cell("WRITE-OFF ALLOCATION & VARIANCE ANALYSIS", STYLE_MASTER_TITLE), "", "", "", ""],
        [
            Cell(
                "Write-off file is the control · wallet amount and signed variance are shown for investigation",
                STYLE_DASH_MUTED,
            ),
            "", "", "", "",
        ],
        ["", "", "", "", ""],
        [
            Cell("Allocated Wallet", STYLE_MASTER_HEADER),
            Cell("Count", STYLE_MASTER_HEADER),
            Cell("Write-off Amount (GHC)", STYLE_MASTER_HEADER),
            Cell("Wallet Amount (GHC)", STYLE_MASTER_HEADER),
            Cell("Variance (GHC)", STYLE_MASTER_HEADER),
        ],
    ]
    for label, status in statuses:
        rows.append([
            Cell(label, STYLE_MASTER_BODY),
            Cell(int(status_counts.get(status, 0)), STYLE_MASTER_COUNT),
            Cell(_decimal_amount(write_off_amounts.get(status)), STYLE_MASTER_MONEY),
            Cell(_decimal_amount(wallet_amounts.get(status)), STYLE_MASTER_MONEY),
            Cell(_decimal_amount(variances.get(status)), STYLE_MASTER_MONEY),
        ])

    rows.extend([
        [
            Cell("TOTAL", STYLE_MASTER_SECTION),
            Cell(int(stats.get("total", 0)), STYLE_MASTER_SECTION_COUNT),
            Cell(_decimal_amount(stats.get("source_amount")), STYLE_MASTER_SECTION_MONEY),
            Cell(_decimal_amount(stats.get("wallet_amount")), STYLE_MASTER_SECTION_MONEY),
            Cell(_decimal_amount(stats.get("amount_variance")), STYLE_MASTER_SECTION_MONEY),
        ],
        ["", "", "", "", ""],
    ])

    zenith_result = snapshot.get("zenith_result", {})
    zenith_stats = zenith_result.get("zvs", {}) if isinstance(zenith_result, Mapping) else {}
    zenith_wallet = _summary_from_status(zenith_stats, ZENITH_WRITE_OFF_STATUS)
    zenith_allocated = _summary_from_status(stats, WRITE_OFF_ZENITH_STATUS)
    count_difference = int(zenith_allocated["count"]) - int(zenith_wallet["count"])
    amount_difference = (
        _decimal_amount(zenith_allocated["amount"])
        - _decimal_amount(zenith_wallet["amount"])
    )
    rows.extend([
        [Cell("ZENITH CONSISTENCY CHECK", STYLE_MASTER_HEADER), "", "", "", ""],
        [
            Cell("Source", STYLE_MASTER_HEADER),
            Cell("Count", STYLE_MASTER_HEADER),
            Cell("Amount (GHC)", STYLE_MASTER_HEADER),
            Cell("Count Difference", STYLE_MASTER_HEADER),
            Cell("Amount Difference (GHC)", STYLE_MASTER_HEADER),
        ],
        [
            Cell("Zenith Wallet Recon · Write Off", STYLE_MASTER_BODY),
            Cell(int(zenith_wallet["count"]), STYLE_MASTER_COUNT),
            Cell(_decimal_amount(zenith_wallet["amount"]), STYLE_MASTER_MONEY),
            Cell(0, STYLE_MASTER_COUNT),
            Cell(Decimal("0"), STYLE_MASTER_MONEY),
        ],
        [
            Cell("Write-off Allocation · Zenith", STYLE_MASTER_BODY),
            Cell(int(zenith_allocated["count"]), STYLE_MASTER_COUNT),
            Cell(_decimal_amount(zenith_allocated["amount"]), STYLE_MASTER_MONEY),
            Cell(count_difference, STYLE_MASTER_COUNT),
            Cell(amount_difference, STYLE_MASTER_MONEY),
        ],
        [
            Cell(
                (
                    "CONSISTENT · both reports classify the same Zenith write-offs"
                    if not count_difference and not amount_difference
                    else "INCONSISTENT · inspect Zenith candidate, routing, date and amount fields in Write_off_Recon.xlsx"
                ),
                STYLE_MASTER_SECTION,
            ),
            "", "", "", "",
        ],
        ["", "", "", "", ""],
        [Cell("INVESTIGATION STATUS", STYLE_MASTER_HEADER), Cell("Count", STYLE_MASTER_HEADER), "", "", ""],
    ])
    for label, count in sorted(stats.get("investigation_counts", {}).items()):
        rows.append([
            Cell(str(label), STYLE_MASTER_BODY),
            Cell(int(count), STYLE_MASTER_COUNT),
            "", "", "",
        ])
    if stats.get("ambiguous_breakdown"):
        rows.extend([
            ["", "", "", "", ""],
            [Cell("AMBIGUOUS BREAKDOWN", STYLE_MASTER_HEADER), Cell("Count", STYLE_MASTER_HEADER), "", "", ""],
        ])
        for label, count in sorted(stats["ambiguous_breakdown"].items()):
            rows.append([
                Cell(str(label), STYLE_MASTER_BODY),
                Cell(int(count), STYLE_MASTER_COUNT),
                "", "", "",
            ])
    zenith_check_start = len(statuses) + 7
    zenith_check_status = zenith_check_start + 4
    return rows, [
        "A1:E1",
        "A2:E2",
        f"A{zenith_check_start}:E{zenith_check_start}",
        f"A{zenith_check_status}:E{zenith_check_status}",
    ]


def _daily_summary_overview_rows(
    daily_sheets: list[tuple[str, list[Mapping[str, Any]]]],
) -> tuple[list[list[Any]], list[str]]:
    rows: list[list[Any]] = [
        [Cell("DAILY TRANSACTION SUMMARY", STYLE_MASTER_TITLE), "", "", ""],
        [
            Cell("Wallet", STYLE_MASTER_HEADER),
            Cell("Date", STYLE_MASTER_HEADER),
            Cell("Daily Amount Total (GHC)", STYLE_MASTER_HEADER),
            Cell("Daily Charges (GHC)", STYLE_MASTER_HEADER),
        ],
    ]
    for sheet_name, daily_rows in daily_sheets:
        wallet = sheet_name.removeprefix("Daily - ")
        if not daily_rows:
            rows.append([
                Cell(wallet, STYLE_MASTER_SECTION),
                Cell("No dated transactions", STYLE_MASTER_BODY),
                Cell(Decimal("0"), STYLE_MASTER_MONEY),
                Cell(Decimal("0"), STYLE_MASTER_MONEY),
            ])
            continue
        for index, daily in enumerate(daily_rows):
            rows.append([
                Cell(wallet if index == 0 else "", STYLE_MASTER_SECTION if index == 0 else STYLE_MASTER_BODY),
                Cell(str(daily.get("date", "")), STYLE_MASTER_BODY),
                Cell(_decimal_amount(daily.get("amount")), STYLE_MASTER_MONEY),
                Cell(_decimal_amount(daily.get("charges")), STYLE_MASTER_MONEY),
            ])
    return rows, ["A1:D1"]


def master_daily_sheet_specs(
    snapshot: Mapping[str, Any],
) -> list[tuple[str, list[Mapping[str, Any]]]]:
    """Collect the wallet-specific daily summaries retained by master runs."""
    sheets: list[tuple[str, list[Mapping[str, Any]]]] = []
    used_names: set[str] = set()
    for result in snapshot.values():
        if not isinstance(result, Mapping):
            continue
        summaries = result.get("daily_summary", {})
        if not isinstance(summaries, Mapping):
            continue
        for wallet, rows in summaries.items():
            if not isinstance(rows, list):
                continue
            base_name = f"Daily - {wallet}"[:31]
            sheet_name = base_name
            suffix = 2
            while sheet_name.casefold() in used_names:
                suffix_text = f" {suffix}"
                sheet_name = f"{base_name[:31 - len(suffix_text)]}{suffix_text}"
                suffix += 1
            used_names.add(sheet_name.casefold())
            sheets.append((sheet_name, [row for row in rows if isinstance(row, Mapping)]))
    return sheets


def _daily_wallet_sheet_rows(
    sheet_name: str,
    daily_rows: list[Mapping[str, Any]],
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [Cell(sheet_name.upper(), STYLE_MASTER_TITLE), "", ""],
        [
            Cell("Date", STYLE_MASTER_HEADER),
            Cell("Daily Amount Total (GHC)", STYLE_MASTER_HEADER),
            Cell("Daily Charges (GHC)", STYLE_MASTER_HEADER),
        ],
    ]
    amount_total = Decimal("0")
    charge_total = Decimal("0")
    for daily in daily_rows:
        amount = _decimal_amount(daily.get("amount"))
        charges = _decimal_amount(daily.get("charges"))
        amount_total += amount
        charge_total += charges
        rows.append([
            Cell(str(daily.get("date", "")), STYLE_MASTER_BODY),
            Cell(amount, STYLE_MASTER_MONEY),
            Cell(charges, STYLE_MASTER_MONEY),
        ])
    rows.append([
        Cell("TOTAL", STYLE_MASTER_SECTION),
        Cell(amount_total, STYLE_MASTER_SECTION_MONEY),
        Cell(charge_total, STYLE_MASTER_SECTION_MONEY),
    ])
    return rows


def _not_found_details(result: Mapping[str, Any], *stats_keys: str) -> list[dict[str, Any]]:
    details: list[dict[str, Any]] = []
    for stats_key in stats_keys:
        stats = result.get(stats_key, {})
        if not isinstance(stats, Mapping):
            continue
        rows = stats.get("not_found_details", [])
        if isinstance(rows, list):
            details.extend(row for row in rows if isinstance(row, Mapping))
    return details


def _grouped_not_found_sheet_rows(
    sheet_name: str,
    groups: list[tuple[str, list[dict[str, Any]]]],
) -> tuple[list[list[Any]], list[str], dict[int, float]]:
    rows: list[list[Any]] = [
        [Cell(sheet_name.upper(), STYLE_MASTER_TITLE), *("" for _ in range(11))],
        [
            Cell(
                "Transaction IDs requiring back-office investigation · grouped by reconciliation and source",
                STYLE_DASH_MUTED,
            ),
            *("" for _ in range(11)),
        ],
        ["" for _ in range(12)],
    ]
    merges = ["A1:L1", "A2:L2"]
    row_heights: dict[int, float] = {1: 28, 2: 22, 3: 8}

    for label, details in groups:
        section_row = len(rows) + 1
        rows.append([
            Cell(f"{label} · {len(details):,} TRANSACTION(S)", STYLE_MASTER_SECTION),
            *("" for _ in range(11)),
        ])
        merges.append(f"A{section_row}:L{section_row}")
        row_heights[section_row] = 22

        header_row = len(rows) + 1
        rows.append([Cell(header, STYLE_MASTER_HEADER) for header in NOT_FOUND_DETAIL_HEADERS])
        row_heights[header_row] = 22

        if details:
            for detail in details:
                output_row: list[Any] = []
                for header in NOT_FOUND_DETAIL_HEADERS:
                    value = detail.get(header, "")
                    if header == "Amount (GHC)":
                        style = STYLE_MASTER_MONEY
                    elif header == "Source Row":
                        style = STYLE_MASTER_COUNT
                    else:
                        style = STYLE_MASTER_BODY
                    output_row.append(Cell(value, style))
                rows.append(output_row)
                row_heights[len(rows)] = 20
        else:
            empty_row = len(rows) + 1
            rows.append([
                Cell("No Not-Found transactions", STYLE_DASH_MUTED),
                *("" for _ in range(11)),
            ])
            merges.append(f"A{empty_row}:L{empty_row}")
            row_heights[empty_row] = 20

        rows.append(["" for _ in range(12)])
        row_heights[len(rows)] = 18

    return rows, merges, row_heights


def _not_found_group(
    snapshot: Mapping[str, Any],
    result_key: str,
    stats_key: str,
    label: str,
) -> tuple[str, list[dict[str, Any]]] | None:
    result = snapshot.get(result_key)
    if not isinstance(result, Mapping):
        return None
    return label, _not_found_details(result, stats_key)


def master_not_found_sheet_specs(
    snapshot: Mapping[str, Any],
) -> list[tuple[str, list[tuple[str, list[dict[str, Any]]]]]]:
    """Return one grouped Not-Found sheet for each master domain."""
    sheet_specs: list[tuple[str, list[tuple[str, list[dict[str, Any]]]]]] = []

    collection_groups = [
        _not_found_group(snapshot, "nsano_result", "mvn", "NSANO · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "nsano_result", "nvm", "NSANO · WALLET AS SOURCE"),
        _not_found_group(snapshot, "itc_result", "mvi", "ITC / VODAFONE · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "itc_result", "ivm", "ITC · WALLET AS SOURCE"),
        _not_found_group(snapshot, "itc_result", "vvm", "VODAFONE · WALLET AS SOURCE"),
        _not_found_group(snapshot, "zenith_result", "mvz", "ZENITH · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "zenith_result", "zvs", "ZENITH · WALLET AS SOURCE"),
    ]
    present_collection_groups = [group for group in collection_groups if group is not None]
    if present_collection_groups:
        sheet_specs.append(("Collections Not Found", present_collection_groups))

    disbursement_groups = [
        _not_found_group(snapshot, "nsano_disb_result", "mvn", "NSANO · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "nsano_disb_result", "nvm", "NSANO · WALLET AS SOURCE"),
        _not_found_group(snapshot, "itc_disb_result", "mvi", "ITC · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "itc_disb_result", "ivm", "ITC · WALLET AS SOURCE"),
        _not_found_group(snapshot, "mtn_manual_disb_result", "mvm", "MTN MANUAL · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "mtn_manual_disb_result", "mtv", "MTN MANUAL · WALLET AS SOURCE"),
        _not_found_group(snapshot, "vodafone_manual_disb_result", "mvv", "VODAFONE MANUAL · MAMBU AS SOURCE"),
        _not_found_group(snapshot, "vodafone_manual_disb_result", "vtv", "VODAFONE MANUAL · WALLET AS SOURCE"),
    ]
    present_disbursement_groups = [group for group in disbursement_groups if group is not None]
    if present_disbursement_groups:
        sheet_specs.append(("Disbursement Not Found", present_disbursement_groups))

    return sheet_specs


def collection_source_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})
    not_found = _sum_summaries(*(_row_summary(row) for row in not_found_summary_rows(state)))

    rows = [
        _summary_row("Nsano", _summary_from_stats(nsano.get("nvm", {}), "mambu", "mambu_amount")),
        _summary_row("Itc", _summary_from_status(itc.get("ivm", {}), ITC_MAMBU_STATUS)),
        _summary_row("Vodafone", _summary_from_status(itc.get("vvm", {}), ITC_MAMBU_STATUS)),
        _summary_row("Zenith", _summary_from_status(zenith.get("zvs", {}), ZENITH_MAMBU_STATUS)),
    ]
    if not_found["count"] or not_found["amount"]:
        rows.append(_summary_row("Not-Found", not_found))
    return rows


def collection_mambu_source_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})

    rows = [
        _summary_row("Nsano", _summary_from_stats(nsano.get("mvn", {}), "matched", "matched_amount")),
        _summary_row("Itc", _summary_from_status(itc.get("mvi", {}), ITC_STATUS)),
        _summary_row("Vodafone", _summary_from_status(itc.get("mvi", {}), VODA_COLL_STATUS)),
        _summary_row("Zenith", _summary_from_status(zenith.get("mvz", {}), ZENITH_MATCHED_STATUS)),
        _summary_row(
            "Not-Found · Nsano",
            _summary_from_stats(nsano.get("mvn", {}), "not_found", "not_found_amount"),
        ),
        _summary_row(
            "Not-Found · ITC / Vodafone",
            _summary_from_status(itc.get("mvi", {}), ITC_NOT_FOUND_STATUS),
        ),
        _summary_row(
            "Not-Found · Zenith",
            _summary_from_status(zenith.get("mvz", {}), ZENITH_NOT_FOUND_STATUS),
        ),
    ]
    return rows


def collection_exception_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    not_found_by_wallet = {
        row["Wallet"]: _row_summary(row)
        for row in not_found_summary_rows(state)
    }
    for wallet, charges, _settlement, write_off, _refund, unidentified in collection_exception_wallet_summaries(state):
        not_found = not_found_by_wallet.get(wallet, _summary())
        rows.extend([
            {"Wallet": wallet, "Item": "Write-off", "Count": write_off["count"], "Amount": write_off["amount"]},
            {"Wallet": wallet, "Item": "Total Charges", "Count": charges["count"], "Amount": charges["amount"]},
            {"Wallet": wallet, "Item": "Unidentified", "Count": unidentified["count"], "Amount": unidentified["amount"]},
            {"Wallet": wallet, "Item": "Not Found", "Count": not_found["count"], "Amount": not_found["amount"]},
        ])

    category_rows = (
        ("Write-off", write_off_summary_rows(state)),
        ("Total Charges", charges_summary_rows(state)),
        ("Unidentified", unidentified_summary_rows(state)),
        ("Not Found", not_found_summary_rows(state)),
    )
    for label, summaries in category_rows:
        total = _sum_summaries(*(_row_summary(row) for row in summaries))
        rows.append({"Wallet": "Total", "Item": label, "Count": total["count"], "Amount": total["amount"]})
    return rows


def collection_exception_wallet_summaries(state: Mapping[str, Any]) -> list[tuple[str, dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})
    return [
        (
            "Nsano",
            nsano.get("charge_summary", _summary()),
            nsano.get("settlement_summary", _summary()),
            _summary_from_stats(nsano.get("nvm", {}), "write_off", "write_off_amount"),
            _summary(),
            _summary_from_stats(nsano.get("nvm", {}), "unidentified", "unidentified_amount"),
        ),
        (
            "Itc",
            itc.get("itc_charge_summary", _summary()),
            itc.get("itc_settlement_summary", _summary()),
            _summary_from_status(itc.get("ivm", {}), ITC_WRITE_OFF_STATUS),
            _summary_from_status(itc.get("ivm", {}), ITC_UPSALE_STATUS),
            _summary_from_status(itc.get("ivm", {}), ITC_UNIDENTIFIED_STATUS),
        ),
        (
            "Vodafone",
            itc.get("vodafone_charge_summary", _summary()),
            itc.get("vodafone_settlement_summary", _summary()),
            _summary_from_status(itc.get("vvm", {}), VODAFONE_WRITE_OFF_STATUS),
            _summary(),
            _summary_from_status(itc.get("vvm", {}), ITC_UNIDENTIFIED_STATUS),
        ),
        (
            "Zenith",
            zenith.get("charge_summary", _summary()),
            zenith.get("settlement_summary", _summary()),
            _summary_from_status(zenith.get("zvs", {}), ZENITH_WRITE_OFF_STATUS),
            _summary(),
            _summary_from_status(zenith.get("zvs", {}), ZENITH_UNIDENTIFIED_STATUS),
        ),
    ]


def _ledger_panel(
    rows: list[list[Any]],
    merges: list[str],
    row_heights: dict[int, float],
    top: int,
    section_label: str,
    panel_title: str,
    wallet_rows: list[dict[str, Any]],
    mambu_rows: list[dict[str, Any]],
    wallet_label: str,
    mambu_label: str,
    *,
    show_variance: bool = True,
    large_text: bool = False,
) -> None:
    """Draw one Wallet/Manual-vs-Mambu ledger reconciliation panel."""
    header_row = top + 4
    provider_row_count = max(5, len(wallet_rows), len(mambu_rows))
    total_row = header_row + provider_row_count + 1
    variance_top = total_row + 2
    bottom = variance_top + 4 if show_variance else total_row + 1
    header_style = STYLE_DASH_TABLE_HEADER_LARGE if large_text else STYLE_DASH_TABLE_HEADER
    label_style = STYLE_DASH_TABLE_LABEL_LARGE if large_text else STYLE_DASH_TABLE_LABEL
    count_style = STYLE_DASH_TABLE_COUNT_LARGE if large_text else STYLE_DASH_TABLE_COUNT
    money_style = STYLE_DASH_TABLE_MONEY_LARGE if large_text else STYLE_DASH_TABLE_MONEY
    total_label_style = STYLE_DASH_TABLE_TOTAL_LABEL_LARGE if large_text else STYLE_DASH_TABLE_TOTAL_LABEL
    total_count_style = STYLE_DASH_TABLE_TOTAL_COUNT_LARGE if large_text else STYLE_DASH_TABLE_TOTAL_COUNT
    total_money_style = STYLE_DASH_TABLE_TOTAL_MONEY_LARGE if large_text else STYLE_DASH_TABLE_TOTAL_MONEY

    dashboard_style_range(rows, top, 2, bottom, 20, STYLE_DASH_PANEL)
    dashboard_merge(rows, merges, top + 1, 2, top + 1, 6, section_label, STYLE_DASH_SECTION_LABEL)
    dashboard_merge(rows, merges, top + 2, 2, top + 3, 10, panel_title, STYLE_DASH_PANEL_TITLE)
    dashboard_merge(rows, merges, top + 1, 17, top + 3, 20, "Count · Amount (GH¢)", STYLE_DASH_MUTED)
    dashboard_merge(rows, merges, header_row, 2, header_row, 6, f"■  {wallet_label}", header_style)
    dashboard_merge(rows, merges, header_row, 7, header_row, 8, "COUNT", header_style)
    dashboard_merge(rows, merges, header_row, 9, header_row, 10, "AMOUNT", header_style)
    dashboard_merge(rows, merges, header_row, 12, header_row, 16, f"■  {mambu_label}", header_style)
    dashboard_merge(rows, merges, header_row, 17, header_row, 18, "COUNT", header_style)
    dashboard_merge(rows, merges, header_row, 19, header_row, 20, "AMOUNT", header_style)

    def add_provider_row(output_row: int, source_row: Mapping[str, Any], side: str) -> None:
        if side == "left":
            label_cols, count_cols, money_cols = (2, 6), (7, 8), (9, 10)
        else:
            label_cols, count_cols, money_cols = (12, 16), (17, 18), (19, 20)
        dashboard_merge(rows, merges, output_row, label_cols[0], output_row, label_cols[1], f"●  {source_row['Wallet']}", label_style)
        dashboard_merge(rows, merges, output_row, count_cols[0], output_row, count_cols[1], int(source_row["Count"]), count_style)
        dashboard_merge(rows, merges, output_row, money_cols[0], output_row, money_cols[1], _decimal_amount(source_row["Amount"]), money_style)

    for index in range(provider_row_count):
        if index < len(wallet_rows):
            add_provider_row(header_row + 1 + index, wallet_rows[index], "left")
        if index < len(mambu_rows):
            add_provider_row(header_row + 1 + index, mambu_rows[index], "right")

    wallet_total = _sum_summaries(*(_row_summary(row) for row in wallet_rows))
    mambu_total = _sum_summaries(*(_row_summary(row) for row in mambu_rows))
    dashboard_merge(rows, merges, total_row, 2, total_row, 6, "TOTAL", total_label_style)
    dashboard_merge(rows, merges, total_row, 7, total_row, 8, int(wallet_total["count"]), total_count_style)
    dashboard_merge(rows, merges, total_row, 9, total_row, 10, _decimal_amount(wallet_total["amount"]), total_money_style)
    dashboard_merge(rows, merges, total_row, 12, total_row, 16, "TOTAL", total_label_style)
    dashboard_merge(rows, merges, total_row, 17, total_row, 18, int(mambu_total["count"]), total_count_style)
    dashboard_merge(rows, merges, total_row, 19, total_row, 20, _decimal_amount(mambu_total["amount"]), total_money_style)

    if show_variance:
        variance_count = int(wallet_total["count"]) - int(mambu_total["count"])
        variance_amount = _decimal_amount(wallet_total["amount"]) - _decimal_amount(mambu_total["amount"])
        wallet_amounts = {row["Wallet"]: _decimal_amount(row["Amount"]) for row in wallet_rows}
        mambu_amounts = {row["Wallet"]: _decimal_amount(row["Amount"]) for row in mambu_rows}
        gaps = {
            name: wallet_amounts.get(name, Decimal("0")) - mambu_amounts.get(name, Decimal("0"))
            for name in set(wallet_amounts) | set(mambu_amounts)
        }
        largest_gap_name = max(gaps, key=lambda name: abs(gaps[name])) if gaps else "—"
        largest_gap_amount = abs(gaps.get(largest_gap_name, Decimal("0")))

        dashboard_merge(rows, merges, variance_top, 2, variance_top + 3, 6, "●  VARIANCE · WALLET − MAMBU", STYLE_DASH_VARIANCE_LABEL)
        dashboard_merge(rows, merges, variance_top, 7, variance_top + 3, 8, variance_count, STYLE_DASH_VARIANCE_COUNT)
        dashboard_merge(rows, merges, variance_top, 9, variance_top + 3, 11, "transaction gap", STYLE_DASH_VARIANCE_LABEL)
        dashboard_merge(rows, merges, variance_top, 12, variance_top + 3, 15, variance_amount, STYLE_DASH_VARIANCE_MONEY)
        dashboard_merge(
            rows, merges, variance_top, 16, variance_top + 3, 20,
            f"Largest gap: {largest_gap_name} · GH¢{largest_gap_amount:,.2f}",
            STYLE_DASH_VARIANCE_LABEL,
        )

    row_heights[top] = 8
    row_heights[top + 1] = 18
    row_heights[top + 2] = 22
    row_heights[top + 3] = 22
    for row_number in range(header_row, total_row + 1):
        row_heights[row_number] = 24 if large_text else 20
    row_heights[total_row + 1] = 8
    if show_variance:
        row_heights[variance_top] = 18
        row_heights[variance_top + 1] = 20
        row_heights[variance_top + 2] = 20
        row_heights[variance_top + 3] = 18
    row_heights[bottom] = 8


def _collection_exceptions_panel(
    rows: list[list[Any]],
    merges: list[str],
    row_heights: dict[int, float],
    top: int,
    categories: list[tuple[str, list[dict[str, Any]]]],
    *,
    subtitle: str = "Write-offs · Charges · Unidentified · Not Found",
) -> None:
    """Draw wallet-level collection exception tables, two per row, in a grid."""
    grid_rows = max(1, -(-len(categories) // 2))
    bottom = top + 5 + 8 * grid_rows
    dashboard_style_range(rows, top, 2, bottom, 20, STYLE_DASH_PANEL)
    dashboard_merge(rows, merges, top + 1, 2, top + 1, 7, "EXCEPTIONS", STYLE_DASH_SECTION_LABEL)
    dashboard_merge(rows, merges, top + 2, 2, top + 3, 10, "WALLET-LEVEL BREAKDOWN", STYLE_DASH_PANEL_TITLE)
    dashboard_merge(
        rows, merges, top + 2, 12, top + 3, 20,
        subtitle,
        STYLE_DASH_MUTED,
    )

    def add_table(table_top: int, start_col: int, end_col: int, title: str, detail_rows: list[dict[str, Any]]) -> None:
        label_end = start_col + 3
        count_start, count_end = start_col + 4, start_col + 5
        money_start = start_col + 6
        dashboard_merge(rows, merges, table_top, start_col, table_top, end_col, title.upper(), STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, start_col, table_top + 1, label_end, "WALLET", STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, count_start, table_top + 1, count_end, "COUNT", STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, money_start, table_top + 1, end_col, "AMOUNT", STYLE_DASH_TABLE_HEADER_LARGE)
        for index, detail in enumerate(detail_rows):
            output_row = table_top + 2 + index
            dashboard_merge(rows, merges, output_row, start_col, output_row, label_end, f"●  {detail['Wallet']}", STYLE_DASH_TABLE_LABEL_LARGE)
            dashboard_merge(rows, merges, output_row, count_start, output_row, count_end, int(detail["Count"]), STYLE_DASH_TABLE_COUNT_LARGE)
            dashboard_merge(rows, merges, output_row, money_start, output_row, end_col, _decimal_amount(detail["Amount"]), STYLE_DASH_TABLE_MONEY_LARGE)
        total_row = table_top + 2 + len(detail_rows)
        total = _sum_summaries(*(_row_summary(detail) for detail in detail_rows))
        dashboard_merge(rows, merges, total_row, start_col, total_row, label_end, "TOTAL", STYLE_DASH_TABLE_TOTAL_LABEL_LARGE)
        dashboard_merge(rows, merges, total_row, count_start, total_row, count_end, int(total["count"]), STYLE_DASH_TABLE_TOTAL_COUNT_LARGE)
        dashboard_merge(rows, merges, total_row, money_start, total_row, end_col, _decimal_amount(total["amount"]), STYLE_DASH_TABLE_TOTAL_MONEY_LARGE)
        for row_number in range(table_top, total_row + 1):
            row_heights[row_number] = 24

    positions = [
        (top + 5 + 8 * row_index, col, col + 8)
        for row_index in range(grid_rows)
        for col in (2, 12)
    ]
    for (title, detail_rows), (table_top, start_col, end_col) in zip(categories, positions):
        add_table(table_top, start_col, end_col, title, detail_rows)

    row_heights[top] = 8
    row_heights[top + 1] = 18
    row_heights[top + 2] = 22
    row_heights[top + 3] = 22
    for row_index in range(1, grid_rows):
        row_heights[top + 5 + 8 * row_index - 1] = 8
    row_heights[bottom] = 8


def _disbursement_exceptions_panel(
    rows: list[list[Any]],
    merges: list[str],
    row_heights: dict[int, float],
    top: int,
    wallet_not_found_rows: list[dict[str, Any]],
    mambu_not_found_rows: list[dict[str, Any]],
    refund_rows: list[dict[str, Any]],
) -> int:
    """Draw Not-Found (both directions) and Refund wallet breakdown tables.

    Row 1: wallet-side not-found (in wallet, missing from Mambu) next to
    Mambu-side not-found (in Mambu, missing from wallet). Row 2: refunds.
    Returns the panel's bottom row so callers can size the sheet around it.
    """
    row_a_top = top + 5
    row_a_height = 3 + max(len(wallet_not_found_rows), len(mambu_not_found_rows), 1)
    row_b_top = row_a_top + row_a_height + 1
    row_b_height = 3 + max(len(refund_rows), 1)
    bottom = row_b_top + row_b_height

    dashboard_style_range(rows, top, 2, bottom, 20, STYLE_DASH_PANEL)
    dashboard_merge(rows, merges, top + 1, 2, top + 1, 7, "EXCEPTIONS", STYLE_DASH_SECTION_LABEL)
    dashboard_merge(rows, merges, top + 2, 2, top + 3, 10, "WALLET-LEVEL BREAKDOWN", STYLE_DASH_PANEL_TITLE)
    dashboard_merge(
        rows, merges, top + 2, 12, top + 3, 20,
        "Not Found (Wallet vs Mambu, both directions) · Refunds",
        STYLE_DASH_MUTED,
    )

    def add_table(table_top: int, start_col: int, end_col: int, title: str, detail_rows: list[dict[str, Any]]) -> None:
        label_end = start_col + 3
        count_start, count_end = start_col + 4, start_col + 5
        money_start = start_col + 6
        dashboard_merge(rows, merges, table_top, start_col, table_top, end_col, title.upper(), STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, start_col, table_top + 1, label_end, "WALLET", STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, count_start, table_top + 1, count_end, "COUNT", STYLE_DASH_TABLE_HEADER_LARGE)
        dashboard_merge(rows, merges, table_top + 1, money_start, table_top + 1, end_col, "AMOUNT", STYLE_DASH_TABLE_HEADER_LARGE)
        for index, detail in enumerate(detail_rows):
            output_row = table_top + 2 + index
            dashboard_merge(rows, merges, output_row, start_col, output_row, label_end, f"●  {detail['Wallet']}", STYLE_DASH_TABLE_LABEL_LARGE)
            dashboard_merge(rows, merges, output_row, count_start, output_row, count_end, int(detail["Count"]), STYLE_DASH_TABLE_COUNT_LARGE)
            dashboard_merge(rows, merges, output_row, money_start, output_row, end_col, _decimal_amount(detail["Amount"]), STYLE_DASH_TABLE_MONEY_LARGE)
        total_row = table_top + 2 + len(detail_rows)
        total = _sum_summaries(*(_row_summary(detail) for detail in detail_rows))
        dashboard_merge(rows, merges, total_row, start_col, total_row, label_end, "TOTAL", STYLE_DASH_TABLE_TOTAL_LABEL_LARGE)
        dashboard_merge(rows, merges, total_row, count_start, total_row, count_end, int(total["count"]), STYLE_DASH_TABLE_TOTAL_COUNT_LARGE)
        dashboard_merge(rows, merges, total_row, money_start, total_row, end_col, _decimal_amount(total["amount"]), STYLE_DASH_TABLE_TOTAL_MONEY_LARGE)
        for row_number in range(table_top, total_row + 1):
            row_heights[row_number] = 24

    add_table(row_a_top, 2, 10, "Not Found · Wallet Side", wallet_not_found_rows)
    add_table(row_a_top, 12, 20, "Not Found · Mambu Side", mambu_not_found_rows)
    add_table(row_b_top, 2, 10, "Refund Per Wallet", refund_rows)

    row_heights[top] = 8
    row_heights[top + 1] = 18
    row_heights[top + 2] = 22
    row_heights[top + 3] = 22
    row_heights[row_a_top + row_a_height] = 8
    row_heights[bottom] = 8
    return bottom


def _adjustments_table(
    rows: list[list[Any]],
    merges: list[str],
    row_heights: dict[int, float],
    top: int,
    detail_rows: list[tuple[str, Mapping[str, Any]]],
) -> None:
    """Draw a compact totals-only adjustments table when both domains ran this session."""
    header_row = top + 2
    bottom = header_row + len(detail_rows)
    dashboard_style_range(rows, top, 2, bottom, 20, STYLE_DASH_PANEL)
    dashboard_merge(rows, merges, top, 2, top + 1, 7, "ADJUSTMENTS & REFUNDS", STYLE_DASH_PANEL_TITLE)
    dashboard_merge(
        rows, merges, top, 8, top + 1, 20,
        "Write-offs, charges, settlement and refunds this session",
        STYLE_DASH_MUTED,
    )
    dashboard_merge(rows, merges, header_row, 2, header_row, 12, "CATEGORY", STYLE_DASH_TABLE_HEADER)
    dashboard_merge(rows, merges, header_row, 13, header_row, 16, "COUNT", STYLE_DASH_TABLE_HEADER)
    dashboard_merge(rows, merges, header_row, 17, header_row, 20, "AMOUNT", STYLE_DASH_TABLE_HEADER)
    for index, (label, summary) in enumerate(detail_rows):
        output_row = header_row + 1 + index
        dashboard_merge(rows, merges, output_row, 2, output_row, 12, f"●  {label}", STYLE_DASH_TABLE_LABEL)
        dashboard_merge(rows, merges, output_row, 13, output_row, 16, int(summary["count"]), STYLE_DASH_TABLE_COUNT)
        dashboard_merge(rows, merges, output_row, 17, output_row, 20, _decimal_amount(summary["amount"]), STYLE_DASH_TABLE_MONEY)

    row_heights[top] = 22
    row_heights[top + 1] = 22
    row_heights[header_row] = 20
    for index in range(len(detail_rows)):
        row_heights[header_row + 1 + index] = 20


def master_summary_dashboard_rows(
    state: Mapping[str, Any],
) -> tuple[list[list[Any]], list[str], dict[int, float]]:
    """Build the presentation-first Master Summary used by every new export.

    The layout adapts to whichever workflows actually ran this session: a
    collection-only run gets a collection-led ledger panel, a disbursement-only
    run gets a disbursement-led one (with matching labels), and a session with
    both gets a panel for each plus a combined adjustments rollup.
    """
    has_collection = any(key in state for key in MASTER_SUMMARY_COLLECTION_KEYS)
    has_disbursement = any(key in state for key in MASTER_SUMMARY_DISBURSEMENT_KEYS)
    collection_only = has_collection and not has_disbursement
    row_count, col_count = (61, 21) if collection_only else ((62, 21) if (has_collection and has_disbursement) else (53, 21))
    rows = dashboard_grid(row_count, col_count)
    merges: list[str] = []
    row_heights: dict[int, float] = {
        1: 8, 2: 18, 3: 24, 4: 24, 5: 18, 6: 22, 7: 10,
        8: 5, 9: 20, 10: 24, 11: 24, 12: 21, 13: 10, 14: 10,
        32: 10, 50: 12,
    }

    collection_wallet = collection_source_summary_rows(state)
    collection_mambu = collection_mambu_source_summary_rows(state)
    disbursement_wallet = disbursement_wallet_summary_rows(state)
    disbursement_mambu = disbursement_mambu_summary_rows(state)
    disbursement_not_found = disbursement_not_found_summary_rows(state)
    disbursement_mambu_not_found = disbursement_mambu_not_found_summary_rows(state)
    write_offs = write_off_summary_rows(state)
    unidentified = unidentified_summary_rows(state)
    charges = charges_summary_rows(state)
    not_found = not_found_summary_rows(state)
    settlement = settlement_summary_rows(state)
    refunds = refund_summary_rows(state)

    write_off_total = _sum_summaries(*(_row_summary(row) for row in write_offs))
    unidentified_total = _sum_summaries(*(_row_summary(row) for row in unidentified))
    charges_total = _sum_summaries(*(_row_summary(row) for row in charges))
    settlement_total = _sum_summaries(*(_row_summary(row) for row in settlement))
    refund_total = _sum_summaries(*(_row_summary(row) for row in refunds))

    if has_collection and has_disbursement:
        scope_label = "Collections & Disbursements"
        subtitle = "Collections and disbursements reconciled against Mambu core banking  •  All amounts in GH¢"
    elif has_disbursement:
        scope_label = "Disbursements"
        subtitle = "Disbursement ledger reconciled against Mambu core banking  •  All amounts in GH¢"
    else:
        scope_label = "Collections"
        subtitle = "Wallet ledger reconciled against Mambu core banking  •  All amounts in GH¢"

    # Header
    dashboard_merge(rows, merges, 2, 2, 5, 3, "₵", STYLE_DASH_LOGO)
    dashboard_merge(rows, merges, 2, 4, 2, 11, "FINANCE OPERATIONS", STYLE_DASH_EYEBROW)
    dashboard_merge(rows, merges, 3, 4, 5, 14, "Master Recon Summary", STYLE_DASH_TITLE)
    dashboard_merge(rows, merges, 2, 15, 2, 17, "GENERATED", STYLE_DASH_META_LABEL)
    dashboard_merge(rows, merges, 3, 15, 5, 17, dt.date.today().strftime("%d %b %Y"), STYLE_DASH_META_VALUE)
    dashboard_merge(rows, merges, 6, 2, 6, 20, subtitle, STYLE_DASH_META_LABEL)

    def add_kpi(top_row: int, start_col: int, end_col: int, label: str, value: Any, value_style: int, detail: str, *, dark_top: bool) -> None:
        top_style = STYLE_DASH_CARD_TOP_DARK if dark_top else STYLE_DASH_CARD_TOP_PRIMARY
        label_style = STYLE_DASH_CARD_LABEL_LARGE if collection_only else STYLE_DASH_CARD_LABEL
        detail_style = STYLE_DASH_CARD_DETAIL_LARGE if collection_only else STYLE_DASH_CARD_DETAIL
        dashboard_merge(rows, merges, top_row, start_col, top_row, end_col, "", top_style)
        dashboard_merge(rows, merges, top_row + 1, start_col, top_row + 1, end_col, label.upper(), label_style)
        dashboard_merge(rows, merges, top_row + 2, start_col, top_row + 3, end_col, value, value_style)
        dashboard_merge(rows, merges, top_row + 4, start_col, top_row + 4, end_col, detail, detail_style)
        dashboard_merge(rows, merges, top_row + 5, start_col, top_row + 5, end_col, "", detail_style)

    kpi_specs: list[tuple[str, Any, int, str, bool]] = []
    collection_variance_count = collection_variance_amount = 0
    disb_variance_count = disb_variance_amount = 0
    if has_collection:
        wallet_total = _sum_summaries(*(_row_summary(row) for row in collection_wallet))
        mambu_total = _sum_summaries(*(_row_summary(row) for row in collection_mambu))
        collection_variance_count = int(wallet_total["count"]) - int(mambu_total["count"])
        collection_variance_amount = _decimal_amount(wallet_total["amount"]) - _decimal_amount(mambu_total["amount"])
        if not has_disbursement:
            kpi_specs.extend([
                (
                    "Total Collection Wallets", _decimal_amount(wallet_total["amount"]), STYLE_DASH_CARD_VALUE_FULL,
                    f"{int(wallet_total['count']):,} txns", False,
                ),
                (
                    "Total Collections for Mambu", _decimal_amount(mambu_total["amount"]), STYLE_DASH_CARD_VALUE_FULL,
                    f"{int(mambu_total['count']):,} txns", False,
                ),
                (
                    "Collections Variances", collection_variance_amount, STYLE_DASH_CARD_VALUE_ALERT_FULL,
                    f"{collection_variance_count:+,} transaction gap", False,
                ),
                (
                    "Total Unidentified", _decimal_amount(unidentified_total["amount"]), STYLE_DASH_CARD_VALUE_FULL,
                    f"{int(unidentified_total['count']):,} txns", True,
                ),
                (
                    "Total Write-Off", _decimal_amount(write_off_total["amount"]), STYLE_DASH_CARD_VALUE_FULL,
                    f"{int(write_off_total['count']):,} txns", True,
                ),
                (
                    "Total Charges", _decimal_amount(charges_total["amount"]), STYLE_DASH_CARD_VALUE_FULL,
                    f"{int(charges_total['count']):,} txns", True,
                ),
            ])
        else:
            kpi_specs.append((
                "Total collections", _decimal_amount(wallet_total["amount"]), STYLE_DASH_CARD_VALUE,
                f"{_decimal_amount(wallet_total['amount']):,.2f}  •  {int(wallet_total['count']):,} txns", False,
            ))
    if has_disbursement:
        disb_wallet_total = _sum_summaries(*(_row_summary(row) for row in disbursement_wallet))
        disb_mambu_total = _sum_summaries(*(_row_summary(row) for row in disbursement_mambu))
        disb_variance_count = int(disb_wallet_total["count"]) - int(disb_mambu_total["count"])
        disb_variance_amount = _decimal_amount(disb_wallet_total["amount"]) - _decimal_amount(disb_mambu_total["amount"])
        if not has_collection:
            kpi_specs.extend([
                (
                    "Total Disbursement", _decimal_amount(disb_wallet_total["amount"]), STYLE_DASH_CARD_VALUE,
                    f"{_decimal_amount(disb_wallet_total['amount']):,.2f}  •  {int(disb_wallet_total['count']):,} txns", False,
                ),
                (
                    "Total Mambu Wallets", _decimal_amount(disb_mambu_total["amount"]), STYLE_DASH_CARD_VALUE,
                    f"{_decimal_amount(disb_mambu_total['amount']):,.2f}  •  {int(disb_mambu_total['count']):,} txns", False,
                ),
                (
                    "Disbursements variance · Wallet ↔ Mambu", disb_variance_amount, STYLE_DASH_CARD_VALUE_ALERT,
                    f"{disb_variance_count:+,} transaction gap", False,
                ),
                (
                    "Total refunds", _decimal_amount(refund_total["amount"]), STYLE_DASH_CARD_VALUE,
                    f"{_decimal_amount(refund_total['amount']):,.2f}  •  {int(refund_total['count']):,} txns", True,
                ),
            ])
        else:
            kpi_specs.append((
                "Total disbursements", _decimal_amount(disb_wallet_total["amount"]), STYLE_DASH_CARD_VALUE,
                f"{_decimal_amount(disb_wallet_total['amount']):,.2f}  •  {int(disb_wallet_total['count']):,} txns", False,
            ))
    if has_collection and has_disbursement:
        kpi_specs.append((
            "Collections variance · Wallet ↔ Mambu", collection_variance_amount, STYLE_DASH_CARD_VALUE_ALERT,
            f"{collection_variance_count:+,} transaction gap", False,
        ))
    if has_disbursement and has_collection:
        kpi_specs.append((
            "Disbursements variance · Wallet ↔ Mambu", disb_variance_amount, STYLE_DASH_CARD_VALUE_ALERT,
            f"{disb_variance_count:+,} transaction gap", False,
        ))
    if has_collection and has_disbursement:
        kpi_specs.append((
            "Total write-offs", _decimal_amount(write_off_total["amount"]), STYLE_DASH_CARD_VALUE,
            f"{_decimal_amount(write_off_total['amount']):,.2f}  •  {int(write_off_total['count']):,} txns", True,
        ))
    if has_disbursement and has_collection:
        kpi_specs.append((
            "Total refunds", _decimal_amount(refund_total["amount"]), STYLE_DASH_CARD_VALUE,
            f"{_decimal_amount(refund_total['amount']):,.2f}  •  {int(refund_total['count']):,} txns", True,
        ))
    if has_collection and has_disbursement:
        kpi_specs.append((
            "Total charges", _decimal_amount(charges_total["amount"]), STYLE_DASH_CARD_VALUE,
            f"{_decimal_amount(charges_total['amount']):,.2f}  •  {int(charges_total['count']):,} txns", True,
        ))
    target_kpi_count = 6 if has_collection and not has_disbursement else 4
    kpi_specs = kpi_specs[:target_kpi_count]
    if len(kpi_specs) < target_kpi_count:
        combined_count = (int(wallet_total["count"]) if has_collection else 0) + (
            int(disb_wallet_total["count"]) if has_disbursement else 0
        )
        kpi_specs.append((
            "Total transactions reconciled", combined_count, STYLE_DASH_CARD_COUNT,
            f"Across {scope_label.lower()} this session", True,
        ))

    kpi_positions = (
        [(8, 2, 7), (8, 8, 14), (8, 15, 20), (15, 2, 7), (15, 8, 14), (15, 15, 20)]
        if collection_only
        else [(8, 2, 5), (8, 7, 10), (8, 12, 15), (8, 17, 20)]
    )
    for (top_row, start_col, end_col), (label, value, style, detail, dark_top) in zip(
        kpi_positions, kpi_specs,
    ):
        add_kpi(top_row, start_col, end_col, label, value, style, detail, dark_top=dark_top)

    if collection_only:
        for top_row in (8, 15):
            row_heights[top_row] = 6
            row_heights[top_row + 1] = 24
            row_heights[top_row + 2] = 30
            row_heights[top_row + 3] = 30
            row_heights[top_row + 4] = 24
            row_heights[top_row + 5] = 10
        row_heights[14] = 10
        row_heights[21] = 10

    variance_flagged = (has_collection and (collection_variance_count or collection_variance_amount)) or (
        has_disbursement and (disb_variance_count or disb_variance_amount)
    )
    status_text = "●  Variance flagged" if variance_flagged else "●  Reconciled"
    dashboard_merge(rows, merges, 2, 18, 5, 20, status_text, STYLE_DASH_STATUS)

    # Primary ledger panel: whichever domain ran leads the page.
    if has_collection:
        ledger_top = 22 if collection_only else 15
        _ledger_panel(
            rows, merges, row_heights, ledger_top, "COLLECTIONS", "COLLECTIONS ANALYSIS",
            collection_wallet, collection_mambu, "COLLECTIONS PER WALLET", "MAMBU COLLECTIONS PER WALLET",
            show_variance=False,
            large_text=collection_only,
        )
    else:
        _ledger_panel(
            rows, merges, row_heights, 15, "DISBURSEMENTS", "DISBURSEMENTS ANALYSIS",
            disbursement_wallet, disbursement_mambu, "DISBURSEMENT PER WALLET", "MAMBU PER WALLETS",
            show_variance=False,
        )

    if has_collection and has_disbursement:
        _ledger_panel(
            rows, merges, row_heights, 33, "DISBURSEMENTS", "DISBURSEMENTS ANALYSIS",
            disbursement_wallet, disbursement_mambu, "DISBURSEMENT PER WALLET", "MAMBU PER WALLETS",
            show_variance=False,
        )
        disbursement_wallet_not_found_total = _sum_summaries(*(_row_summary(row) for row in disbursement_not_found))
        disbursement_mambu_not_found_total = _sum_summaries(*(_row_summary(row) for row in disbursement_mambu_not_found))
        _adjustments_table(rows, merges, row_heights, 51, [
            ("Write-offs", write_off_total),
            ("Unidentified", unidentified_total),
            ("Charges", charges_total),
            ("Settlement", settlement_total),
            ("Refunds", refund_total),
            ("Not-Found · Wallet Side", disbursement_wallet_not_found_total),
            ("Not-Found · Mambu Side", disbursement_mambu_not_found_total),
        ])
    elif has_disbursement:
        _disbursement_exceptions_panel(
            rows, merges, row_heights, 33,
            disbursement_not_found, disbursement_mambu_not_found, refunds,
        )
    else:
        _collection_exceptions_panel(
            rows, merges, row_heights, 38,
            [
                ("Write-Off by Wallet", write_offs),
                ("Charges by Wallet", charges),
                ("Unidentified by Wallet", unidentified),
                ("Not Found by Wallet", not_found),
            ],
        )

    footer_row = row_count
    row_heights[footer_row] = 18
    dashboard_merge(
        rows,
        merges,
        footer_row,
        2,
        footer_row,
        10,
        f"Master Recon Summary · Finance Operations · {REPORT_LOGIC_VERSION}",
        STYLE_DASH_FOOTER,
    )
    dashboard_merge(rows, merges, footer_row, 12, footer_row, 20, f"Source: {scope_label}", STYLE_DASH_FOOTER)
    return rows, merges, row_heights


def master_summary_sheet_rows(state: Mapping[str, Any]) -> list[list[Any]]:
    collection_wallet_rows = collection_source_summary_rows(state)
    collection_mambu_rows = collection_mambu_source_summary_rows(state)
    disbursement_wallet_rows = disbursement_wallet_summary_rows(state)
    disbursement_mambu_rows = disbursement_mambu_summary_rows(state)
    rows: list[list[Any]] = [
        [Cell("MASTER RECON SUMMARY", STYLE_MASTER_TITLE), "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        _paired(
            _header_cells("Collections Recon", "Count", "Amount"),
            _header_cells("Disbursement Recon", "Count", "Amount"),
        ),
    ]

    def add_blocks(left: list[list[Any]], right: list[list[Any]] | None = None) -> None:
        rows.extend(_paired_blocks(left, right or []))
        rows.append(_blank_row())

    add_blocks(
        _section_with_rows("TOTAL COLLECTIONS(Wallet)", collection_wallet_rows),
        _section_with_rows("TOTAL DISBURSEMENT(Wallet)", disbursement_wallet_rows),
    )
    add_blocks(
        _section_with_rows("TOTAL COLLECTIONS(Mambu)", collection_mambu_rows),
        _section_with_rows("TOTAL DISBURSEMENT(Mambu)", disbursement_mambu_rows),
    )
    add_blocks(
        _section_with_rows(
            "Collections Analysis (Wallet - Mambu)",
            [_variance_summary_row("Difference", collection_wallet_rows, collection_mambu_rows)],
        ),
        _section_with_rows(
            "Disbursement Analysis (Wallet - Mambu)",
            [_variance_summary_row("Difference", disbursement_wallet_rows, disbursement_mambu_rows)],
        ),
    )
    add_blocks(
        _section_with_rows("Total Write off", write_off_summary_rows(state)),
        _section_with_rows("Total Refund", refund_summary_rows(state)),
    )
    add_blocks(_section_with_rows("Total Unidentified", unidentified_summary_rows(state)))
    add_blocks(_section_with_rows("Total Charges", charges_summary_rows(state)))
    add_blocks(_section_with_rows("Total Not Found", not_found_summary_rows(state)))
    return rows[:-1] if rows and rows[-1] == _blank_row() else rows


def disbursement_wallet_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    not_found = _sum_summaries(*(_row_summary(row) for row in disbursement_not_found_summary_rows(state)))

    rows = [
        _summary_row("Nsano", _summary_from_status(nsano.get("nvm", {}), DISB_MAMBU_STATUS)),
        _summary_row("Itc", _summary_from_status(itc.get("ivm", {}), DISB_MAMBU_STATUS)),
        _summary_row("Vodafone(Manual)", _summary_from_status(vodafone.get("vtv", {}), VODAFONE_MANUAL_MAMBU_STATUS)),
        _summary_row("MTN Manual", _summary_from_status(mtn.get("mtv", {}), MTN_MANUAL_MAMBU_STATUS)),
    ]
    if not_found["count"] or not_found["amount"]:
        rows.append(_summary_row("Not-Found", not_found))
    return rows


def disbursement_not_found_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Wallet-side not-found: wallet transactions with no matching Mambu record."""
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    return [
        _summary_row("Nsano", _summary_from_status(nsano.get("nvm", {}), DISB_NOT_FOUND_STATUS)),
        _summary_row("Itc", _summary_from_status(itc.get("ivm", {}), ITC_DISB_NOT_FOUND_STATUS)),
        _summary_row("Vodafone(Manual)", _summary_from_status(vodafone.get("vtv", {}), VODAFONE_MANUAL_NOT_FOUND_STATUS)),
        _summary_row("MTN Manual", _summary_from_status(mtn.get("mtv", {}), MTN_MANUAL_NOT_FOUND_STATUS)),
    ]


def disbursement_mambu_not_found_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Mambu-side not-found: Mambu transactions with no matching wallet record."""
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    return [
        _summary_row("Nsano", _summary_from_status(nsano.get("mvn", {}), DISB_NOT_FOUND_STATUS)),
        _summary_row("Itc", _summary_from_status(itc.get("mvi", {}), ITC_DISB_NOT_FOUND_STATUS)),
        _summary_row("Vodafone(Manual)", _summary_from_status(vodafone.get("mvv", {}), VODAFONE_MANUAL_NOT_FOUND_STATUS)),
        _summary_row("MTN Manual", _summary_from_status(mtn.get("mvm", {}), MTN_MANUAL_NOT_FOUND_STATUS)),
    ]


def disbursement_mambu_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    nsano = state.get("nsano_disb_result", {})
    itc = state.get("itc_disb_result", {})
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    not_found = _sum_summaries(*(_row_summary(row) for row in disbursement_mambu_not_found_summary_rows(state)))

    rows = [
        _summary_row("Nsano", _summary_from_status(nsano.get("mvn", {}), DISB_NSANO_STATUS)),
        _summary_row("Itc", _summary_from_status(itc.get("mvi", {}), ITC_DISB_STATUS)),
        _summary_row("Vodafone(Manual)", _summary_from_status(vodafone.get("mvv", {}), VODAFONE_MANUAL_MATCHED_STATUS)),
        _summary_row("MTN Manual", _summary_from_status(mtn.get("mvm", {}), MTN_MANUAL_MATCHED_STATUS)),
    ]
    if not_found["count"] or not_found["amount"]:
        rows.append(_summary_row("Not-Found", not_found))
    return rows


def write_off_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    write_off = state.get("write_off_recon_result", {})
    stats = write_off.get("stats", {}) if isinstance(write_off, Mapping) else {}
    if stats:
        # The allocation report is the control: these are write-off-file amounts
        # assigned using channel-aware wallet matching.
        return [
            _summary_row("Nsano", _summary_from_status(stats, WRITE_OFF_NSANO_STATUS)),
            _summary_row("Itc", _summary_from_status(stats, WRITE_OFF_ITC_STATUS)),
            _summary_row("Vodafone", _summary_from_status(stats, WRITE_OFF_VODAFONE_STATUS)),
            _summary_row("Zenith", _summary_from_status(stats, WRITE_OFF_ZENITH_STATUS)),
        ]

    # Retain wallet-side summaries as a fallback when the standalone allocation
    # workflow was not run.
    return [
        _summary_row(wallet, write_off)
        for wallet, _charges, _settlement, write_off, _refund, _unidentified
        in collection_exception_wallet_summaries(state)
    ]


def unidentified_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [
        _summary_row(wallet, unidentified)
        for wallet, _charges, _settlement, _write_off, _refund, unidentified
        in collection_exception_wallet_summaries(state)
    ]


def not_found_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    nsano = state.get("nsano_result", {})
    itc = state.get("itc_result", {})
    zenith = state.get("zenith_result", {})
    return [
        _summary_row("Nsano", _summary_from_stats(nsano.get("nvm", {}), "not_found", "not_found_amount")),
        _summary_row("Itc", _summary_from_status(itc.get("ivm", {}), ITC_NOT_FOUND_STATUS)),
        _summary_row("Vodafone", _summary_from_status(itc.get("vvm", {}), ITC_NOT_FOUND_STATUS)),
        _summary_row("Zenith", _summary_from_status(zenith.get("zvs", {}), ZENITH_NOT_FOUND_STATUS)),
    ]


def refund_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    vodafone = state.get("vodafone_manual_disb_result", {})
    mtn = state.get("mtn_manual_disb_result", {})
    return [
        _summary_row("Vodafone(Manual)", _summary_from_status(vodafone.get("vtv", {}), VODAFONE_MANUAL_REFUND_STATUS)),
        _summary_row("MTN Manual", _summary_from_status(mtn.get("mtv", {}), MTN_MANUAL_REFUND_STATUS)),
    ]


def charges_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [
        _summary_row(wallet, charges)
        for wallet, charges, _settlement, _write_off, _refund, _unidentified
        in collection_exception_wallet_summaries(state)
    ]


def settlement_summary_rows(state: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [
        _summary_row(wallet, settlement)
        for wallet, _charges, settlement, _write_off, _refund, _unidentified
        in collection_exception_wallet_summaries(state)
    ]


def _section_with_rows(label: str, rows: list[dict[str, Any]]) -> list[list[Cell]]:
    total = _sum_summaries(*(_row_summary(row) for row in rows))
    return [
        _section_cells(label, total),
        *(_body_cells(row["Wallet"], _row_summary(row)) for row in rows),
    ]


def _paired_blocks(left_rows: list[list[Any]], right_rows: list[list[Any]]) -> list[list[Any]]:
    row_count = max(len(left_rows), len(right_rows))
    rows: list[list[Any]] = []
    for idx in range(row_count):
        left = left_rows[idx] if idx < len(left_rows) else ["", "", ""]
        right = right_rows[idx] if idx < len(right_rows) else ["", "", ""]
        rows.append(_paired(left, right))
    return rows


def _paired(left: list[Any], right: list[Any]) -> list[Any]:
    return [*left, "", *right]


def _header_cells(label: str, count_label: str, amount_label: str) -> list[Cell]:
    return [
        Cell(label, STYLE_MASTER_HEADER),
        Cell(count_label, STYLE_MASTER_HEADER),
        Cell(amount_label, STYLE_MASTER_HEADER),
    ]


def _section_cells(label: str, summary: Mapping[str, Any]) -> list[Cell]:
    return [
        Cell(label, STYLE_MASTER_SECTION),
        Cell(int(summary.get("count", 0)), STYLE_MASTER_SECTION_COUNT),
        Cell(_decimal_amount(summary.get("amount", Decimal("0"))), STYLE_MASTER_SECTION_MONEY),
    ]


def _body_cells(label: str, summary: Mapping[str, Any]) -> list[Cell]:
    label_style = STYLE_MASTER_BODY_BOLD if str(label).casefold() == "total" else STYLE_MASTER_BODY
    return [
        Cell(label, label_style),
        Cell(int(summary.get("count", 0)), STYLE_MASTER_COUNT),
        Cell(_decimal_amount(summary.get("amount", Decimal("0"))), STYLE_MASTER_MONEY),
    ]


def _blank_row() -> list[str]:
    return ["", "", "", "", "", "", ""]


def _summary_row(wallet: str, summary: Mapping[str, Any]) -> dict[str, Any]:
    return {"Wallet": wallet, "Count": int(summary.get("count", 0)), "Amount": _decimal_amount(summary.get("amount", Decimal("0")))}


def _variance_summary_row(label: str, wallet_rows: list[dict[str, Any]], mambu_rows: list[dict[str, Any]]) -> dict[str, Any]:
    wallet_total = _sum_summaries(*(_row_summary(row) for row in wallet_rows))
    mambu_total = _sum_summaries(*(_row_summary(row) for row in mambu_rows))
    return _summary_row(
        label,
        _summary(
            int(wallet_total["count"]) - int(mambu_total["count"]),
            _decimal_amount(wallet_total["amount"]) - _decimal_amount(mambu_total["amount"]),
        ),
    )


def _row_summary(row: Mapping[str, Any]) -> dict[str, Any]:
    return _summary(row.get("Count", 0), row.get("Amount", Decimal("0")))


def _summary(count: int = 0, amount: Any = Decimal("0")) -> dict[str, Any]:
    return {"count": int(count or 0), "amount": _decimal_amount(amount)}


def _summary_from_stats(stats: Mapping[str, Any], count_key: str, amount_key: str) -> dict[str, Any]:
    if not stats:
        return _summary()
    return _summary(stats.get(count_key, 0), stats.get(amount_key, Decimal("0")))


def _summary_from_status(stats: Mapping[str, Any], status: str) -> dict[str, Any]:
    if not stats:
        return _summary()
    return _summary(
        stats.get("status_counts", {}).get(status, 0),
        stats.get("status_amounts", {}).get(status, Decimal("0")),
    )


def _sum_summaries(*summaries: Mapping[str, Any]) -> dict[str, Any]:
    return _summary(
        sum(int(summary.get("count", 0)) for summary in summaries),
        sum((_decimal_amount(summary.get("amount", Decimal("0"))) for summary in summaries), Decimal("0")),
    )


def _decimal_amount(value: Any) -> Decimal:
    text = str(value or "").strip().replace(",", "").replace("\t", "")
    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _master_summary_style(_row_number: int, _col_number: int, value: Any) -> int | None:
    if isinstance(value, Cell):
        return value.style
    return None
