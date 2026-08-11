#!/usr/bin/env python3
"""Build the finance automation logic reference DOCX.

This intentionally uses only the Python standard library and direct OOXML so
the document can be generated in the current environment without python-docx.
"""

from __future__ import annotations

import html
import zipfile
from pathlib import Path


OUT = Path("Finance_Automation_Logic_Reference.docx")


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def run(text: str, *, bold: bool = False) -> str:
    props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def para(text: str = "", *, style: str = "Normal", num_id: int | None = None, level: int = 0) -> str:
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ""
    num_xml = ""
    if num_id is not None:
        num_xml = (
            "<w:numPr>"
            f'<w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{num_id}"/>'
            "</w:numPr>"
        )
    return f"<w:p><w:pPr>{style_xml}{num_xml}</w:pPr>{run(text)}</w:p>"


def label_para(label: str, text: str) -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
        f"{run(label, bold=True)}{run(text)}"
        "</w:p>"
    )


def cell(text: object, width: int, *, header: bool = False) -> str:
    fill = '<w:shd w:fill="E8EEF5"/>' if header else ""
    style = "TableHeader" if header else "TableText"
    return (
        "<w:tc>"
        "<w:tcPr>"
        f'<w:tcW w:w="{width}" w:type="dxa"/>'
        "<w:tcMar>"
        '<w:top w:w="80" w:type="dxa"/>'
        '<w:bottom w:w="80" w:type="dxa"/>'
        '<w:start w:w="120" w:type="dxa"/>'
        '<w:end w:w="120" w:type="dxa"/>'
        "</w:tcMar>"
        '<w:vAlign w:val="center"/>'
        f"{fill}"
        "</w:tcPr>"
        f"{para(str(text), style=style)}"
        "</w:tc>"
    )


def table(headers: list[str], rows: list[list[object]], widths: list[int]) -> str:
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    body = [
        "<w:tbl>",
        "<w:tblPr>",
        '<w:tblW w:w="9360" w:type="dxa"/>',
        '<w:tblInd w:w="120" w:type="dxa"/>',
        '<w:tblLayout w:type="fixed"/>',
        "<w:tblBorders>",
        '<w:top w:val="single" w:sz="4" w:color="B8C2CC"/>',
        '<w:left w:val="single" w:sz="4" w:color="B8C2CC"/>',
        '<w:bottom w:val="single" w:sz="4" w:color="B8C2CC"/>',
        '<w:right w:val="single" w:sz="4" w:color="B8C2CC"/>',
        '<w:insideH w:val="single" w:sz="4" w:color="D6DEE6"/>',
        '<w:insideV w:val="single" w:sz="4" w:color="D6DEE6"/>',
        "</w:tblBorders>",
        "</w:tblPr>",
        f"<w:tblGrid>{grid}</w:tblGrid>",
    ]
    body.append("<w:tr>" + "".join(cell(h, w, header=True) for h, w in zip(headers, widths)) + "</w:tr>")
    for row in rows:
        padded = row + [""] * max(0, len(headers) - len(row))
        body.append("<w:tr>" + "".join(cell(value, w) for value, w in zip(padded, widths)) + "</w:tr>")
    body.append("</w:tbl>")
    return "".join(body)


def section(title: str) -> str:
    return para(title, style="Heading1")


def subsection(title: str) -> str:
    return para(title, style="Heading2")


def numbered(items: list[str]) -> str:
    return "".join(para(item, style="ListParagraph", num_id=1) for item in items)


def document_xml() -> str:
    parts: list[str] = []
    parts.append(para("Finance Reconciliation Automation Logic Reference", style="Title"))
    parts.append(para("Prompt, input, matching, classification, and output logic for the current Streamlit finance automation project.", style="Subtitle"))
    parts.append(label_para("Prepared on: ", "June 5, 2026"))
    parts.append(label_para("Source files reviewed: ", "app.py, build_reconciliation_template.py, disbursement_reconciliation.py, collection_filtering.py, zenith_collection_reconciliation.py"))
    parts.append(para("Important note: The project does not use AI prompts inside the automation logic. The automations are deterministic Python rules. In this document, prompt means the user-facing app prompt, upload instruction, button label, or logic decision shown to the finance user."))

    parts.append(section("1. Global Application Logic"))
    parts.append(numbered([
        "The Streamlit app is split into Collection, Disbursement, and Monthly Summary tabs. Each workflow saves uploaded files into a temporary folder, loads rows from CSV or XLSX, runs deterministic matching or filtering logic, writes an output XLSX workbook, and exposes it with a download button.",
        "CSV files are read with UTF-8-SIG handling. XLSX files are read directly from the workbook XML, so the app avoids external spreadsheet dependencies.",
        "The Column Inspector reads the first row of each uploaded file and compares actual columns against expected columns. It reports Found, Missing, and Extra columns before the user runs the workflow.",
        "Amounts are normalized to cents using Decimal arithmetic and rounded to two decimal places. Displayed amounts are converted back to GHC decimal values.",
        "Most transaction identifiers are normalized by trimming spaces, taking the text before the first slash, removing whitespace, stripping leading apostrophes, uppercasing, and removing leading zeroes for numeric identifiers.",
        "Disbursement identifiers are normalized by removing whitespace, uppercasing, converting scientific-notation integer values to normal integers, and stripping a leading FIDO prefix.",
        "Dates are normalized from Excel serial dates and common text date formats where the workflow requires date matching or cleaned output values.",
    ]))

    parts.append(subsection("Automation Overview"))
    parts.append(table(
        ["Automation", "Required inputs", "Optional/manual inputs", "Output workbook"],
        [
            ["Collection: Nsano Coll Recon", "Mambu, Nsano", "Write Off, Unidentified", "NSANO Coll Recon.xlsx"],
            ["Collection: ITC/Voda Coll Recon", "Mambu, ITC", "Vodafone Collections, Write Off, Unidentified", "ITC/Voda Coll Recon.xlsx"],
            ["Collection: Zenith Coll Recon", "Combined workbook or Mambu + Zenith Bank", "Unidentified, Write Off", "Zenith Coll Recon.xlsx"],
            ["Collection: Write-off Recon", "Write Off", "Nsano, ITC, Zenith Collections, Vodafone Collections", "Write_off_Recon.xlsx"],
            ["Collection: Filtering", "One or more source files", "Automation type selected by user", "Filtered or cleaned workbook"],
            ["Disb: Nsano Wallet vs Mambu", "Mambu Disbursement, NSANO DISB", "None", "Nsano_DISB_Recon.xlsx"],
            ["Disb: Nsano Wallet vs Ledger", "Combined workbook or Mambu + Nsano Disb + Nsano Transfers", "Ledger, unidentified, settlement, reversal amounts", "Nsano_Wallet_vs_Ledger.xlsx"],
            ["Disb: ITC Wallet vs Mambu", "Mambu Disbursement, ITC Wallet", "None", "ITC_Wallet_vs_Mambu.xlsx"],
            ["Disb: ITC Wallet vs Ledger", "Combined workbook or ITC Statement + Debit Transfers + Credit Transfers", "Ledger balance, unidentified amount", "ITC_Wallet_vs_Ledger.xlsx"],
            ["Disb: MTN Manual Disbursement", "Combined workbook or Mambu + MTN MANUAL", "REFUND", "MTN_MANUAL_DISB_Recon.xlsx"],
            ["Disb: Vodafone Manual Disbursement", "Combined workbook or Mambu + VODAFONE MANUAL", "REFUND", "VF_MANUAL_DISB_Recon.xlsx"],
            ["Monthly Summary", "Completed recon results from current session", "No extra uploads", "Monthly_Reconciliation_Summary.xlsx"],
        ],
        [2300, 3000, 2600, 1460],
    ))

    parts.append(section("2. Collection Reconciliation Logic"))
    parts.append(subsection("2.1 Nsano Collection Reconciliation"))
    parts.append(label_para("User-facing prompt: ", "Mambu and Nsano are required. Write Off and Unidentified are optional. Button: Run Nsano Reconciliation."))
    parts.append(table(
        ["Source", "Key column logic", "Amount logic", "Date logic", "Special handling"],
        [
            ["Mambu", "Identifier (Key), else Identifier", "Amount (GHC), else Amount", "Date/Time, else Value Date (Entry Date)", "Loaded as COLLECTION records."],
            ["Nsano", "Ext Debit Ref (Key), else External_Debit_Reference", "Amount (GHC), else Amount_GHC", "Date/Time, else DateTime", "Rows where Result is failed are ignored."],
            ["Write Off", "Identifier (Key), Transaction ID, Transactions Id, or Transactions ID", "Amount (GHC), C, or Amount", "Date/Time, else Date", "Optional fallback source."],
            ["Unidentified", "Key (Identifier), else Identifier", "Amount (GHC), else Amount", "Date/Time, else Value Date (Entry Date)", "Optional fallback source."],
        ],
        [1300, 3000, 1900, 2100, 1060],
    ))
    parts.append(numbered([
        "Mambu vs Nsano compares by direction plus normalized key. If no key exists, the reason is blank reconciliation key. If the key is missing in Nsano, the status is not_found.",
        "If the reference and amount both match, the status is matched. If the reference matches but amount differs, the status is still matched, with Match Reason set to reference matched; amount mismatch and Amount_Difference populated.",
        "Nsano vs Mambu tries sources in this order: Mambu, Write-Off, Unidentified. A found exact amount match receives that source status. A key-only amount mismatch is still assigned to the first source where the key is found, with an amount mismatch reason.",
        "Summary metrics include total, matched, matched exact, amount mismatch, blank key, not found, match rate, source amount, matched amount, and source-specific amounts for Mambu, Write Off, and Unidentified.",
        "The Summary sheet includes Nsano Charge count and amount from the Nsano Charge column.",
    ]))

    parts.append(subsection("2.2 ITC/Vodafone Collection Reconciliation"))
    parts.append(label_para("User-facing prompt: ", "Mambu and ITC are required. Vodafone Collections, Write Off, and Unidentified are optional. Button: Run ITC Reconciliation."))
    parts.append(table(
        ["Source", "Key column logic", "Amount logic", "Date logic", "Special handling"],
        [
            ["Mambu", "Identifier, else Identifier (Key)", "Amount, else Amount (GHC)", "Value Date (Entry Date), else Date/Time", "Normal Mambu collection source."],
            ["ITC", "channel_transaction_id, Channel Transaction ID, or ITC Trans ID (Key)", "amount, Amount, Amount (GHC), or ITC Amount (GHC)", "transaction_date, Transaction Date, or ITC Date", "Narration and source are retained for breakdowns."],
            ["Vodafone Collections", "Receipt No., Receipt No, or Receipt Number", "Paid In, PaidIn, Amount, or Amount (GHC)", "Completion Time, Initiation Time, or Date", "Optional fallback/source comparison."],
            ["Write Off", "Transaction ID, Transactions Id, Transactions ID, or Identifier (Key)", "Amount, Amount (GHC), or C", "Date, else Date/Time", "Optional fallback source."],
            ["Unidentified", "Identifier, else Key (Identifier)", "Amount, else Amount (GHC)", "Value Date (Entry Date), else Date/Time", "Optional fallback source."],
        ],
        [1300, 3100, 2100, 2100, 760],
    ))
    parts.append(numbered([
        "Mambu vs ITC compares Mambu keys to ITC first, then Vodafone Collections. Status values are ITC, Voda_Coll, or not_found.",
        "ITC vs Mambu compares ITC keys to Mambu first, then Vodafone Collections, Unidentified, and Write Off. Status values include mambu, Voda_Coll, unidentified, write-off, upsale, and not_found.",
        "When an ITC row is not found and its narration ends with _3, the row is reclassified as upsale with reason not found; narration _3 classified as upsale.",
        "Vodafone Collections vs Mambu compares Vodafone collection keys to Mambu, then Unidentified, then Write Off. Status values include mambu, unidentified, write off, and not_found.",
        "These ITC/Vodafone comparisons are key-presence comparisons. Amount difference can be shown in comparison rows, but amount mismatch does not create a separate status in this workflow.",
        "ITC narration/source breakdowns are calculated for successful and not_found groups by narration suffixes _1, _2, _3, _4, source containing genpay, and rows with any narration.",
        "The Summary sheet splits ITC fees into Upsales Transaction fees for narration values ending _3, and Commission Charge ITC Payment for the remaining ITC rows.",
    ]))

    parts.append(subsection("2.3 Zenith Collection Reconciliation"))
    parts.append(label_para("User-facing prompt: ", "Mambu and Zenith Bank are required. Matching uses date and amount together. A combined workbook may be uploaded instead of separate files. Button: Run Zenith Reconciliation."))
    parts.append(table(
        ["Source", "Date candidates", "Amount candidates", "Blank-key handling"],
        [
            ["Mambu", "Value Date (Entry Date), Date/Time, Date", "Amount (GHC), Amount", "Kept even when date or amount is blank."],
            ["Zenith Bank", "Create Date (KEY date - DD/MM/YYYY), Create Date, Date", "Credit (KEY amount), Credit, Amount", "Rows with blank date or amount are skipped."],
            ["Unidentified", "Value Date (Entry Date), Date/Time, Date", "Amount (KEY), Amount (GHC), Amount", "Rows with blank date or amount are skipped."],
            ["Write Off", "Date (KEY), Date/Time, Date", "Amount (KEY), Amount (GHC), Amount", "Rows with blank date or amount are skipped."],
        ],
        [1500, 3600, 2700, 1560],
    ))
    parts.append(numbered([
        "The Zenith reconciliation key is normalized_date plus amount_cents, joined as date|amount_cents.",
        "Mambu vs Zenith Bank matches only when both date and amount match. Status values are Matched or Not Found.",
        "Zenith Bank as Source tries Mambu first, then Unidentified, then Write Off. Status values are Mambu, Unidentified, Write Off, or Not Found.",
        "Date normalization supports Excel serial dates, ISO formats, day-first formats, short-year day-first formats, and month-first formats.",
        "The output workbook includes Instructions, Summary, Mambu vs Zenith Bank, Zenith Bank as Source, Mambu, Zenith Bank, Unidentified, and Write Off sheets.",
    ]))

    parts.append(subsection("2.4 Write-off Reconciliation"))
    parts.append(label_para("User-facing prompt: ", "Write Off is the source. Identifiers are matched against Nsano and ITC, then date plus amount is matched against Zenith Collections, then identifier is matched against Vodafone Collections. Button: Run Write-off Reconciliation."))
    parts.append(table(
        ["Priority", "Target source", "Matching logic", "Returned status"],
        [
            ["1", "Nsano Collection", "Ext Debit Ref (Key), else External_Debit_Reference", "Nsano"],
            ["2", "ITC Collection", "channel_transaction_id, Channel Transaction ID, or ITC Trans ID (Key)", "ITC"],
            ["3", "Zenith Collections", "Write-off date and amount matched to Zenith Collections date and amount", "Zenith"],
            ["4", "Vodafone Collections", "Receipt No., Receipt No, or Receipt Number", "Vodafone Collections"],
            ["5", "No target match", "No matching identifier/date+amount in any target source", "Not Found"],
        ],
        [900, 2200, 4300, 1960],
    ))
    parts.append(numbered([
        "Write Off rows are the source rows and are summarized by source amount.",
        "The first matching target wins. For example, if the same identifier exists in both Nsano and ITC, the returned status is Nsano.",
        "The Summary sheet reports count and amount by Nsano, ITC, Zenith, Vodafone Collections, Not Found, and Total.",
        "The output workbook includes Summary, Write Off Recon, Write Off, Nsano, ITC, Zenith Collections, and Vodafone Collections sheets.",
    ]))

    parts.append(section("3. Disbursement Reconciliation Logic"))
    parts.append(subsection("3.1 Nsano Wallet vs Mambu"))
    parts.append(label_para("User-facing prompt: ", "Mambu and NSANO DISB are required. Matching uses column I on both files and ignores Nsano's FIDO prefix. Button: Run Nsano Disbursement Reconciliation."))
    parts.append(table(
        ["Direction", "Source key", "Counterparty", "Status logic"],
        [
            ["Mambu vs Nsano", "Mambu column I / Identifier", "Nsano SendingHse_ID (KEY) / SendingHse_ID", "If key is found in Nsano, status is nsano; otherwise not found."],
            ["Nsano vs Mambu", "Nsano SendingHse_ID stripped of FIDO prefix", "Mambu Identifier", "If key is found in Mambu, status is mambu; otherwise not found."],
        ],
        [1800, 2500, 2700, 2360],
    ))
    parts.append(numbered([
        "The match is key-only. Amount is displayed and summarized, but amount mismatch does not create a separate status.",
        "Mambu amount is taken from column E / Amount (GHC) / Amount. Nsano amount is taken from Amount_GHC, Amount (GHC), or Amount.",
    ]))

    parts.append(subsection("3.2 ITC Wallet vs Mambu"))
    parts.append(label_para("User-facing prompt: ", "Mambu and ITC Wallet are required. Matching uses ITC Wallet column G and Mambu column I. Button: Run ITC Wallet vs Mambu Reconciliation."))
    parts.append(table(
        ["Direction", "Source key", "Counterparty", "Status logic"],
        [
            ["Mambu vs ITC Wallet", "Mambu Identifier (KEY) / Identifier", "ITC thirdparty_id (KEY) / thirdparty_id", "If key is found in ITC, status is ITC; otherwise not_found."],
            ["ITC Wallet vs Mambu", "ITC thirdparty_id", "Mambu Identifier", "If key is found in Mambu, status is mambu. Otherwise narration suffix decides special not-found status."],
        ],
        [1800, 2600, 2600, 2360],
    ))
    parts.append(table(
        ["ITC narration ending", "Status when no Mambu match"],
        [
            ["_2", "referral bonus"],
            ["_3", "upsales refound"],
            ["_6", "savings reward"],
            ["Anything else", "not_found"],
        ],
        [3000, 6360],
    ))

    parts.append(subsection("3.3 ITC Wallet vs Ledger"))
    parts.append(label_para("User-facing prompt: ", "Use combined ITC workbook, or upload ITC Statement, ITC Debit Transfers, and ITC Credit Transfers separately. Ledger balance and unidentified are entered manually. Button: Run ITC Wallet vs Ledger Reconciliation."))
    parts.append(numbered([
        "Credit transfers are classified from the credit transfer narration. If compact narration contains settlement, status is settlement. If it contains prepaidtransactionreversal or prepaidtr, status is prepaid_reversal. If it contains reversalof or starts with reversal, status is reversal. Otherwise status is transfer_to_wallet.",
        "Debit transfers are indexed by third_party_transaction_id and checked against statement processor_transaction_id and credit transfer third_party_transaction_id.",
        "If a statement row is found, the statement narration suffix classifies the debit: _1 or _4 means disbursement, _2 means referral_award, _3 means upsales_refund, _6 means savings, and any other suffix defaults to disbursement.",
        "If no statement row is found and the debit row text contains withdrawal, the debit is transfers_from_bank. If no statement row is found but a matching credit transfer exists, the debit is credit transfer (prepaid_reversal). Otherwise it is not_found.",
    ]))
    parts.append(table(
        ["Formula item", "Logic"],
        [
            ["Unidentified", "Positive manual unidentified is added to credit/available funds. Negative manual unidentified is treated as a positive debit."],
            ["Available funds", "ledger_balance + settlement + reversal + transfer_to_wallet + unidentified_credit"],
            ["Total debit", "transfer_to_bank + referral_awards + savings + upsales_refund + disbursement + unidentified_debit"],
            ["Wallet statement balance", "available_funds - total_debit"],
        ],
        [2600, 6760],
    ))

    parts.append(subsection("3.4 Nsano Wallet vs Ledger"))
    parts.append(label_para("User-facing prompt: ", "Use combined Nsano workbook, or upload Mambu Disbursement, Nsano Disbursement, and Nsano Transfers separately. Ledger, unidentified, recovery from write off, and reversal are entered manually. Button: Run Nsano Wallet vs Ledger Reconciliation."))
    parts.append(table(
        ["Transfer purpose text", "Transfer Category"],
        [
            ["Contains topup and collection", "top_up_through_collections"],
            ["Contains topup and stanbic", "bank_to_wallet"],
            ["Contains reversaladjustment", "reversal_adjustment"],
            ["Contains transfertobank", "transfer_to_bank"],
            ["Anything else", "other"],
        ],
        [4200, 5160],
    ))
    parts.append(table(
        ["Formula item", "Logic"],
        [
            ["Unidentified", "Positive unidentified is added to available funds. Negative unidentified is added to total debit as a positive value."],
            ["Reversal", "Positive reversal is added to ledger balance. Negative reversal is added to available funds as a positive value."],
            ["Starting ledger balance", "ledger_balance + reversal_ledger"],
            ["Available funds", "starting_ledger_balance + topup_collections + bank_to_wallet + reversal_adjustment + recovery_from_write_off + unidentified_available_funds + reversal_available_funds"],
            ["Disbursement", "Sum of Mambu disbursement records."],
            ["Total debit", "transfer_to_bank + disbursement + charges + unidentified_debit"],
            ["Wallet statement balance", "available_funds - total_debit"],
        ],
        [2600, 6760],
    ))
    parts.append(para("Nsano Disbursement statement rows are retained in the output workbook for source/audit visibility. The wallet balance formula uses Mambu disbursement totals and Nsano transfer classifications."))

    parts.append(subsection("3.5 MTN Manual Disbursement"))
    parts.append(label_para("User-facing prompt: ", "Mambu and MTN MANUAL are required. Refund is used as fallback when MTN Manual is not found in Mambu. A combined workbook may be used. Button: Run MTN Manual Disbursement Reconciliation."))
    parts.append(table(
        ["Source", "Key logic", "Match direction", "Status logic"],
        [
            ["Mambu", "Column I / Identifier", "Mambu vs MTN MANUAL", "If key exists in MTN Manual, status is Matched; otherwise Not Found."],
            ["MTN MANUAL", "Column A / Id / ID", "MTN MANUAL as Source", "If key exists in Mambu, status is Mambu. Else if key exists in REFUND, status is Refund. Else Not Found."],
            ["REFUND", "Column D / TRANSACTION ID / Transaction ID", "Fallback source", "Used only for MTN MANUAL as Source fallback."],
        ],
        [1400, 2600, 2400, 2960],
    ))

    parts.append(subsection("3.6 Vodafone Manual Disbursement"))
    parts.append(label_para("User-facing prompt: ", "Mambu and VODAFONE MANUAL are required. Refund is used as fallback when Vodafone Manual is not found in Mambu. A combined workbook may be used. Button: Run Vodafone Manual Disbursement Reconciliation."))
    parts.append(table(
        ["Source", "Key logic", "Match direction", "Status logic"],
        [
            ["Mambu", "Column I / Identifier", "Mambu vs VODAFONE MANUAL", "If key exists in Vodafone Manual, status is Matched; otherwise Not Found."],
            ["VODAFONE MANUAL", "Column A / Id / ID", "VODAFONE MANUAL as Source", "If key exists in Mambu, status is Mambu. Else if key exists in REFUND, status is Refund. Else Not Found."],
            ["REFUND", "Column D / TRANSACTION ID / Transaction ID", "Fallback source", "Used only for VODAFONE MANUAL as Source fallback."],
        ],
        [1400, 2600, 2400, 2960],
    ))

    parts.append(section("4. Filtering Automation Logic"))
    parts.append(label_para("User-facing prompt: ", "Upload one or more files. Files can be XLSX or CSV and will be combined before filtering. Button text changes based on the selected filtering automation."))
    parts.append(table(
        ["Filtering automation", "Rule", "Workbook behavior"],
        [
            ["Mambu Collection", "Channel matches one of the checked ITC or Vodafone collection channels from the finance channel list", "Sheet: ITC Vodafone Collection."],
            ["Mambu Collection", "Channel matches the checked ITC collection channels, excluding Vodafone Coll. Account - vodafone", "Sheet: ITC Collection."],
            ["Mambu Collection", "Channel equals Vodafone Coll. Account - vodafone", "Sheet: Vodafone Collection."],
            ["Mambu Collection", "Channel equals Zenith Coll. Account - 6010159660", "Sheet: Zenith Collections."],
            ["ITC Collection", "transaction_type equals inflow", "Sheet: Inflow. No All Filtered sheet."],
            ["ITC Collection", "transaction_type equals outflow", "Sheet: Outflow. No All Filtered sheet."],
            ["Nsano Disb and Collections", "Result equals Successful and Type equals W2A", "Sheet: Successful W2A. No All Filtered sheet."],
            ["Nsano Disb and Collections", "Result equals Successful and Type equals A2W", "Sheet: Successful A2W. No All Filtered sheet."],
            ["Mambu Disbursement", "Channel matches one of the configured ITC disbursement channels", "Sheet: ITC Disbursement. No All Filtered sheet."],
            ["Mambu Disbursement", "Channel equals Nsano - Client Disbursement", "Sheet: Nsano Client Disb. No All Filtered sheet."],
            ["Mambu Disbursement", "Channel equals Maxbuy - Client Disbursement", "Sheet: Maxbuy Client Disbursement. No All Filtered sheet."],
            ["Voda Collection Cleanup", "First detects the real header row where Receipt No. appears, removes the 8th column, skips blank/repeated header rows, and keeps all cleaned rows", "Sheet: Cleaned Voda Coll. No All Filtered sheet."],
        ],
        [1900, 4600, 2860],
    ))
    parts.append(para("Filter amounts are summed from the first available amount-like column in this order: Amount, amount, Amount (GHC), Amount_GHC, Paid In, Principal Amount, net_amount."))

    parts.append(section("5. Monthly Summary Logic"))
    parts.append(label_para("User-facing prompt: ", "Run the collection and disbursement recons first, then generate one combined monthly summary. Button: Generate Monthly Summary."))
    parts.append(para(
        "The monthly summary does not ask for the monthly files again. It reads the completed reconciliation result objects already saved in Streamlit session state and excludes each workflow's raw output bytes."
    ))
    parts.append(table(
        ["Sheet", "Content"],
        [
            ["Run Status", "Shows every supported collection, disbursement, and wallet workflow as Ready or Not run, with main row count, amount, and notes."],
            ["Collection Summary", "Combines Nsano Coll Recon, ITC/Voda Coll Recon, Zenith Coll Recon, and Write-off Recon counts, amounts, match rates, matched amounts, and not-found amounts."],
            ["Disbursement Summary", "Combines Nsano Wallet vs Mambu, ITC Wallet vs Mambu, MTN Manual, and Vodafone Manual counts, amounts, match rates, matched amounts, and not-found amounts."],
            ["Wallet Summary", "Combines Nsano Wallet vs Ledger and ITC Wallet vs Ledger wallet categories, transaction counts where available, manual ledger inputs, debit totals, and wallet statement balances."],
        ],
        [2200, 7360],
    ))
    parts.append(para(
        "If a reconciliation is run again after the monthly summary is generated, the app warns that the master summary should be regenerated before download."
    ))

    parts.append(section("6. Status And Review Notes"))
    parts.append(numbered([
        "Reference-key workflows generally normalize keys before matching. A row can look different in the raw file and still match if the normalized key is the same.",
        "Zenith collection is the main exception: it has no stable transaction reference, so it uses date plus amount as the key.",
        "Nsano collection failed rows are excluded from collection reconciliation. Nsano disbursement rows are not filtered by failed result in the current disbursement loader.",
        "ITC/Vodafone collection and disbursement workflows use narration suffixes to classify certain unmatched ITC rows into operational buckets such as upsale, referral bonus, upsales refund, or savings.",
        "Manual wallet-vs-ledger values are part of the calculation and should be captured from finance-approved month-end balances before running those templates.",
        "The generated XLSX workbooks retain source sheets and comparison sheets so finance can review both the summary and row-level evidence.",
    ]))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14">'
        "<w:body>"
        + "".join(parts)
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:color w:val="000000"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="40"/><w:color w:val="0B2545"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:qFormat/><w:pPr><w:spacing w:after="180"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:color w:val="555555"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="200"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="32"/><w:color w:val="2E74B5"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="26"/><w:color w:val="2E74B5"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:ind w:left="540" w:hanging="270"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TableHeader"><w:name w:val="Table Header"/><w:basedOn w:val="TableText"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="0B2545"/></w:rPr></w:style>
</w:styles>"""


def numbering_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def document_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""


def settings_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:compat/>
</w:settings>"""


def core_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Finance Reconciliation Automation Logic Reference</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-06-05T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-06-05T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""


def app_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex OOXML Builder</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>Fido Finance</Company>
</Properties>"""


def build() -> None:
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml())
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("word/document.xml", document_xml())
        zf.writestr("word/styles.xml", styles_xml())
        zf.writestr("word/numbering.xml", numbering_xml())
        zf.writestr("word/settings.xml", settings_xml())
        zf.writestr("word/_rels/document.xml.rels", document_rels_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml())


if __name__ == "__main__":
    build()
    print(f"Wrote {OUT}")
