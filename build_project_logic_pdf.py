#!/usr/bin/env python3
"""Build plain-English documentation for the Finance automation project."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


OUT = Path("output/pdf/Finance_Automation_Project_Logic_Documentation.pdf")

NAVY = colors.HexColor("#14253D")
BLUE = colors.HexColor("#176B87")
TEAL = colors.HexColor("#2A9D8F")
PALE = colors.HexColor("#EAF3F5")
LIGHT = colors.HexColor("#F5F7F9")
GOLD = colors.HexColor("#E9C46A")
RED = colors.HexColor("#B7443E")
GREY = colors.HexColor("#5C6875")
WHITE = colors.white

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(
    name="CoverTitle", parent=styles["Title"], fontName="Helvetica-Bold",
    fontSize=25, leading=30, textColor=NAVY, alignment=TA_CENTER, spaceAfter=12,
))
styles.add(ParagraphStyle(
    name="CoverSub", parent=styles["Normal"], fontName="Helvetica",
    fontSize=11, leading=17, textColor=GREY, alignment=TA_CENTER, spaceAfter=8,
))
styles.add(ParagraphStyle(
    name="H1x", parent=styles["Heading1"], fontName="Helvetica-Bold",
    fontSize=17, leading=21, textColor=NAVY, spaceBefore=10, spaceAfter=8,
))
styles.add(ParagraphStyle(
    name="H2x", parent=styles["Heading2"], fontName="Helvetica-Bold",
    fontSize=12.5, leading=16, textColor=BLUE, spaceBefore=9, spaceAfter=5,
))
styles.add(ParagraphStyle(
    name="H3x", parent=styles["Heading3"], fontName="Helvetica-Bold",
    fontSize=10.5, leading=14, textColor=NAVY, spaceBefore=7, spaceAfter=4,
))
styles.add(ParagraphStyle(
    name="Bodyx", parent=styles["BodyText"], fontName="Helvetica",
    fontSize=8.7, leading=12.4, textColor=colors.HexColor("#263442"),
    spaceAfter=5, alignment=TA_LEFT,
))
styles.add(ParagraphStyle(
    name="Smallx", parent=styles["BodyText"], fontName="Helvetica",
    fontSize=7.4, leading=10.2, textColor=colors.HexColor("#263442"),
))
styles.add(ParagraphStyle(
    name="Bulletx", parent=styles["BodyText"], fontName="Helvetica",
    fontSize=8.5, leading=12, leftIndent=12, firstLineIndent=-7,
    bulletIndent=4, spaceAfter=3, textColor=colors.HexColor("#263442"),
))
styles.add(ParagraphStyle(
    name="Callout", parent=styles["BodyText"], fontName="Helvetica",
    fontSize=8.8, leading=12.5, leftIndent=9, rightIndent=9,
    borderColor=TEAL, borderWidth=0.8, borderPadding=8,
    backColor=PALE, spaceBefore=4, spaceAfter=7,
))
styles.add(ParagraphStyle(
    name="Formula", parent=styles["BodyText"], fontName="Courier",
    fontSize=7.8, leading=11, leftIndent=8, rightIndent=8,
    borderColor=colors.HexColor("#CAD3DC"), borderWidth=0.5,
    borderPadding=6, backColor=LIGHT, spaceAfter=5,
))
styles.add(ParagraphStyle(
    name="TableHead", parent=styles["BodyText"], fontName="Helvetica-Bold",
    fontSize=7.5, leading=9.5, textColor=WHITE,
))
styles.add(ParagraphStyle(
    name="TableCell", parent=styles["BodyText"], fontName="Helvetica",
    fontSize=7.2, leading=9.6, textColor=colors.HexColor("#24313D"),
))


def p(text: str, style: str = "Bodyx"):
    return Paragraph(text, styles[style])


def h1(text: str):
    return Paragraph(text, styles["H1x"])


def h2(text: str):
    return Paragraph(text, styles["H2x"])


def h3(text: str):
    return Paragraph(text, styles["H3x"])


def bullet(text: str):
    return Paragraph(f"• {text}", styles["Bulletx"])


def formula(text: str):
    return Paragraph(text, styles["Formula"])


def table(headers, rows, widths=None):
    data = [[p(str(x), "TableHead") for x in headers]]
    data.extend([[p(str(x), "TableCell") for x in row] for row in rows])
    t = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD4DC")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT]),
    ]))
    return t


def workflow(title, prompt, inputs, logic, outputs, notes=None):
    parts = [h2(title), p(f"<b>User-facing prompt:</b> {prompt}")]
    parts.append(h3("Inputs"))
    parts.extend(bullet(x) for x in inputs)
    parts.append(h3("Logic"))
    parts.extend(bullet(x) for x in logic)
    parts.append(h3("Outputs"))
    parts.extend(bullet(x) for x in outputs)
    if notes:
        parts.append(h3("Important notes"))
        parts.extend(bullet(x) for x in notes)
    return parts


class LogicDoc(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(
            filename, pagesize=A4,
            leftMargin=18 * mm, rightMargin=18 * mm,
            topMargin=18 * mm, bottomMargin=17 * mm,
            title="Finance Automation Project Logic Documentation",
            author="Project logic review",
        )
        frame = Frame(
            self.leftMargin, self.bottomMargin, self.width, self.height,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        self.addPageTemplates(PageTemplate(id="main", frames=frame, onPage=self._page))

    def _page(self, canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D7DEE5"))
        canvas.setLineWidth(0.5)
        canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GREY)
        canvas.drawString(18 * mm, 9.5 * mm, "Finance Automation - Logic Reference")
        canvas.drawRightString(A4[0] - 18 * mm, 9.5 * mm, f"Page {doc.page}")
        canvas.restoreState()


def story():
    s = []
    s += [
        Spacer(1, 35 * mm),
        p("PROJECT LOGIC REFERENCE", "CoverSub"),
        p("Finance Automation", "CoverTitle"),
        p("A plain-English guide to the application tabs, prompts, filtering rules, reconciliation decisions, wallet calculations, summaries, and archive behaviour.", "CoverSub"),
        Spacer(1, 14 * mm),
        Table([[""], [""]], colWidths=[135 * mm], rowHeights=[2 * mm, 2 * mm],
              style=TableStyle([("BACKGROUND", (0, 0), (-1, 0), TEAL), ("BACKGROUND", (0, 1), (-1, 1), GOLD)])),
        Spacer(1, 14 * mm),
        p(f"Reviewed from the project source code and tests<br/>Document date: {date.today().isoformat()}", "CoverSub"),
        Spacer(1, 28 * mm),
        p("<b>Audience:</b> Finance users, reviewers, operators, developers, and auditors.", "Callout"),
        PageBreak(),
    ]

    s += [
        h1("1. Purpose and scope"),
        p("This document explains the complete finance automation project in clear English. It describes what each major tab does, what the user is asked to provide, how the code makes decisions, and what appears in each output workbook."),
        p("<b>Meaning of “prompt” in this project:</b> the application does not use artificial-intelligence prompts to decide financial results. The prompts are screen instructions, upload labels, input fields, warnings, buttons, and captions. Every financial decision is made by fixed Python rules."),
        p("<b>Files reviewed:</b> app.py, collection_filtering.py, build_reconciliation_template.py, zenith_collection_reconciliation.py, disbursement_reconciliation.py, collections_master_summary.py, monthly_summary.py, master_ledger.py, recon_archive.py, build_reconciliation_template.py, helper/build scripts, Google Drive setup notes, requirements, and tests.", "Callout"),
        h2("1.1 Major application tabs"),
        table(
            ["Major tab", "Minor sections", "Purpose"],
            [
                ["Filtering", "Collection filters; disbursement filters", "Splits or cleans uploaded operational files before reconciliation."],
                ["Collection", "Nsano, ITC/Vodafone, Zenith, Write-Off; master collection run", "Compares collection transactions between Mambu and external collection sources."],
                ["Disbursement", "Nsano, ITC, MTN Manual, Vodafone Manual; master disbursement run", "Compares Mambu disbursements with wallet/manual sources and refunds."],
                ["Ledger", "Nsano Collections, Nsano Disbursement, ITC, Voda vs Ledger, Zenith Wallet vs Ledger", "Builds wallet-to-ledger calculations from transactions and manual month-end balances."],
                ["Monthly Summary", "Combined summary", "Combines completed results held in the current Streamlit session."],
                ["Recon Archive", "Search, filter, download, ZIP, delete, refresh", "Stores and retrieves versioned workbooks from Google Drive."],
            ],
            [27 * mm, 47 * mm, 96 * mm],
        ),
        h2("1.2 End-to-end flow"),
        formula("Upload or select archived source -> validate and normalize -> filter/classify -> match -> calculate totals -> write XLSX -> keep result in session -> save to Drive archive -> build master/monthly summary"),
        bullet("CSV and XLSX are accepted in most source workflows. Multiple chunks are combined where the screen allows multiple uploads."),
        bullet("The output format is XLSX. This PDF documents the logic but is not itself a transaction result."),
        bullet("Temporary files are used only while a run is executing. Archive persistence is Google Drive-only."),
        h2("1.3 Common normalization rules"),
        table(
            ["Area", "Rule"],
            [
                ["Text", "Leading and trailing spaces are removed; comparisons are generally case-insensitive."],
                ["Identifiers", "Punctuation, prefixes, leading zeros, and provider-specific decorations are cleaned by the workflow before matching."],
                ["Nsano disbursement", "The FIDO prefix is ignored when the disbursement identifier is normalized."],
                ["Amounts", "Text amounts are converted to cents or Decimal values. Commas are removed. Invalid or blank values generally become zero or blank, depending on context."],
                ["Dates", "Excel serial dates and common date text are converted to normalized date values where date matching is required."],
                ["Phone values", "Phone formatting is normalized for display or comparison when used."],
                ["Headers", "Several expected header alternatives are accepted. Some loaders can fall back to fixed column positions."],
            ],
            [35 * mm, 135 * mm],
        ),
        PageBreak(),
    ]

    s += [h1("2. Filtering")]
    s += [
        p("Filtering does not reconcile transactions. It prepares source data by placing rows into named sheets. A row may qualify for more than one filter when the rules overlap. The result counters show input rows, unique filtered rows, row counts by filter, and summed amounts."),
        h2("2.1 Mambu collection filtering"),
        p("<b>Prompt:</b> Upload one or more Mambu collection files. CSV and XLSX chunks are combined. Run the selected filtering automation."),
        table(
            ["Filter / output sheet", "Exact rule"],
            [
                ["ITC / Vodafone Collection", "Channel equals any configured ITC collection channel or Vodafone collection channel."],
                ["ITC Collection", "Channel equals a configured ITC collection channel, excluding “Vodafone Coll. Account - vodafone”."],
                ["Vodafone Collection", "Channel equals “Vodafone Coll. Account - vodafone”."],
                ["Nsano Client Collection", "Channel equals “Nsano - Client Collections” or “Nsano - Client Collection”."],
                ["Zenith Collections", "Channel equals “Zenith Coll. Account - 6010159660”."],
            ],
            [48 * mm, 122 * mm],
        ),
        h3("Configured ITC/Vodafone collection channels"),
        p("AirtelTigo Coll. Account - ITC Fido; Business Loans (AirtelTigo) - Client Collection - ITC; Business Loans (MTN) - Client Collection - ITC; Business Loans (Vodafone) - Client Collection - ITC; ITC Collections Account - GIP; MTN Coll. Account - ITC Fido; Vodafone Coll. Account - ITC Fido; Vodafone Coll. Account - vodafone; Vodafone Collections Account - GIP.", "Smallx"),
        h2("2.2 ITC collection filtering"),
        bullet("<b>Inflow:</b> transaction_type equals “inflow”."),
        bullet("<b>Outflow:</b> transaction_type equals “outflow”."),
        h2("2.3 Nsano collection filtering"),
        bullet("<b>Successful W2A:</b> Result equals “Successful” and Type equals “W2A”."),
        bullet("<b>Successful A2W:</b> Result equals “Successful” and Type equals “A2W”."),
        h2("2.4 Mambu disbursement filtering"),
        table(
            ["Filter / output sheet", "Rule"],
            [
                ["ITC Disbursement", "Channel equals one of the six configured ITC automatic/business-loan disbursement channels."],
                ["Nsano Client Disbursement", "Channel equals “Nsano - Client Disbursement”."],
                ["Maxbuy Client Disbursement", "Channel equals “Maxbuy - Client Disbursement”."],
                ["MTN Manual Disbursement", "Channel text contains MTN, Manual, and Disbursement."],
                ["Vodafone Manual Disbursement", "Channel text contains Vodafone, Manual, and Disbursement."],
            ],
            [50 * mm, 120 * mm],
        ),
        h3("Configured ITC disbursement channels"),
        p("AirtelTigo Disb. Account - Auto - 0266930184; Business Loans (AirtelTigo) - Client Disbursement - ITC; Business Loans (MTN) - Client Disbursement - ITC; Business Loans (Vodafone) - Client Disbursement - ITC; MTN Disb. Account - Auto - ITC; Vodafone Disb. Account - Auto - ITC.", "Smallx"),
        h2("2.5 Vodafone collection cleanup"),
        bullet("The loader scans the first 25 rows for a header row containing “Receipt No.” so raw exports with an account-information block can be accepted."),
        bullet("Blank rows and repeated header rows are ignored. All discovered source columns are retained."),
        bullet("A cleaned file is expected to contain: Receipt No., Completion Time, Initiation Time, Details, Transaction Status, Currency, Paid In, Withdrawn, Balance, Reason Type, and Opposite Party."),
        bullet("The output sheet is “Cleaned Voda Coll”. CSV files can be normalized in place before later use."),
        h2("2.6 Filtering amount logic"),
        p("When the screen displays a filter amount, the first available value is taken from Amount, amount, Amount (GHC), Amount_GHC, Paid In, Principal Amount, or net_amount. Invalid numbers count as zero."),
        p("The “unique filtered rows” total removes duplicate row objects created by overlapping filters. It does not perform business-key deduplication.", "Callout"),
    ]

    s += [h1("3. Collection reconciliation")]
    s += workflow(
        "3.1 Nsano collections vs Mambu",
        "Mambu and Nsano are required. Write Off and Unidentified are optional. Run Nsano Reconciliation.",
        [
            "Mambu collection rows and Nsano collection rows.",
            "Optional Write-Off and Unidentified rows, either uploaded separately or read from an accepted combined workbook.",
        ],
        [
            "Mambu records are divided into normal Mambu, Write-Off, and Unidentified groups by descriptive text.",
            "The matching index uses transaction direction plus the normalized reconciliation identifier.",
            "Mambu -> Nsano: blank keys are Not Found. A key not present in Nsano is Not Found. A key and amount match is Matched. A matching key with a different amount is still Matched, but the reason says “amount mismatch” and the difference is recorded.",
            "Nsano -> sources checks in this order: normal Mambu, Write-Off, then Unidentified. The first source with the key wins. An exact amount match is preferred within that source.",
            "Failed Nsano collection rows are excluded by the collection loader and counted as ignored.",
        ],
        [
            "Summary, source-data, comparison, exception, and not-found evidence sheets.",
            "Counts, source amounts, matched amounts, amount-mismatch totals, blank-key totals, not-found totals, match rates, and source-specific totals.",
        ],
        ["Matching does not consume a counterparty row. Repeated source keys can point to the same first candidate."],
    )
    s += workflow(
        "3.2 ITC and Vodafone collections vs Mambu",
        "Mambu and ITC are required. Vodafone Collections, Write Off, and Unidentified are optional. Run ITC Reconciliation.",
        [
            "Mambu, ITC collection, optional cleaned Vodafone collection, Write-Off, and Unidentified data.",
        ],
        [
            "This workflow matches by normalized key presence, not by an exact amount requirement.",
            "Mambu -> external sources checks ITC first and Vodafone Collections second.",
            "ITC -> sources checks Mambu, Vodafone Collections, Unidentified, then Write-Off. First source found wins.",
            "An unmatched ITC narration ending “_3” becomes “upsale” instead of ordinary not_found.",
            "Vodafone -> sources checks Mambu, Unidentified, then Write-Off.",
            "Amount differences may be displayed, but an amount mismatch does not create a different match status.",
            "Vodafone “Pay Bill Charge” rows are counted and their absolute Withdrawn amounts are summed.",
            "ITC fee reporting separates narration ending “_3” as Upsales Transaction fees; remaining ITC rows are Commission Charge ITC Payment.",
        ],
        [
            "Bidirectional comparison sheets, raw/source sheets, summary metrics, narration breakdowns, charge totals, and not-found details.",
        ],
    )
    s += workflow(
        "3.3 Zenith collections vs Mambu",
        "Mambu and Zenith Bank are required. Matching uses date and amount together. Optional Unidentified and Write Off sources may be included. Run Zenith Reconciliation.",
        [
            "Mambu collection rows, Zenith bank rows, and optional Unidentified and Write-Off rows.",
        ],
        [
            "Zenith does not have a reliable common transaction identifier, so the key is normalized date plus amount in cents.",
            "Mambu -> Zenith matches only when date and amount are both equal.",
            "Zenith -> sources checks Mambu, Unidentified, and Write-Off using the same date-and-amount key.",
            "Statuses identify the source that matched; unmatched rows are Not Found.",
        ],
        [
            "Summary, source sheets, Mambu-vs-Zenith comparison, Zenith-vs-sources comparison, and not-found evidence.",
        ],
        ["Different transactions on the same date for the same amount share the same key; the first indexed candidate is used."],
    )
    s += workflow(
        "3.4 Write-Off reconciliation",
        "Write Off is the source. It is checked against Nsano, ITC, Zenith Collections, and Vodafone Collections. Run Write-off Reconciliation.",
        [
            "Write-Off rows; optional/available Nsano, ITC, Zenith, and Vodafone collection sources.",
        ],
        [
            "Channel text routes a Write-Off row to Nsano, ITC, Zenith, or Vodafone Collections. Matching is case-insensitive after punctuation is replaced by spaces; Vodafone also accepts the word Voda.",
            "Nsano, ITC, and Vodafone use the normalized Write-Off identifier. All identifier-wallet matches are collected so cross-wallet duplicates remain visible.",
            "If the routed wallet is among the matches, that wallet is allocated. Other matches are reported as Multiple Wallet Matches - Routed by Channel.",
            "If there is one match but it disagrees with the channel route, the row is allocated to the found wallet and flagged Channel Mismatch. Multiple matches with no usable route are Ambiguous - Multiple Wallet Matches.",
            "Zenith-routed rows ignore all identifier matches and compare only normalized date plus amount. An exact date-and-amount match allocates to Zenith. A same-date/different-amount candidate also allocates to Zenith but is flagged Amount Mismatch; same-amount/different-date remains Not Found with a date-mismatch investigation.",
            "For duplicate Vodafone receipt rows, the first non-zero amount candidate is preferred so the Paid In collection is used instead of a zero-amount/charge row.",
        ],
        [
            "A consolidated comparison containing Routed Wallet, Candidate Wallets, Match Status, Investigation Status, Zenith date/amount checks, Write-Off amount, Wallet amount, and Amount Variance; plus counts and control totals by allocated wallet.",
        ],
    )
    s += [
        PageBreak(),
        h2("3.5 Collection master run"),
        p("The master collection screen accepts a combined workbook and/or individual sources, shows sheet and column diagnostics, runs the applicable collection workflows, and builds one dashboard-style master summary."),
        bullet("It shows source totals, wallet totals, Mambu totals, exceptions, write-offs, unidentified items, charges, settlements, daily wallet summaries, and grouped not-found details."),
        bullet("Not-found detail sheets keep both the raw reconciliation identifier and the cleaned key so a reviewer can see why matching succeeded or failed."),
        bullet("Daily summaries are separated into Nsano, ITC, and Vodafone sheets and include totals."),
        bullet("The app warns about missing expected sheets, data issues, ignored Nsano failures, and zero key overlap where that likely indicates the wrong Mambu sheet or wrong source file."),
    ]

    s += [h1("4. Disbursement reconciliation")]
    s += workflow(
        "4.1 Nsano disbursement vs Mambu",
        "Mambu and NSANO DISB are required. Matching uses column I on both files and ignores Nsano's FIDO prefix. Run Nsano Disbursement Reconciliation.",
        ["Mambu disbursement rows and Nsano disbursement rows."],
        [
            "Identifiers are normalized, including removal of the provider FIDO decoration.",
            "Mambu -> Nsano and Nsano -> Mambu both use key presence.",
            "The match is not rejected when amounts differ. Amounts remain visible for review.",
            "Statuses are mambu/nsano or not found, depending on the comparison direction.",
        ],
        ["Summary, instructions, both source sheets, both comparison sheets, and not-found details."],
    )
    s += workflow(
        "4.2 ITC Wallet vs Mambu",
        "Mambu and ITC Wallet are required. Matching uses ITC Wallet column G and Mambu column I. Run ITC Wallet vs Mambu Reconciliation.",
        ["Mambu disbursement and ITC wallet rows."],
        [
            "Both directions use the normalized identifier as the key.",
            "When an ITC row does not match Mambu, narration suffix “_2” becomes referral bonus, “_3” becomes upsales refound, and “_6” becomes savings reward.",
            "Other unmatched ITC rows are not_found.",
            "Amount differences are evidence, not a separate match status.",
        ],
        ["Summary with narration classifications, source sheets, comparisons, and detailed exceptions."],
    )
    s += workflow(
        "4.3 MTN Manual disbursement",
        "Mambu and MTN MANUAL are required. Refund is a fallback when MTN Manual is not found in Mambu. Run MTN Manual Disbursement Reconciliation.",
        ["Mambu, MTN Manual, and optional Refund data."],
        [
            "Mambu -> MTN Manual checks the normalized transaction key.",
            "MTN Manual -> sources checks Mambu first, then Refund.",
            "The first source with the key wins. Status is Mambu, Refund, or Not Found.",
        ],
        ["Summary, source data, two-way comparisons, refund attribution, and not-found evidence."],
    )
    s += workflow(
        "4.4 Vodafone Manual disbursement",
        "Mambu and VODAFONE MANUAL are required. Refund is a fallback when Vodafone Manual is not found in Mambu. Run Vodafone Manual Disbursement Reconciliation.",
        ["Mambu, Vodafone Manual, and optional Refund data."],
        [
            "The raw Vodafone loader scans up to 25 rows to find the correct header.",
            "Mambu -> Vodafone Manual checks the normalized key.",
            "Vodafone Manual -> sources checks Mambu first and Refund second.",
            "Status is Mambu, Refund, or Not Found.",
        ],
        ["Summary, source sheets, comparisons, refund totals, and not-found details."],
    )
    s += [
        h2("4.5 Disbursement master run"),
        bullet("The master run can load workflow-specific Mambu sheets from a combined workbook."),
        bullet("It measures source-to-Mambu key overlap and warns when overlap is zero, which helps catch wrong sheets and wrong months."),
        bullet("The master summary includes wallet-source totals, Mambu-source totals, refunds, not-found values, daily analysis where available, and grouped detail sheets."),
        bullet("Completed workflow outputs are placed in Streamlit session state and become available to summary and archive functions."),
        PageBreak(),
    ]

    s += [h1("5. Wallet and ledger logic")]
    s += [
        p("Ledger workflows combine transaction classifications with manual month-end inputs. Positive and negative manual values can intentionally go to different sides of the calculation. Users should enter finance-approved balances."),
        h2("5.1 Nsano transfer classification"),
        table(
            ["Purpose text after compacting", "Category"],
            [
                ["Contains “topup” and “collection”", "top_up_through_collections"],
                ["Contains “topup” and “stanbic”", "bank_to_wallet"],
                ["Contains “reversaladjustment”", "reversal_adjustment"],
                ["Contains “transfertobank” or “transferbank”", "transfer_to_bank"],
                ["Anything else", "other"],
            ],
            [90 * mm, 80 * mm],
        ),
        h2("5.2 Nsano disbursement wallet vs ledger"),
        p("<b>Prompt:</b> Enter ledger balance, delayed/unidentified amount, recovery from write-off/settlement, and reversal. Upload a combined workbook or Mambu Disbursement, Nsano Disbursement, and Nsano Transfers separately."),
        bullet("Top-up through collections, bank-to-wallet, reversal adjustment, and transfer-to-bank totals come from classified Nsano transfer rows."),
        bullet("Disbursement is the sum of selected Mambu disbursement records."),
        bullet("A positive delayed amount is available funds; a negative delayed amount becomes an absolute debit."),
        bullet("A positive reversal is added to the starting ledger balance. A negative reversal is added to available funds as a positive amount."),
        formula("Starting ledger balance = ledger balance + positive reversal"),
        formula("Available funds = starting ledger balance + top-up collections + bank to wallet + reversal adjustment + recovery from write-off + positive delayed amount + absolute negative reversal"),
        formula("Total debit = transfer to bank + disbursement + charges (when supplied) + absolute negative delayed amount"),
        formula("Wallet statement balance = available funds - total debit"),
        h2("5.3 Nsano collection wallet vs ledger"),
        p("This reuses the Nsano wallet calculation structure but labels the wallet as Nsano Collections. Credit collections and counts can come directly from successful Nsano W2A rows. Optional charge rows supply charge total and count. Mambu collection data supplies the wallet debit/transaction side selected by the workflow."),
        h2("5.4 ITC wallet vs ledger - credit classifications"),
        p("Credit transfer narration is classified into settlement, reversal, prepaid reversal, or transfer to wallet. The exact classification helper reads the narration suffix/token used in the ITC exports."),
        h2("5.5 ITC wallet vs ledger - debit classifications"),
        bullet("The debit identifier is compared with the ITC statement and credit-transfer indexes."),
        bullet("A matched statement narration determines operational categories such as disbursement, referral award, upsales refund, or savings."),
        bullet("If no statement row is found and the debit text contains “withdrawal”, it is transfers_from_bank."),
        bullet("If there is no statement match but the same key exists in credit transfers, it is credit transfer (prepaid_reversal)."),
        bullet("Otherwise it is not_found."),
        formula("Available funds = ledger balance + settlement + reversal + transfer to wallet + positive unidentified amount"),
        formula("Total debit = transfer to bank + referral awards + savings + upsales refund + disbursement + absolute negative unidentified amount"),
        formula("Wallet statement balance = available funds - total debit"),
        h2("5.6 Vodafone wallet vs ledger"),
        p("<b>Prompt:</b> Enter Balance and Delayed Transactions, then upload one or more Vodafone wallet files."),
        bullet("Balance is forced to a positive value. Paid In, Withdrawn, and Balance values in the cleaned output are also shown as positive values."),
        bullet("Total Collections is the sum of absolute Paid In values."),
        bullet("Total Charges includes rows whose Details contains “Pay Bill Charge”; zero withdrawals are excluded from the count."),
        bullet("Transfer to Bank includes rows whose Details contains “Funds Transfer from Utility Account to Working Account”."),
        formula("Available funds = absolute balance + total collections + positive delayed amount"),
        formula("Total debit = total charges + transfer to bank + absolute negative delayed amount"),
        formula("Wallet statement balance = absolute(available funds - total debit)"),
        h2("5.7 Zenith wallet vs ledger"),
        p("The project also provides a Zenith Wallet vs Ledger filtering/ledger automation. It accepts ledger balance and delayed transactions, processes uploaded Zenith wallet files, and produces a ledger workbook with summary metrics and cleaned supporting data. It is separate from the Zenith collection reconciliation, which matches Mambu to bank rows by date and amount."),
        h2("5.8 Master ledger workbook"),
        bullet("The master ledger snapshot includes Nsano Collection Ledger, Nsano Disbursement Ledger, and ITC Wallet Ledger results when present."),
        bullet("The Overview sheet displays ledger balance, available funds, total debit, and calculated wallet statement balance per wallet, plus a combined total."),
        bullet("The Detail sheet expands every manual input and transaction category with count, amount, and explanatory notes."),
        PageBreak(),
    ]

    s += [h1("6. Summaries and dashboards")]
    s += [
        h2("6.1 Collection/disbursement master summary"),
        p("The master summary uses result objects already produced by the workflows. It deliberately removes raw output bytes from its snapshot, keeping only the values required to build the combined workbook."),
        table(
            ["Section", "What it reports"],
            [
                ["Dashboard", "High-level collection/disbursement totals, exception values, variances, completion status, and visual status panels."],
                ["Source summaries", "External wallet/source counts and amounts compared with Mambu counts and amounts."],
                ["Exceptions", "Write-Off, Unidentified, Not Found, Refund, Charges, Settlement, and other adjustments."],
                ["Daily summaries", "Separate daily views for Nsano, ITC, and Vodafone, with amount and charge totals."],
                ["Not-found details", "Grouped row-level evidence by workflow and comparison direction."],
            ],
            [42 * mm, 128 * mm],
        ),
        h2("6.2 Monthly Summary"),
        p("<b>Prompt:</b> Run collection and/or disbursement reconciliations first, optionally choose whether collection workflows are included, then select Generate Monthly Summary."),
        bullet("No transaction files are uploaded again. The generator reads current-session result objects."),
        bullet("A disbursement-only summary is allowed when collection workflows are excluded."),
        bullet("Run Status shows which expected workflows are ready or missing, with rows, amount, and notes."),
        bullet("Collection Summary combines Nsano, ITC/Vodafone, Zenith, and Write-Off results."),
        bullet("Disbursement Summary combines Nsano, ITC, MTN Manual, and Vodafone Manual results."),
        bullet("Wallet Summary combines Nsano Collection, Nsano Disbursement, and ITC ledger metrics."),
        bullet("Granular summary sheets expose category-level figures and status totals used by the dashboard."),
        bullet("If any reconciliation result changes after generation, the app warns that the monthly summary is stale and should be regenerated."),
        h2("6.3 Session-state dependency"),
        p("Results exist in Streamlit session state during the active browser session. A summary can only include workflows that have been completed or restored into that state. Download bytes are not treated as business metrics."),
        PageBreak(),
    ]

    s += [h1("7. Google Drive archive")]
    s += [
        p("The archive is Google Drive-only. The application does not maintain a local archive of generated finance workbooks."),
        h2("7.1 Save logic"),
        bullet("Each generated workbook is assigned a unique ID, a cleaned storage name, creation time, size, SHA-256 content hash, month, category, workflow, and archive scope."),
        bullet("Category folders are Collection Recon, Disbursement Recon, Filtering, and Summary. Month folders are also used."),
        bullet("If a current record has the same archive key and the same content hash, the existing Drive record is returned instead of uploading a duplicate."),
        bullet("A new workbook with the same month/category/workflow/file name becomes current. Older versions become previous and point to the new record."),
        bullet("If upload succeeds but updating the Drive index fails, the code attempts to delete the newly uploaded orphan file."),
        h2("7.2 Archive screen prompts and controls"),
        bullet("Filters: month, type/category, workflow, show previous versions, and free-text search."),
        bullet("Actions: refresh from Drive, download selected, prepare visible-results ZIP, prepare full-month ZIP, delete selected, and delete a confirmed month."),
        bullet("Refresh prunes index records whose Drive files were deleted or trashed outside the app."),
        bullet("Current and previous versions are labelled. Previous versions are hidden by default."),
        h2("7.3 Configuration"),
        bullet("A Google Drive folder ID and service-account credentials can come from Streamlit secrets or environment variables."),
        bullet("The recommended target is a Shared Drive folder. A service account normally has no personal My Drive storage quota."),
        bullet("Optional domain-wide delegation is supported when configured by a Google Workspace administrator."),
        bullet("If Drive is unavailable, the active result can still be downloaded during the session, but it is not marked as archived."),
        h2("7.4 Security boundary"),
        p("Credential values are configuration, not workbook content. They should remain in Streamlit secrets, environment variables, or a protected credential file and should never be embedded in generated reports.", "Callout"),
        PageBreak(),
    ]

    s += [h1("8. Validation, diagnostics, and error behaviour")]
    s += [
        bullet("Run buttons are disabled until required files and manual numeric inputs are present."),
        bullet("The app inspects detected columns and workbook sheet names, then shows expected-versus-found information."),
        bullet("Vodafone cleanup and Vodafone Manual files have special header detection because their exports may contain metadata before the real header."),
        bullet("Wrong or missing sheets can be reported before a long reconciliation starts."),
        bullet("Zero key overlap warnings help detect a wrong month, wrong file, or wrong Mambu sheet."),
        bullet("Archive failures are shown separately from reconciliation failures. A valid workbook can still be downloaded even when Drive save fails."),
        bullet("Timing rows record how long important processing stages take and can be shown to the user."),
        h2("8.1 Important edge cases"),
        table(
            ["Case", "Current behaviour"],
            [
                ["Blank key", "Usually Not Found with a blank-key reason."],
                ["Same key, different amount", "Nsano collection reference matching still treats it as matched but records the amount mismatch. Several key-only workflows also remain matched."],
                ["Duplicate keys", "Indexes store all candidates, but many comparisons select the first candidate or the first exact-amount candidate. Rows are not consumed."],
                ["Invalid number", "Filtering totals generally treat it as zero; stricter manual inputs show an input error."],
                ["Nsano failed collection", "Ignored in collection reconciliation and counted for diagnostics."],
                ["Nsano failed disbursement", "The current disbursement loader does not apply the same failed-result exclusion."],
                ["Positive/negative delayed value", "Positive is normally credit/available funds; negative becomes an absolute debit."],
                ["Stale monthly summary", "The app warns after an underlying result changes."],
            ],
            [52 * mm, 118 * mm],
        ),
        PageBreak(),
    ]

    s += [h1("9. Workbook generation and technical design")]
    s += [
        h2("9.1 Application design"),
        bullet("app.py is the Streamlit interface and workflow coordinator. It controls tabs, prompts, uploads, state, diagnostics, runs, downloads, and archive calls."),
        bullet("Business logic is split into filtering, collection reconciliation, Zenith reconciliation, disbursement/ledger reconciliation, summary, master ledger, and archive modules."),
        bullet("The XLSX writers build Office Open XML files directly with zipfile and XML. This avoids depending on a spreadsheet library for core production output."),
        bullet("Temporary “.tmp” output files are replaced atomically after workbook writing finishes."),
        bullet("Source row caches use path, modification time, file size, and sheet name so repeated reads can be reused safely during a run."),
        h2("9.2 Workbook presentation"),
        bullet("Workbooks use shared style constants for titles, headers, currency, percentages, matched/not-found statuses, wallet statements, and dashboard panels."),
        bullet("Column widths are estimated from headers and a sample of up to 1,000 rows."),
        bullet("Large worksheets are written in row batches and ZIP compression is deliberately set for faster generation."),
        bullet("Sheet names are sanitized to Excel limits; data sheets normally keep original source columns plus comparison evidence."),
        h2("9.3 Command-line support"),
        p("The principal reconciliation builders also expose command-line entry points. These accept source paths and write the same XLSX outputs without the Streamlit interface. The Streamlit app remains the normal operator-facing entry point."),
        h2("9.4 Dependencies"),
        table(
            ["Dependency", "Use"],
            [
                ["streamlit >= 1.58.0", "Interactive browser application."],
                ["google-api-python-client", "Google Drive files, folders, upload, download, delete, and index operations."],
                ["google-auth", "Service-account authentication and optional delegated credentials."],
                ["google-auth-httplib2", "Authenticated HTTP transport for Google APIs."],
                ["Python standard library", "CSV, XML, ZIP/XLSX, dates, decimals, hashing, temporary files, and tests."],
            ],
            [55 * mm, 115 * mm],
        ),
        PageBreak(),
    ]

    s += [h1("10. Tests and verification coverage")]
    s += [
        p("The included unit tests protect several high-risk rules:"),
        bullet("Collection master not-found totals remain separated by Nsano, ITC/Vodafone, and Zenith."),
        bullet("Collection not-found detail is grouped into a single collection sheet while preserving workflow and direction groups."),
        bullet("Not-found evidence retains the raw ID and normalized reconciliation key."),
        bullet("Daily wallet sheets are separated, totalled, and represented in a daily overview."),
        bullet("Vodafone Pay Bill Charge totals use absolute non-zero Withdrawn amounts."),
        bullet("Vodafone charge metrics appear in the ITC/Vodafone reconciliation summary."),
        bullet("Nsano “Transfer to Bank” and “Transfer Bank” purpose variations map to transfer_to_bank."),
        bullet("Write-Off allocation obeys channel routing, exposes ambiguous cross-wallet identifiers, uses Vodafone Paid In rows, and reports Zenith date/amount mismatches explicitly."),
        h2("10.1 Current coverage limitation"),
        p("The tests are focused and useful, but they do not cover every loader, every status path, every workbook layout, Drive API behaviour, or the full Streamlit interface. Finance approval should still include controlled sample files and reconciliation against known totals.", "Callout"),
        PageBreak(),
    ]

    s += [h1("11. Project file map")]
    s += [
        table(
            ["File", "Responsibility"],
            [
                ["app.py", "Streamlit tabs, prompts, uploads, state, diagnostics, workflow orchestration, downloads, and archive UI."],
                ["collection_filtering.py", "Mambu/ITC/Nsano filters, Vodafone cleanup, Vodafone ledger, and filtering workbook output."],
                ["build_reconciliation_template.py", "Core records, normalization, Nsano/ITC/Vodafone/Write-Off collection matching, summaries, and common XLSX writer/styles."],
                ["zenith_collection_reconciliation.py", "Zenith date-and-amount matching, Zenith summaries, and Zenith ledger helpers."],
                ["disbursement_reconciliation.py", "Nsano, ITC, MTN Manual, Vodafone Manual, and wallet-ledger logic."],
                ["collections_master_summary.py", "Master dashboard, source/exception tables, daily sheets, and grouped not-found sheets."],
                ["monthly_summary.py", "Session snapshot, run status, dashboard, granular, collection, disbursement, and wallet summary sheets."],
                ["master_ledger.py", "Combined ledger overview and detailed line-item workbook."],
                ["recon_archive.py", "Google Drive configuration, upload, index, versions, download, prune, and delete logic."],
                ["build_automation_logic_doc.py", "Legacy Word logic reference generator."],
                ["render_preview_fallback.py", "Fallback workbook preview rendering helper."],
                ["inspect_config_palette.py", "Small helper for examining configuration colour values."],
                ["test_*.py", "Focused regression tests for summaries, raw IDs, wallet transfer text, and Vodafone charges."],
                ["outputs/", "Preview workbooks and verification material; not the live application logic."],
                ["GOOGLE_DRIVE_SETUP.md", "Operator setup notes for the Drive-only archive."],
                ["requirements.txt", "Runtime Python package requirements."],
            ],
            [52 * mm, 118 * mm],
        ),
        PageBreak(),
    ]

    s += [h1("12. Operator checklist")]
    s += [
        h2("Before running"),
        bullet("Confirm all source files belong to the same reporting month."),
        bullet("Use the Filtering tab first when a raw source requires splitting or Vodafone cleanup."),
        bullet("Confirm expected sheets and columns in the diagnostics."),
        bullet("Enter ledger balances and manual adjustments from approved month-end records."),
        h2("After running"),
        bullet("Review total source rows and amounts before relying on match rates."),
        bullet("Investigate blank keys, not-found rows, amount mismatches, and zero-overlap warnings."),
        bullet("Review raw and normalized IDs together in exception sheets."),
        bullet("Confirm that the generated output shows as saved in Drive, or download it immediately if Drive is unavailable."),
        bullet("Regenerate the monthly/master summary after rerunning any underlying reconciliation."),
        h2("Approval focus"),
        bullet("Confirm the reporting month, source completeness, manual ledger values, classification totals, and all exception categories."),
        bullet("Pay special attention to duplicate identifiers, key-only matches with amount differences, and same-date/same-amount Zenith transactions."),
        p("This reference describes the code as reviewed. If business policy changes, update the rules and this documentation together so the application and operator guidance remain aligned.", "Callout"),
    ]
    return s


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = LogicDoc(str(OUT))
    doc.build(story())
    print(OUT)


if __name__ == "__main__":
    main()
