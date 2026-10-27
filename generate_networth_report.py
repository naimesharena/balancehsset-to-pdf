#!/usr/bin/env python3
"""
Build a Net Worth Report (Excel) from the balance sheet PDF in this repository.

Pipeline
--------
1.  Parse the PDF with pdfplumber, preserving the two-column (Liabilities / Assets)
    layout, the bold group headings and every account line (including wrapped names).
2.  Reconcile: sum of account lines == printed group total for every group, and
    sum of group totals == printed grand total on both sides.
3.  Classify each balance into an asset / liability / equity category and write a
    formatted, formula-driven .xlsx workbook (6 sheets + charts).

Run:  python3 generate_networth_report.py
Deps: pdfplumber, openpyxl   (see requirements.txt)
"""

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from collections import defaultdict
from decimal import Decimal
from itertools import zip_longest
from pathlib import Path

import pdfplumber
import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
DEFAULT_PDF = HERE / "Vipulbhai BALANCESHEET_1-23102026161349.PDF"
DEFAULT_OUT = HERE / "Vipulbhai_Patel_Net_Worth_Report_10-09-2026.xlsx"

ENTITY = "VIPULBHAI LALJIBHAI PATEL"
AS_ON = "10/09/2026"

# ---------------------------------------------------------------------------
# 1. PDF parsing
# ---------------------------------------------------------------------------

QUIRKS: list[str] = []


def parse_amount(text: str) -> Decimal:
    """'1,70,75,736.77' -> Decimal; the source prints '.37' for 0.37."""
    raw = text.strip().replace(",", "").replace(" ", "")
    if raw.startswith("."):
        QUIRKS.append(
            f"The source prints '{raw}' with no leading zero (credit card HDFC-3988); "
            f"it is interpreted as 0{raw}."
        )
        raw = "0" + raw
    if raw.startswith("-."):
        raw = "-0" + raw[1:]
    if not re.fullmatch(r"-?\d+(\.\d+)?", raw):
        raise ValueError(f"Unrecognised amount: {text!r}")
    return Decimal(raw)


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def join_name(existing: str, continuation: str) -> str:
    """Rejoin a wrapped account name, respecting hyphens used as line breaks."""
    if existing.endswith(" -"):
        return existing + " " + continuation
    if existing.endswith(("-", "(")):
        return existing + continuation
    return existing + " " + continuation


def parse_pdf(path: Path) -> tuple[list[dict], dict[str, Decimal]]:
    """Return (groups, grand_totals).

    A group = one bold heading plus its indented account lines, e.g.
    'INVESTMENT A/C  67,42,939.15'.  The same group may be repeated on the
    next page (the asset-side LOAN & ADVANCES A/C is): those are merged.
    """
    groups: list[dict] = []
    grand: dict[str, Decimal] = {}

    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            for side, x_lo, x_hi, name_cut in (
                ("Liabilities", 0, 292, 215),
                ("Assets", 292, 612, 490),
            ):
                chars = [
                    c
                    for c in page.chars
                    if x_lo <= c["x0"] < x_hi and c["top"] > 85
                ]
                entries = []
                for c in chars:
                    kind = "N" if c["x0"] < name_cut else "A"
                    entries.append(
                        (c["top"], c["x0"], kind, c["text"], "Bold" in c["fontname"])
                    )
                entries.sort(key=lambda e: (e[0], e[1]))

                # Cluster characters into visual lines (name and amount are drawn
                # at slightly different y positions by the report writer).
                clusters: list[list] = []
                for e in entries:
                    if clusters and abs(e[0] - clusters[-1][0]) <= 2.6:
                        clusters[-1][1].append(e)
                    else:
                        clusters.append([e[0], [e]])

                current: dict | None = None
                last_item: dict | None = None
                for top, es in clusters:
                    name = "".join(
                        e[3] for e in sorted(es, key=lambda x: x[1]) if e[2] == "N"
                    ).strip()
                    amount = "".join(
                        e[3] for e in sorted(es, key=lambda x: x[1]) if e[2] == "A"
                    ).strip()
                    is_bold = any(e[4] for e in es)
                    if not name and not amount:
                        continue

                    if "T O T A L" in name or name.replace(" ", "") == "TOTAL:":
                        grand[side] = parse_amount(amount)
                        continue

                    if is_bold:  # group heading carrying the printed group total
                        gname = norm(name)
                        current = next(
                            (
                                g
                                for g in groups
                                if g["side"] == side and g["name"] == gname
                            ),
                            None,
                        )
                        if current is None:
                            current = {
                                "side": side,
                                "name": gname,
                                "total": parse_amount(amount),
                                "items": [],
                                "pages": [],
                            }
                            groups.append(current)
                        if page_no not in current["pages"]:
                            current["pages"].append(page_no)
                        last_item = None
                        continue

                    if current is None:
                        raise ValueError(f"Orphan line on page {page_no}: {name!r}")
                    if amount:  # a new account line
                        last_item = {
                            "name": norm(name),
                            "amount": parse_amount(amount),
                            "page": page_no,
                        }
                        current["items"].append(last_item)
                    elif name and last_item is not None:  # wrapped continuation
                        last_item["name"] = join_name(last_item["name"], norm(name))

    return groups, grand


def validate(groups: list[dict], grand: dict[str, Decimal]) -> list[dict]:
    """Check every group and both sides; raise if anything does not reconcile."""
    checks = []
    for g in groups:
        s = sum((i["amount"] for i in g["items"]), Decimal("0"))
        checks.append(
            {
                "side": g["side"],
                "group": g["name"],
                "pdf_total": g["total"],
                "items_sum": s,
                "diff": s - g["total"],
                "pages": g["pages"],
                "n": len(g["items"]),
            }
        )
        if s != g["total"]:
            raise AssertionError(
                f"{g['side']} / {g['name']}: items {s} != printed {g['total']}"
            )

    for side, total in grand.items():
        side_sum = sum(
            (g["total"] for g in groups if g["side"] == side), Decimal("0")
        )
        print(f"  {side:<12} groups={sum(1 for g in groups if g['side']==side):>2}  "
              f"sum={side_sum:>16,.2f}   printed total={total:>16,.2f}")
        if side_sum != total:
            raise AssertionError(f"{side}: group sum {side_sum} != {total}")
    if grand.get("Assets") != grand.get("Liabilities"):
        raise AssertionError("The two sides of the balance sheet do not agree.")
    print("  All group sub-totals and both grand totals reconcile exactly.")
    return checks


# ---------------------------------------------------------------------------
# 2. Classification
# ---------------------------------------------------------------------------

ASSET_CATEGORIES = [
    "Land & Property",
    "Gold & Valuables",
    "Vehicles",
    "Investments (MF, Shares, Trading)",
    "Deposits & Fixed Deposits",
    "Loans & Advances Given",
    "Receivables & Tax Recoverable",
    "Cash & Bank Balances",
]

LIABILITY_CATEGORIES = [
    "Secured Loans (Housing)",
    "Unsecured Loans & Borrowings",
    "Trade Creditors & Payables",
    "Current Liabilities",
    "Credit Card Dues",
    "Other Credit / Contra Balances",
]

EQUITY_CATEGORIES = [
    "Owner's Capital",
    "Reserves & Surplus",
    "Unappropriated Income",
]

LIAB_GROUP_CATEGORY = {
    "SECURED LOAN A/C": "Secured Loans (Housing)",
    "LOAN & ADVANCES A/C": "Unsecured Loans & Borrowings",
    "UN-SECURED LOAN A/C": "Unsecured Loans & Borrowings",
    "SUNDRY CREDITORS/DEBTORS A/C": "Trade Creditors & Payables",
    "CURRENT LIABILITIES A/C": "Current Liabilities",
    "EXPENSE(CAPITAL)A/C": "Credit Card Dues",
    "ASSETS A/C": "Other Credit / Contra Balances",
    "INVESTMENT A/C": "Other Credit / Contra Balances",
    "SUNDRY DEBTORS/CREDITORS A/C": "Other Credit / Contra Balances",
    "CASH BOOK A/C": "Other Credit / Contra Balances",
}

EQUITY_GROUP_CATEGORY = {
    "CAPITAL A/C": "Owner's Capital",
    "CAPITAL RESERVE AND SURPLUS A/C": "Reserves & Surplus",
    "INCOME A/C": "Unappropriated Income",
}


def asset_category(group: str, item_name: str) -> str:
    n = item_name.upper()
    if group == "ASSETS A/C":
        if "GOLD ORNAMENTS" in n:
            return "Gold & Valuables"
        if "AKSHAR AUTO ELECTRIC" in n:
            return "Vehicles"
        if n.startswith(("GSPC GAS", "SHOP RENT DEPOSIT")):
            return "Deposits & Fixed Deposits"
        return "Land & Property"
    if group == "INVESTMENT A/C":
        if n.startswith("ICICI BANK F.D"):
            return "Deposits & Fixed Deposits"
        return "Investments (MF, Shares, Trading)"
    if group in ("LOAN & ADVANCES A/C", "UN-SECURED LOAN A/C"):
        return "Loans & Advances Given"
    if group in ("SUNDRY DEBTORS/CREDITORS A/C", "PROVISION A/C"):
        return "Receivables & Tax Recoverable"
    if group == "BANK BOOK A/C":
        return "Cash & Bank Balances"
    if group == "DEPOSIT A/C":
        return "Deposits & Fixed Deposits"
    raise ValueError(f"Unmapped asset group: {group}")


def build_records(groups: list[dict]) -> tuple[list[dict], list[dict], dict]:
    assets, left = [], []
    group_order = {g["name"]: n for n, g in enumerate(groups)}
    for g in groups:
        for pos, i in enumerate(g["items"]):
            rec = {
                "group": g["name"],
                "name": i["name"],
                "amount": i["amount"],
                "page": i["page"],
                "pages": g["pages"],
                "order": (group_order[g["name"]], pos),
            }
            if g["side"] == "Assets":
                rec["category"] = asset_category(g["name"], i["name"])
                assets.append(rec)
            else:
                if g["name"] in EQUITY_GROUP_CATEGORY:
                    rec["classification"] = "Equity"
                    rec["category"] = EQUITY_GROUP_CATEGORY[g["name"]]
                else:
                    rec["classification"] = "External Liability"
                    rec["category"] = LIAB_GROUP_CATEGORY[g["name"]]
                left.append(rec)

    assets.sort(
        key=lambda r: (ASSET_CATEGORIES.index(r["category"]), r["order"])
    )
    left.sort(
        key=lambda r: (
            0 if r["classification"] == "External Liability" else 1,
            (
                LIABILITY_CATEGORIES.index(r["category"])
                if r["classification"] == "External Liability"
                else EQUITY_CATEGORIES.index(r["category"])
            ),
            r["order"],
        )
    )

    cat_tot = defaultdict(lambda: Decimal("0"))
    for r in assets + left:
        cat_tot[r["category"]] += r["amount"]

    totals = {
        "assets": sum((r["amount"] for r in assets), Decimal("0")),
        "external_liabilities": sum(
            (r["amount"] for r in left if r["classification"] == "External Liability"),
            Decimal("0"),
        ),
        "equity": sum(
            (r["amount"] for r in left if r["classification"] == "Equity"),
            Decimal("0"),
        ),
    }
    totals["net_worth"] = totals["assets"] - totals["external_liabilities"]
    if totals["net_worth"] != totals["equity"]:
        raise AssertionError("Assets - Liabilities does not equal the equity accounts.")
    return assets, left, {"cat_tot": dict(cat_tot), **totals}


# ---------------------------------------------------------------------------
# 3. Excel helpers / styling
# ---------------------------------------------------------------------------

NAVY = "1F3864"
TEAL = "0F6B78"
GOLD = "B8860B"
LIGHT_BLUE = "EAF1F8"
LIGHT_GREY = "F4F6F8"
WHITE = "FFFFFF"
INK = "1F2933"
GREEN = "E2EFDA"
AMBER = "FFF2CC"

INR = '"₹ "#,##,##0.00'          # Indian digit grouping, e.g. 7,39,01,963.50
PCT = "0.0%"
NUM_X = '0.00"x"'


def inr(value) -> str:
    """Format a Decimal/float with Indian digit grouping for prose text."""
    s = f"{Decimal(value):.2f}"  # no separators yet — add Indian grouping below
    sign = "-" if s.startswith("-") else ""
    whole, frac = s.lstrip("-").split(".")
    if len(whole) > 3:
        tail = whole[-3:]
        head = whole[:-3]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [tail])
    return f"{sign}{whole}.{frac}"

THIN = Side(style="thin", color="D5DCE5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class Book:
    """Thin wrapper that also caches the numeric result of every formula.

    openpyxl cannot store cached formula results, so patch_cached_values()
    writes them into the sheet XML after saving.  That way the workbook shows
    correct figures in previewers that do not recalculate, while staying fully
    formula-driven in Excel / Sheets / LibreOffice.
    """

    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        self.cache: dict[tuple[str, str], Decimal | float] = {}

    def sheet(self, title: str, widths: dict[str, float], tab: str = TEAL):
        ws = self.wb.create_sheet(title)
        ws.sheet_properties.tabColor = tab
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        ws.sheet_view.showGridLines = False
        return ws

    def formula(self, ws, coord: str, expression: str, value, fmt: str | None = None):
        c = ws[coord]
        c.value = expression if expression.startswith("=") else "=" + expression
        self.cache[(ws.title, coord)] = value
        if fmt:
            c.number_format = fmt
        return c

    def value(self, ws, coord: str, value, fmt: str | None = None):
        c = ws[coord]
        c.value = float(value) if isinstance(value, Decimal) else value
        if fmt:
            c.number_format = fmt
        return c

    def patch_cached_values(self, path: Path):
        with zipfile.ZipFile(path) as z:
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            rid2target = {}
            for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
                rid = re.search(r'Id="([^"]+)"', tag)
                target = re.search(r'Target="([^"]+)"', tag)
                if rid and target:
                    rid2target[rid.group(1)] = target.group(1)
            title2path = {}
            for name, rid in re.findall(
                r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml
            ):
                name = (
                    name.replace("&amp;", "&")
                    .replace("&lt;", "<")
                    .replace("&gt;", ">")
                    .replace("&quot;", '"')
                    .replace("&apos;", "'")
                )
                target = rid2target[rid]
                if target.startswith("/"):
                    target = target.lstrip("/")
                elif not target.startswith("xl/"):
                    target = "xl/" + target
                title2path[name] = target

            patches: dict[str, bytes] = {}
            for title, sheet_path in title2path.items():
                sheet_cache = {
                    coord: v
                    for (t, coord), v in self.cache.items()
                    if t == title
                }
                if not sheet_cache:
                    continue
                xml = z.read(sheet_path).decode("utf-8")
                parts, pos = [], 0
                # Iterate over every <c ...> opening tag.  Empty cells are
                # self-closing (<c ... />) and must be skipped, otherwise a
                # non-greedy regex would swallow the following formula cell.
                for m in re.finditer(r"<c\b[^>]*>", xml):
                    tag = m.group(0)
                    parts.append(xml[pos:m.start()])
                    if tag.endswith("/>"):
                        parts.append(tag)
                        pos = m.end()
                        continue
                    end = xml.find("</c>", m.end())
                    if end < 0:
                        parts.append(tag)
                        pos = m.end()
                        continue
                    inner = xml[m.end():end]
                    ref_m = re.search(r'r="([A-Z]+\d+)"', tag)
                    if ref_m and ref_m.group(1) in sheet_cache:
                        v = sheet_cache[ref_m.group(1)]
                        text = format(float(v), ".10f").rstrip("0").rstrip(".")
                        if text in ("", "-0"):
                            text = "0"
                        if re.search(r"<v\s*/>", inner):
                            inner = re.sub(r"<v\s*/>", f"<v>{text}</v>", inner)
                        elif re.search(r"<v>.*?</v>", inner, re.S):
                            inner = re.sub(
                                r"<v>.*?</v>", f"<v>{text}</v>", inner, flags=re.S
                            )
                        else:
                            inner += f"<v>{text}</v>"
                        tag = re.sub(r'\s+t="[^"]*"', "", tag)
                    parts.append(tag + inner + "</c>")
                    pos = end + len("</c>")
                parts.append(xml[pos:])
                patches[sheet_path] = "".join(parts).encode("utf-8")

        tmp = path.with_suffix(".tmp.xlsx")
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                zout.writestr(item, patches.get(item.filename, zin.read(item.filename)))
        shutil.move(str(tmp), str(path))


def cell_style(
    c,
    *,
    bold=False,
    size=10,
    color=INK,
    bg=None,
    align="left",
    fmt=None,
    wrap=False,
    border=True,
    italic=False,
):
    c.font = Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)
    c.alignment = Alignment(
        horizontal=align, vertical="center", wrap_text=wrap, indent=0
    )
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    return c


def band(ws, row: int, text: str, last_col: int, *, bg=NAVY, size=11, height=22):
    ws.merge_cells(
        start_row=row, start_column=1, end_row=row, end_column=last_col
    )
    c = ws.cell(row=row, column=1, value=text)
    cell_style(c, bold=True, size=size, color=WHITE, bg=bg, border=False)
    for col in range(2, last_col + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=bg)
    ws.row_dimensions[row].height = height
    return c


def table_header(ws, row: int, labels: list[str], *, bg=TEAL):
    for col, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=col, value=label)
        cell_style(c, bold=True, color=WHITE, bg=bg, wrap=True, align="left")
    ws.row_dimensions[row].height = 28


def note_lines(ws, row: int, lines: list[str], last_col: int = 8):
    chars_per_line = {3: 135, 4: 110, 8: 118}.get(last_col, 115)
    for i, text in enumerate(lines):
        r = row + i
        ws.merge_cells(
            start_row=r, start_column=1, end_row=r, end_column=last_col
        )
        c = ws.cell(row=r, column=1, value=text)
        cell_style(c, size=10, wrap=True, border=False)
        n_lines = max(1, -(-len(text) // chars_per_line))
        ws.row_dimensions[r].height = 15 * n_lines
    return row + len(lines)


def section_heading(ws, row: int, text: str, last_col: int = 8):
    c = ws.cell(row=row, column=1, value=text)
    cell_style(c, bold=True, size=11, color=NAVY, bg=LIGHT_BLUE, border=False)
    for col in range(2, last_col + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws.row_dimensions[row].height = 20


def total_row(ws, row: int, label: str, last_col: int, cells: dict[int, object], *, bg=GREEN):
    c = ws.cell(row=row, column=1, value=label)
    cell_style(c, bold=True, bg=bg)
    for col in range(2, last_col + 1):
        x = ws.cell(row=row, column=col)
        if col in cells:
            x.value = cells[col]
        cell_style(x, bold=True, bg=bg, fmt=cells.get(f"fmt{col}"))
    ws.row_dimensions[row].height = 19


def internal_link(ws, coord: str, label: str, location: str):
    c = ws[coord]
    c.value = label
    c.hyperlink = Hyperlink(ref=coord, location=location, tooltip=label)
    cell_style(c, size=10, color=TEAL, bold=True, bg=LIGHT_GREY, border=False)


def setup_print(ws, landscape=True, header_rows: str | None = None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    if header_rows:
        ws.print_title_rows = header_rows


# ---------------------------------------------------------------------------
# 4. Workbook sheets
# ---------------------------------------------------------------------------


def build_workbook(
    groups: list[dict],
    grand: dict[str, Decimal],
    checks: list[dict],
    assets: list[dict],
    left: list[dict],
    totals: dict,
    pdf_name: str,
) -> Book:
    bk = Book()
    cat = totals["cat_tot"]
    A = totals["assets"]
    L = totals["external_liabilities"]
    E = totals["equity"]
    NW = totals["net_worth"]
    GT = grand["Assets"]

    # ---------------- Sheet 1: Summary ----------------
    ws = bk.sheet(
        "Net Worth Summary",
        {"A": 46, "B": 22, "C": 16, "D": 30, "E": 12, **{c: 10 for c in "FGHIJKLM"}},
        tab=NAVY,
    )
    band(ws, 1, "NET WORTH REPORT", 4, size=18, height=34)
    band(ws, 2, ENTITY, 4, bg=TEAL, size=12, height=24)
    for coord, val, fmt, align in [
        ("A3", "As on", None, "left"),
        ("B3", AS_ON, None, "left"),
        ("C3", "Currency", None, "left"),
        ("D3", "INR (₹) — Indian Rupees", None, "left"),
    ]:
        c = ws[coord]
        c.value = val
        cell_style(c, bold=(coord in ("A3", "C3")), border=False, align=align)
    ws.merge_cells("B4:D4")
    ws["A4"] = "Prepared from"
    ws["B4"] = f"{pdf_name} (2 pages, {sum(len(g['items']) for g in groups)} account lines)"
    cell_style(ws["A4"], bold=True, border=False)
    cell_style(ws["B4"], border=False)
    for col in "CD":
        ws[f"{col}4"].fill = PatternFill("solid", fgColor=WHITE)
    ws.merge_cells("A5:D5")
    ws["A5"] = (
        f"Balance Sheet total on both sides: ₹ {inr(GT)}  |  "
        "All amounts are book values as per the source, not market valuations."
    )
    cell_style(ws["A5"], size=9, italic=True, color="5A6B7B", border=False)

    for col in "BCDEF":
        ws[f"{col}5"].fill = PatternFill("solid", fgColor=WHITE)

    ws["A6"] = "Jump to:"
    cell_style(ws["A6"], bold=True, size=9, border=False)
    for coord, label, loc in [
        ("B6", "Calculation", "'Net Worth Calculation'!A1"),
        ("C6", "Assets Detail", "'Assets Detail'!A1"),
        ("D6", "Liabilities Detail", "'Liabilities Detail'!A1"),
        ("E6", "Source B/S", "'Balance Sheet (Source)'!A1"),
        ("F6", "Notes & Method", "'Notes & Method'!A1"),
    ]:
        internal_link(ws, coord, label, loc)
    ws.row_dimensions[6].height = 18

    band(ws, 8, "HEADLINE FIGURES", 4, bg=NAVY, size=11)
    table_header(ws, 9, ["Particulars", "Amount (₹)", "% of Assets", "Basis"])

    # Asset / liability / net-worth headline.  Links to the detail sheets are
    # filled in after those sheets are created (see bottom of this function).
    headline = [
        (10, "Total Assets", None, "Asset side of the Balance Sheet"),
        (11, "Total External Liabilities", None, "Liabilities side, excluding equity"),
        (12, "NET WORTH  (Assets − Liabilities)", None, "Owner's equity"),
        (13, "Check: Assets − Liabilities − Net Worth", None, "Must be 0.00"),
    ]
    for r, label, _, basis in headline:
        cell_style(ws.cell(row=r, column=1, value=label), bold=(r == 12))
        cell_style(ws.cell(row=r, column=4, value=basis), size=9, color="5A6B7B")
        ws.row_dimensions[r].height = 20 if r != 12 else 24
        if r == 12:
            for col in range(1, 5):
                ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(row=r, column=1).font = Font(
                name="Calibri", size=11, bold=True, color=NAVY
            )

    band(ws, 15, "ASSET MIX BY CATEGORY", 4, bg=TEAL)
    table_header(ws, 16, ["Asset Category", "Amount (₹)", "% of Assets", "No. of Accounts"])
    asset_first, asset_last = 17, 17 + len(ASSET_CATEGORIES) - 1
    asset_total_row = asset_last + 1
    for idx, name in enumerate(ASSET_CATEGORIES):
        r = asset_first + idx
        cell_style(ws.cell(row=r, column=1, value=name))
        cell_style(ws.cell(row=r, column=2), fmt=INR)
        cell_style(ws.cell(row=r, column=3), fmt=PCT)
        cell_style(ws.cell(row=r, column=4), fmt="0")
        ws.row_dimensions[r].height = 18
    cell_style(ws.cell(row=asset_total_row, column=1, value="TOTAL ASSETS"), bold=True, bg=GREEN)
    cell_style(ws.cell(row=asset_total_row, column=2), bold=True, bg=GREEN, fmt=INR)
    cell_style(ws.cell(row=asset_total_row, column=3), bold=True, bg=GREEN, fmt=PCT)
    cell_style(ws.cell(row=asset_total_row, column=4), bold=True, bg=GREEN, fmt="0")

    band(ws, 27, "LIABILITY MIX BY CATEGORY", 4, bg=TEAL)
    table_header(ws, 28, ["Liability Category", "Amount (₹)", "% of Assets", "No. of Accounts"])
    liab_first, liab_last = 29, 29 + len(LIABILITY_CATEGORIES) - 1
    liab_total_row = liab_last + 1
    for idx, name in enumerate(LIABILITY_CATEGORIES):
        r = liab_first + idx
        cell_style(ws.cell(row=r, column=1, value=name))
        cell_style(ws.cell(row=r, column=2), fmt=INR)
        cell_style(ws.cell(row=r, column=3), fmt=PCT)
        cell_style(ws.cell(row=r, column=4), fmt="0")
        ws.row_dimensions[r].height = 18
    cell_style(ws.cell(row=liab_total_row, column=1, value="TOTAL EXTERNAL LIABILITIES"), bold=True, bg=GREEN)
    cell_style(ws.cell(row=liab_total_row, column=2), bold=True, bg=GREEN, fmt=INR)
    cell_style(ws.cell(row=liab_total_row, column=3), bold=True, bg=GREEN, fmt=PCT)
    cell_style(ws.cell(row=liab_total_row, column=4), bold=True, bg=GREEN, fmt="0")

    band(ws, 37, "NET WORTH COMPOSITION  (equity accounts on the Liabilities side)", 4, bg=TEAL)
    table_header(ws, 38, ["Equity Component", "Amount (₹)", "% of Net Worth", "Source Group"])
    eq_first, eq_last = 39, 39 + len(EQUITY_CATEGORIES) - 1
    eq_total_row, eq_check_row = eq_last + 1, eq_last + 2
    eq_groups = {
        "Owner's Capital": "CAPITAL A/C",
        "Reserves & Surplus": "CAPITAL RESERVE AND SURPLUS A/C",
        "Unappropriated Income": "INCOME A/C",
    }
    for idx, name in enumerate(EQUITY_CATEGORIES):
        r = eq_first + idx
        cell_style(ws.cell(row=r, column=1, value=name))
        cell_style(ws.cell(row=r, column=2), fmt=INR)
        cell_style(ws.cell(row=r, column=3), fmt=PCT)
        cell_style(ws.cell(row=r, column=4, value=eq_groups[name]), size=9)
        ws.row_dimensions[r].height = 18
    cell_style(ws.cell(row=eq_total_row, column=1, value="TOTAL EQUITY = NET WORTH"), bold=True, bg=GREEN)
    cell_style(ws.cell(row=eq_total_row, column=2), bold=True, bg=GREEN, fmt=INR)
    cell_style(ws.cell(row=eq_total_row, column=3), bold=True, bg=GREEN, fmt=PCT)
    cell_style(ws.cell(row=eq_total_row, column=4), bold=True, bg=GREEN)
    cell_style(ws.cell(row=eq_check_row, column=1, value="Check: Equity − Net Worth (must be 0.00)"), size=9, italic=True)
    cell_style(ws.cell(row=eq_check_row, column=2), fmt=INR, size=9, italic=True)

    band(ws, 45, "KEY RATIOS", 4, bg=NAVY)
    table_header(ws, 46, ["Ratio", "Value", "", "What it shows"])
    ratio_notes = {
        47: "Share of the asset base funded by the owner",
        48: "Rupees of external debt for every rupee of net worth",
        49: "Share of the asset base funded by debt",
        50: "Concentration in loans/advances given to third parties",
        51: "Readily available cash & bank balances",
    }
    ratio_labels = {
        47: "Net Worth / Total Assets",
        48: "Debt / Equity",
        49: "Debt / Total Assets",
        50: "Loans & Advances Given / Total Assets",
        51: "Cash & Bank / Total Assets",
    }
    for r in range(47, 52):
        cell_style(ws.cell(row=r, column=1, value=ratio_labels[r]))
        cell_style(ws.cell(row=r, column=2), fmt=NUM_X if r == 48 else PCT)
        cell_style(ws.cell(row=r, column=3), size=9, color="5A6B7B")
        cell_style(ws.cell(row=r, column=4, value=ratio_notes[r]), size=9, color="5A6B7B")
        ws.row_dimensions[r].height = 18

    band(ws, 53, "BASIS AT A GLANCE", 4, bg=TEAL)
    quick_notes = [
        "Net Worth = Total Assets − Total External Liabilities = Owner's Capital + Reserves + untransferred Income.",
        "Equity accounts on the Liabilities side are excluded from external debt; all other credit balances are included.",
        "Asset-side balances are taken at book value; contra / credit balances were kept on the liabilities side (conservative).",
        "See 'Notes & Method' for classification rules, source quirks and reconciliation checks.",
    ]
    for i, text in enumerate(quick_notes):
        r = 54 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        c = ws.cell(row=r, column=1, value="•  " + text)
        cell_style(c, size=9, border=False, wrap=False)
        ws.row_dimensions[r].height = 16

    ws.freeze_panes = "A10"
    setup_print(ws, header_rows="1:9")

    # ---------------- Sheet 2: Calculation ----------------
    wc = bk.sheet(
        "Net Worth Calculation",
        {"A": 52, "B": 22, "C": 66},
        tab=TEAL,
    )
    band(wc, 1, "NET WORTH — CALCULATION & RECONCILIATION", 3, size=14, height=30)
    wc.merge_cells("A2:C2")
    wc["A2"] = f"{ENTITY}   |   As on {AS_ON}   |   Amounts in INR (₹)"
    cell_style(wc["A2"], border=False, size=10, color="5A6B7B")
    section_heading(wc, 4, "STEP 1 — TOTAL ASSETS", 3)
    cell_style(wc.cell(row=5, column=1, value="Total Assets (asset side of the Balance Sheet)"))
    cell_style(wc.cell(row=5, column=2), fmt=INR)
    cell_style(wc.cell(row=5, column=3, value="Links to 'Assets Detail' total."), size=9, color="5A6B7B")
    section_heading(wc, 7, "STEP 2 — TOTAL EXTERNAL LIABILITIES", 3)
    table_header(wc, 8, ["Liability Category", "Amount (₹)", "Basis"])
    for idx, name in enumerate(LIABILITY_CATEGORIES):
        r = 9 + idx
        cell_style(wc.cell(row=r, column=1, value=name))
        cell_style(wc.cell(row=r, column=2), fmt=INR)
        cell_style(
            wc.cell(row=r, column=3, value="From 'Net Worth Summary' → Liability Mix"),
            size=9,
            color="5A6B7B",
        )
    ext_total_row = 9 + len(LIABILITY_CATEGORIES)
    cell_style(wc.cell(row=ext_total_row, column=1, value="TOTAL EXTERNAL LIABILITIES"), bold=True, bg=GREEN)
    cell_style(wc.cell(row=ext_total_row, column=2), bold=True, bg=GREEN, fmt=INR)
    cell_style(wc.cell(row=ext_total_row, column=3), bold=True, bg=GREEN)
    section_heading(wc, ext_total_row + 2, "STEP 3 — NET WORTH", 3)
    nw_row = ext_total_row + 3
    cell_style(wc.cell(row=nw_row, column=1, value="NET WORTH  =  Total Assets − External Liabilities"), bold=True, size=12, color=NAVY, bg=AMBER)
    cell_style(wc.cell(row=nw_row, column=2), bold=True, size=12, color=NAVY, bg=AMBER, fmt=INR)
    cell_style(wc.cell(row=nw_row, column=3, value="Owner's stake in the business / personal balance sheet."), size=9, bg=AMBER)
    wc.row_dimensions[nw_row].height = 26
    section_heading(wc, nw_row + 2, "STEP 4 — RECONCILIATION WITH THE EQUITY ACCOUNTS IN THE BALANCE SHEET", 3)
    eq_hdr = nw_row + 3
    table_header(wc, eq_hdr, ["Equity Account", "Amount (₹)", "Source Group"])
    for idx, name in enumerate(EQUITY_CATEGORIES):
        r = eq_hdr + 1 + idx
        cell_style(wc.cell(row=r, column=1, value=name))
        cell_style(wc.cell(row=r, column=2), fmt=INR)
        cell_style(wc.cell(row=r, column=3, value=eq_groups[name]), size=9, color="5A6B7B")
    eq_sum_row = eq_hdr + 1 + len(EQUITY_CATEGORIES)
    cell_style(wc.cell(row=eq_sum_row, column=1, value="Equity per Balance Sheet"), bold=True, bg=GREEN)
    cell_style(wc.cell(row=eq_sum_row, column=2), bold=True, bg=GREEN, fmt=INR)
    cell_style(wc.cell(row=eq_sum_row, column=3), bold=True, bg=GREEN)
    diff_row = eq_sum_row + 1
    cell_style(wc.cell(row=diff_row, column=1, value="Difference: Net Worth − Equity (must be 0.00)"), size=9, italic=True)
    cell_style(wc.cell(row=diff_row, column=2), fmt=INR, size=9, italic=True)

    section_heading(wc, diff_row + 2, "MEMO ITEMS & CROSS-CHECKS", 3)
    memo = [
        ("Book capital (Capital + Reserves only, before untransferred income)", None),
        ("Net profit transferred from P&L (included in Reserves & Surplus)", None),
        ("Income A/C not yet transferred to P&L (Salary Income)", None),
        ("Balance Sheet grand total — Liabilities side", None),
        ("Balance Sheet grand total — Assets side", None),
        ("Check: Total Assets − B/S grand total (must be 0.00)", None),
        ("Check: External Liabilities + Equity − B/S total (must be 0.00)", None),
    ]
    memo_start = diff_row + 3
    for idx, (label, _) in enumerate(memo):
        r = memo_start + idx
        cell_style(wc.cell(row=r, column=1, value=label), size=9 if idx >= 5 else 10)
        cell_style(wc.cell(row=r, column=2), fmt=INR, size=9 if idx >= 5 else 10)

    method_row = memo_start + len(memo) + 2
    section_heading(wc, method_row, "METHOD", 3)
    note_lines(
        wc,
        method_row + 1,
        [
            "• Total Assets = all balances reported on the Assets (right) side of the source Balance Sheet.",
            "• External Liabilities = every balance on the Liabilities (left) side except the three equity accounts.",
            "• Net Worth = Total Assets − External Liabilities, cross-checked against Capital + Reserves + Income.",
            "• Because the source Balance Sheet balances, the two computations agree exactly (see difference rows above).",
        ],
        last_col=3,
    )
    wc.freeze_panes = "A5"
    setup_print(wc, header_rows="1:4")

    # ---------------- Sheet 3: Assets Detail ----------------
    wa = bk.sheet(
        "Assets Detail",
        {"A": 36, "B": 34, "C": 56, "D": 20, "E": 15, "F": 12},
        tab="2E8B57",
    )
    band(wa, 1, "ASSETS — DETAILED SCHEDULE", 6, size=14, height=30)
    wa.merge_cells("A2:F2")
    wa["A2"] = f"{ENTITY}   |   As on {AS_ON}   |   Amounts in INR (₹)   |   Source: Assets side of the Balance Sheet"
    cell_style(wa["A2"], border=False, size=9, color="5A6B7B")
    table_header(
        wa, 4,
        ["Asset Category", "Source Group (B/S)", "Account / Particulars", "Amount (₹)", "% of Total Assets", "Source Page"],
    )
    a_first = 5
    for idx, rec in enumerate(assets):
        r = a_first + idx
        cell_style(wa.cell(row=r, column=1, value=rec["category"]), size=9)
        cell_style(wa.cell(row=r, column=2, value=rec["group"]), size=9)
        cell_style(wa.cell(row=r, column=3, value=rec["name"]), size=9)
        bk.value(wa, f"D{r}", rec["amount"], INR)
        cell_style(wa[f"D{r}"], size=9, align="right", fmt=INR)
        cell_style(wa.cell(row=r, column=6, value=rec["page"]), size=9, align="center")
        wa.row_dimensions[r].height = 16
    a_last = a_first + len(assets) - 1
    a_total = a_last + 1
    cell_style(wa.cell(row=a_total, column=1, value="TOTAL ASSETS"), bold=True, bg=GREEN)
    for col in range(2, 7):
        cell_style(wa.cell(row=a_total, column=col), bold=True, bg=GREEN)
    for r in range(a_first, a_last + 1):
        bk.formula(wa, f"E{r}", f"=D{r}/$D${a_total}", assets[r - a_first]["amount"] / A, PCT)
        cell_style(wa[f"E{r}"], size=9, align="right", fmt=PCT)
    bk.formula(wa, f"D{a_total}", f"=SUM(D{a_first}:D{a_last})", A, INR)
    cell_style(wa[f"D{a_total}"], bold=True, bg=GREEN, align="right", fmt=INR)
    bk.formula(wa, f"E{a_total}", f"=SUM(E{a_first}:E{a_last})", Decimal(1), PCT)
    cell_style(wa[f"E{a_total}"], bold=True, bg=GREEN, align="right", fmt=PCT)
    wa.auto_filter.ref = f"A4:F{a_last}"
    wa.freeze_panes = "A5"
    setup_print(wa, header_rows="1:4")

    # Group reconciliation under the asset detail
    rec_hdr = a_total + 2
    section_heading(wa, rec_hdr, "RECONCILIATION TO THE SOURCE BALANCE SHEET GROUPS", 6)
    table_header(
        wa, rec_hdr + 1,
        ["Source Group", "Group Total (PDF)", "Sum of Lines (above)", "Difference", "Accounts", "Page(s)"],
    )
    asset_groups = [g for g in groups if g["side"] == "Assets"]
    for idx, g in enumerate(asset_groups):
        r = rec_hdr + 2 + idx
        s = sum((i["amount"] for i in g["items"]), Decimal("0"))
        cell_style(wa.cell(row=r, column=1, value=g["name"]), size=9)
        bk.value(wa, f"B{r}", g["total"], INR)
        cell_style(wa[f"B{r}"], size=9, align="right", fmt=INR)
        bk.formula(
            wa, f"C{r}",
            f'=SUMIF($B${a_first}:$B${a_last},A{r},$D${a_first}:$D${a_last})',
            s, INR,
        )
        cell_style(wa[f"C{r}"], size=9, align="right", fmt=INR)
        bk.formula(wa, f"D{r}", f"=C{r}-B{r}", Decimal(0), INR)
        cell_style(wa[f"D{r}"], size=9, align="right", fmt=INR)
        cell_style(wa.cell(row=r, column=5, value=len(g["items"])), size=9, align="center")
        cell_style(wa.cell(row=r, column=6, value=", ".join(map(str, g["pages"]))), size=9, align="center")
    a_rec_last = rec_hdr + 1 + len(asset_groups)

    # ---------------- Sheet 4: Liabilities Detail ----------------
    wl = bk.sheet(
        "Liabilities Detail",
        {"A": 20, "B": 34, "C": 38, "D": 52, "E": 20, "F": 12},
        tab="B8860B",
    )
    band(wl, 1, "LIABILITIES & EQUITY — DETAILED SCHEDULE", 6, size=14, height=30)
    wl.merge_cells("A2:F2")
    wl["A2"] = f"{ENTITY}   |   As on {AS_ON}   |   Amounts in INR (₹)   |   Source: Liabilities side of the Balance Sheet"
    cell_style(wl["A2"], border=False, size=9, color="5A6B7B")
    table_header(
        wl, 4,
        ["Classification", "Category", "Source Group (B/S)", "Account / Particulars", "Amount (₹)", "Source Page"],
    )
    l_first = 5
    for idx, rec in enumerate(left):
        r = l_first + idx
        cell_style(wl.cell(row=r, column=1, value=rec["classification"]), size=9,
                   bold=rec["classification"] == "Equity",
                   color="1E6B3A" if rec["classification"] == "Equity" else INK)
        cell_style(wl.cell(row=r, column=2, value=rec["category"]), size=9)
        cell_style(wl.cell(row=r, column=3, value=rec["group"]), size=9)
        cell_style(wl.cell(row=r, column=4, value=rec["name"]), size=9)
        bk.value(wl, f"E{r}", rec["amount"], INR)
        cell_style(wl[f"E{r}"], size=9, align="right", fmt=INR)
        cell_style(wl.cell(row=r, column=6, value=rec["page"]), size=9, align="center")
        wl.row_dimensions[r].height = 16
    l_last = l_first + len(left) - 1
    l_total = l_last + 1
    labels = {
        l_total: "TOTAL — LIABILITIES SIDE (as per Balance Sheet)",
        l_total + 1: "of which: EXTERNAL LIABILITIES (deducted from assets)",
        l_total + 2: "of which: EQUITY / NET WORTH (capital, reserves, income)",
        l_total + 3: "Check: Total − External − Equity (must be 0.00)",
    }
    for r, label in labels.items():
        cell_style(wl.cell(row=r, column=1, value=label), bold=True, bg=GREEN)
        wl.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        for col in range(2, 7):
            cell_style(wl.cell(row=r, column=col), bold=True, bg=GREEN)
    bk.formula(wl, f"E{l_total}", f"=SUM(E{l_first}:E{l_last})", GT, INR)
    bk.formula(wl, f"E{l_total+1}",
               f'=SUMIF($A${l_first}:$A${l_last},"External Liability",$E${l_first}:$E${l_last})', L, INR)
    bk.formula(wl, f"E{l_total+2}",
               f'=SUMIF($A${l_first}:$A${l_last},"Equity",$E${l_first}:$E${l_last})', E, INR)
    bk.formula(wl, f"E{l_total+3}", f"=E{l_total}-E{l_total+1}-E{l_total+2}", Decimal(0), INR)
    for r in labels:
        cell_style(wl[f"E{r}"], bold=True, bg=GREEN, align="right", fmt=INR)
    wl.auto_filter.ref = f"A4:F{l_last}"
    wl.freeze_panes = "A5"
    setup_print(wl, header_rows="1:4")

    rec_hdr_l = l_total + 5
    section_heading(wl, rec_hdr_l, "RECONCILIATION TO THE SOURCE BALANCE SHEET GROUPS", 6)
    table_header(
        wl, rec_hdr_l + 1,
        ["Source Group", "Group Total (PDF)", "Sum of Lines (above)", "Difference", "Accounts", "Page(s)"],
    )
    left_groups = [g for g in groups if g["side"] == "Liabilities"]
    for idx, g in enumerate(left_groups):
        r = rec_hdr_l + 2 + idx
        s = sum((i["amount"] for i in g["items"]), Decimal("0"))
        cell_style(wl.cell(row=r, column=1, value=g["name"]), size=9)
        bk.value(wl, f"B{r}", g["total"], INR)
        cell_style(wl[f"B{r}"], size=9, align="right", fmt=INR)
        bk.formula(
            wl, f"C{r}",
            f'=SUMIF($C${l_first}:$C${l_last},A{r},$E${l_first}:$E${l_last})',
            s, INR,
        )
        cell_style(wl[f"C{r}"], size=9, align="right", fmt=INR)
        bk.formula(wl, f"D{r}", f"=C{r}-B{r}", Decimal(0), INR)
        cell_style(wl[f"D{r}"], size=9, align="right", fmt=INR)
        cell_style(wl.cell(row=r, column=5, value=len(g["items"])), size=9, align="center")
        cell_style(wl.cell(row=r, column=6, value=", ".join(map(str, g["pages"]))), size=9, align="center")

    # ---------------- Sheet 5: Source transcript ----------------
    wsrc = bk.sheet(
        "Balance Sheet (Source)",
        {"A": 52, "B": 20, "C": 56, "D": 20, "E": 20, "F": 18},
        tab="6B7280",
    )
    band(wsrc, 1, "BALANCE SHEET — FAITHFUL TRANSCRIPT OF THE SOURCE PDF", 6, size=14, height=30)
    wsrc.merge_cells("A2:F2")
    wsrc["A2"] = (
        f"{ENTITY}   |   BALANCE SHEET AS ON : {AS_ON}   |   "
        f"Source: {pdf_name}   |   Account names reproduced as printed (typos included)"
    )
    cell_style(wsrc["A2"], border=False, size=9, color="5A6B7B")
    table_header(
        wsrc, 4,
        ["LIABILITIES — Name of the Accounts", "Amount (₹)",
         "ASSETS — Name of the Accounts", "Amount (₹)", "", ""],
    )

    def side_lines(side: str):
        out = []
        for g in [x for x in groups if x["side"] == side]:
            out.append(("group", g["name"], g["total"]))
            for i in g["items"]:
                out.append(("item", "    " + i["name"], i["amount"]))
        return out

    l_lines, r_lines = side_lines("Liabilities"), side_lines("Assets")
    src_first = 5
    for idx, (ll, rr) in enumerate(zip_longest(l_lines, r_lines)):
        r = src_first + idx
        if ll:
            kind, name, amt = ll
            c = wsrc.cell(row=r, column=1, value=name)
            cell_style(c, bold=(kind == "group"), size=9,
                       bg=LIGHT_BLUE if kind == "group" else None)
            av = wsrc.cell(row=r, column=2, value=float(amt))
            cell_style(av, bold=(kind == "group"), size=9, align="right", fmt=INR,
                       bg=LIGHT_BLUE if kind == "group" else None)
        if rr:
            kind, name, amt = rr
            c = wsrc.cell(row=r, column=3, value=name)
            cell_style(c, bold=(kind == "group"), size=9,
                       bg=LIGHT_GREY if kind == "group" else None)
            av = wsrc.cell(row=r, column=4, value=float(amt))
            cell_style(av, bold=(kind == "group"), size=9, align="right", fmt=INR,
                       bg=LIGHT_GREY if kind == "group" else None)
        wsrc.row_dimensions[r].height = 15
    src_last = src_first + max(len(l_lines), len(r_lines)) - 1
    trow = src_last + 1
    for col, label in ((1, "T O T A L :"), (3, "T O T A L :")):
        c = wsrc.cell(row=trow, column=col, value=label)
        cell_style(c, bold=True, bg=NAVY, color=WHITE)
    bk.value(wsrc, f"B{trow}", grand["Liabilities"], INR)
    bk.value(wsrc, f"D{trow}", grand["Assets"], INR)
    for coord in (f"B{trow}", f"D{trow}"):
        cell_style(wsrc[coord], bold=True, bg=NAVY, color=WHITE, align="right", fmt=INR)
    for col in (5, 6):
        cell_style(wsrc.cell(row=trow, column=col), bold=True, bg=NAVY)
    wsrc.row_dimensions[trow].height = 20

    # recomputed totals from group rows only
    group_rows = {
        side: [r for r in range(src_first, src_last + 1)
               if wsrc.cell(row=r, column=1 if side == "Liabilities" else 3).value
               and not str(wsrc.cell(row=r, column=1 if side == "Liabilities" else 3).value).startswith("    ")]
        for side in ("Liabilities", "Assets")
    }
    chk = trow + 1
    cell_style(wsrc.cell(row=chk, column=1, value="Sum of group totals (recomputed)"), bold=True, bg=GREEN)
    cell_style(wsrc.cell(row=chk, column=2), bold=True, bg=GREEN)
    cell_style(wsrc.cell(row=chk, column=3, value="Sum of group totals (recomputed)"), bold=True, bg=GREEN)
    cell_style(wsrc.cell(row=chk, column=4), bold=True, bg=GREEN)
    bk.formula(wsrc, f"B{chk}",
               "=" + "+".join(f"B{r}" for r in group_rows["Liabilities"]),
               grand["Liabilities"], INR)
    bk.formula(wsrc, f"D{chk}",
               "=" + "+".join(f"D{r}" for r in group_rows["Assets"]),
               grand["Assets"], INR)
    for coord in (f"B{chk}", f"D{chk}"):
        cell_style(wsrc[coord], bold=True, bg=GREEN, align="right", fmt=INR)
    dchk = trow + 2
    cell_style(wsrc.cell(row=dchk, column=1, value="Difference vs printed total (must be 0.00)"), size=9, italic=True)
    cell_style(wsrc.cell(row=dchk, column=3, value="Difference vs printed total (must be 0.00)"), size=9, italic=True)
    bk.formula(wsrc, f"B{dchk}", f"=B{chk}-B{trow}", Decimal(0), INR)
    bk.formula(wsrc, f"D{dchk}", f"=D{chk}-D{trow}", Decimal(0), INR)
    for coord in (f"B{dchk}", f"D{dchk}"):
        cell_style(wsrc[coord], size=9, italic=True, align="right", fmt=INR)

    val_hdr = dchk + 2
    section_heading(wsrc, val_hdr, "FULL GROUP-BY-GROUP VALIDATION  (printed total vs sum of account lines)", 6)
    table_header(
        wsrc, val_hdr + 1,
        ["Side", "Group / Account Head", "Printed Group Total", "Sum of Account Lines", "Difference", "Page(s)"],
    )
    for idx, c in enumerate(checks):
        r = val_hdr + 2 + idx
        cell_style(wsrc.cell(row=r, column=1, value=c["side"]), size=9)
        cell_style(wsrc.cell(row=r, column=2, value=c["group"]), size=9)
        bk.value(wsrc, f"C{r}", c["pdf_total"], INR)
        bk.value(wsrc, f"D{r}", c["items_sum"], INR)
        bk.value(wsrc, f"E{r}", c["diff"], INR)
        for col in "CDE":
            cell_style(wsrc[f"{col}{r}"], size=9, align="right", fmt=INR)
        cell_style(wsrc.cell(row=r, column=6, value=", ".join(map(str, c["pages"]))), size=9, align="center")
    v_last = val_hdr + 1 + len(checks)
    wsrc.freeze_panes = "A5"
    setup_print(wsrc, header_rows="1:4")

    # ---------------- Sheet 6: Notes & Method ----------------
    wn = bk.sheet(
        "Notes & Method",
        {c: 15 for c in "ABCDEFGH"},
        tab="6B7280",
    )
    band(wn, 1, "NOTES, METHOD & DATA QUALITY", 8, size=14, height=30)
    wn.merge_cells("A2:H2")
    wn["A2"] = f"Net Worth Report — {ENTITY} — as on {AS_ON}"
    cell_style(wn["A2"], border=False, size=10, color="5A6B7B")

    r = 4
    section_heading(wn, r, "1.  PURPOSE & DEFINITION", 8); r += 1
    r = note_lines(wn, r, [
        "This workbook converts the two-page Balance Sheet PDF into an Excel net worth statement with detailed schedules.",
        "Net Worth = Total Assets − Total External Liabilities.  It is cross-checked against the owner's equity accounts.",
        f"Result: Total Assets ₹ {inr(A)}  −  External Liabilities ₹ {inr(L)}  =  Net Worth ₹ {inr(NW)}.",
    ], 8) + 1

    section_heading(wn, r, "2.  SOURCE DATA", 8); r += 1
    r = note_lines(wn, r, [
        f"Source file: {pdf_name} — 'BALANCE SHEET AS ON : {AS_ON}' for {ENTITY} (Oracle Reports PDF, 2 pages).",
        f"Printed grand total: Liabilities ₹ {inr(grand['Liabilities'])} / Assets ₹ {inr(grand['Assets'])} (both sides agree).",
        f"Parsed: {len(groups)} account groups, {sum(len(g['items']) for g in groups)} account lines. All amounts are in INR (₹).",
        "Account names, spellings and abbreviations are reproduced exactly as printed in the source.",
    ], 8) + 1

    section_heading(wn, r, "3.  CLASSIFICATION RULES", 8); r += 1
    r = note_lines(wn, r, [
        "Assets — all groups printed on the right-hand ASSETS side of the Balance Sheet, split into 8 asset classes.",
        "   • Land & Property: agricultural land, plots, residential property construction; Gold & Valuables; Vehicles.",
        "   • Investments: mutual funds, shares, share-trading account; Deposits & Fixed Deposits: bank FD, gas/rent/electricity deposits.",
        "   • Loans & Advances Given: all loan/advance debit balances; Receivables & Tax Recoverable: sundry debtor and TDS/professional tax balances.",
        "   • Cash & Bank Balances: all five savings/current bank accounts.",
        "External Liabilities — all groups on the left-hand LIABILITIES side EXCEPT the three equity accounts below.",
        "Equity (excluded from debt) — CAPITAL A/C, CAPITAL RESERVE AND SURPLUS A/C and INCOME A/C (salary income not yet",
        "   transferred to P&L). These three balances together equal the net worth computed as Assets − Liabilities.",
        "Conservative treatment: credit balances of asset-natured heads appearing on the liabilities side (ASSETS A/C,",
        "   INVESTMENT A/C, SUNDRY DEBTORS/CREDITORS A/C, CASH BOOK A/C) are reported as 'Other Credit / Contra Balances'.",
        "   Treating them as negative assets instead would change gross Assets/Liabilities but NOT the net worth figure.",
    ], 8) + 1

    section_heading(wn, r, "4.  DATA QUALITY NOTES (SOURCE QUIRKS)", 8); r += 1
    r = note_lines(wn, r, [
        "• The credit-card line 'CREDIT CARD EXP (HDFC)-3988 A/C' is printed as '.37' (no leading zero); read as ₹ 0.37.",
        "   Its two lines (0.37 + 11,734.00) reconcile to the group 'EXPENSE(CAPITAL)A/C' total of 11,734.37.",
        "• The Assets-side group 'LOAN & ADVANCES A/C' (₹ 4,78,28,472.18) is printed on BOTH pages; the repeated heading",
        "   was merged and its 10 account lines combined (3 lines from page 1, 7 from page 2).",
        "• Multi-line account names (e.g. 'AGRI. LAND PURCHASE(ASTGAM-801)1/4 A/C') were rejoined without adding",
        "   spaces where the PDF used a hyphen as a line break.",
        "• Obvious source spellings such as 'BOB HOUSIONG LOAN', 'DRIV TRUKING', 'KASHMIRA FEBRICS' and",
        "   'DRIVE TRUSCKING SHARE' are preserved verbatim for traceability.",
        "• 'INCOME A/C' (Salary Income ₹ 4,22,400.00) is a nominal account still open on the Balance Sheet; it is",
        "   included in net worth. Book capital excluding it is shown as a memo item on 'Net Worth Calculation'.",
        "• The 'as on' date is reproduced verbatim as 10/09/2026; the PDF does not state whether it is dd/mm or mm/dd.",
        "   The source filename contains 23102026, which appears to be the print/export date, not the 'as on' date.",
    ], 8) + 1

    section_heading(wn, r, "5.  VALIDATION PERFORMED", 8); r += 1
    r = note_lines(wn, r, [
        f"• Every one of the {len(checks)} groups: sum of its account lines = printed group total (difference 0.00).",
        f"• Sum of all group totals on each side = printed grand total ₹ {inr(GT)}.",
        "• Assets − External Liabilities = Capital + Reserves + Income (difference 0.00).",
        "• See 'Balance Sheet (Source)' for the group-by-group validation table and the full two-column transcript.",
    ], 8) + 1

    section_heading(wn, r, "6.  LIMITATIONS & DISCLAIMER", 8); r += 1
    r = note_lines(wn, r, [
        "• Figures are BOOK VALUES from the source Balance Sheet, not current market valuations of land, property, gold,",
        "   shares or mutual funds. Recoverability of loans/advances and receivables has not been assessed.",
        "• Classification into asset/liability classes is an analytical view prepared for this report; the grouping in the",
        "   source Balance Sheet is authoritative for accounting purposes.",
        "• This is an information worksheet, not an audited financial statement, valuation report or tax/legal advice.",
        "   Please have it reviewed by a qualified accountant before submitting it to a bank, lender or authority.",
    ], 8) + 1

    section_heading(wn, r, "7.  HOW TO REGENERATE", 8); r += 1
    r = note_lines(wn, r, [
        "• Run:  python3 -m pip install -r requirements.txt  &&  python3 generate_networth_report.py",
        "• The script re-parses the PDF, re-runs every reconciliation check and rebuilds this workbook from scratch.",
    ], 8) + 2

    section_heading(wn, r, "SHEET GUIDE", 8); r += 1
    for i, (label, loc) in enumerate([
        ("Net Worth Summary", "'Net Worth Summary'!A1"),
        ("Net Worth Calculation", "'Net Worth Calculation'!A1"),
        ("Assets Detail", "'Assets Detail'!A1"),
        ("Liabilities Detail", "'Liabilities Detail'!A1"),
        ("Balance Sheet (Source)", "'Balance Sheet (Source)'!A1"),
        ("Notes & Method", "'Notes & Method'!A1"),
    ]):
        internal_link(wn, f"A{r + i}", label, loc)
        wn.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=3)
    setup_print(wn)

    # ---------------- Now wire up all formulas ----------------
    # Summary headline
    bk.formula(ws, "B10", f"='Assets Detail'!D{a_total}", A, INR)
    bk.formula(ws, "B11", f"='Liabilities Detail'!E{l_total+1}", L, INR)
    bk.formula(ws, "B12", "=B10-B11", NW, INR)
    bk.formula(ws, "B13", "=B10-B11-B12", Decimal(0), INR)
    for coord, val in [("C10", A / A), ("C11", L / A), ("C12", NW / A)]:
        bk.formula(ws, coord, f"=B{coord[1:]}/$B$10", val, PCT)
        cell_style(ws[coord], align="right", fmt=PCT)
    for r in (10, 11, 12, 13):
        cell_style(ws[f"B{r}"], bold=(r == 12), align="right", fmt=INR)
        if r == 12:
            ws[f"B{r}"].fill = PatternFill("solid", fgColor=AMBER)
    cell_style(ws["B13"], size=9, italic=True, align="right", fmt=INR)
    cell_style(ws["C13"], size=9, color="5A6B7B")

    # Asset mix
    for idx, name in enumerate(ASSET_CATEGORIES):
        r = asset_first + idx
        v = cat[name]
        bk.formula(
            ws, f"B{r}",
            f"=SUMIF('Assets Detail'!$A${a_first}:$A${a_last},A{r},"
            f"'Assets Detail'!$D${a_first}:$D${a_last})",
            v, INR,
        )
        bk.formula(ws, f"C{r}", f"=B{r}/$B$10", v / A, PCT)
        bk.formula(ws, f"D{r}",
                   f"=COUNTIF('Assets Detail'!$A${a_first}:$A${a_last},A{r})",
                   sum(1 for x in assets if x["category"] == name), "0")
        for col in "BCD":
            cell_style(ws[f"{col}{r}"], align="right", fmt={"B": INR, "C": PCT, "D": "0"}[col])
    bk.formula(ws, f"B{asset_total_row}", f"=SUM(B{asset_first}:B{asset_last})", A, INR)
    bk.formula(ws, f"C{asset_total_row}", f"=B{asset_total_row}/$B$10", Decimal(1), PCT)
    bk.formula(ws, f"D{asset_total_row}", f"=SUM(D{asset_first}:D{asset_last})", len(assets), "0")
    for col in "BCD":
        cell_style(ws[f"{col}{asset_total_row}"], bold=True, bg=GREEN, align="right",
                   fmt={"B": INR, "C": PCT, "D": "0"}[col])

    # Liability mix
    for idx, name in enumerate(LIABILITY_CATEGORIES):
        r = liab_first + idx
        v = cat[name]
        bk.formula(
            ws, f"B{r}",
            f"=SUMIF('Liabilities Detail'!$B${l_first}:$B${l_last},A{r},"
            f"'Liabilities Detail'!$E${l_first}:$E${l_last})",
            v, INR,
        )
        bk.formula(ws, f"C{r}", f"=B{r}/$B$10", v / A, PCT)
        bk.formula(ws, f"D{r}",
                   f"=COUNTIF('Liabilities Detail'!$B${l_first}:$B${l_last},A{r})",
                   sum(1 for x in left if x["category"] == name), "0")
        for col in "BCD":
            cell_style(ws[f"{col}{r}"], align="right", fmt={"B": INR, "C": PCT, "D": "0"}[col])
    bk.formula(ws, f"B{liab_total_row}", f"=SUM(B{liab_first}:B{liab_last})", L, INR)
    bk.formula(ws, f"C{liab_total_row}", f"=B{liab_total_row}/$B$10", L / A, PCT)
    bk.formula(ws, f"D{liab_total_row}", f"=SUM(D{liab_first}:D{liab_last})",
               sum(1 for x in left if x["classification"] == "External Liability"), "0")
    for col in "BCD":
        cell_style(ws[f"{col}{liab_total_row}"], bold=True, bg=GREEN, align="right",
                   fmt={"B": INR, "C": PCT, "D": "0"}[col])

    # Equity composition
    for idx, name in enumerate(EQUITY_CATEGORIES):
        r = eq_first + idx
        v = cat[name]
        bk.formula(
            ws, f"B{r}",
            f"=SUMIF('Liabilities Detail'!$B${l_first}:$B${l_last},A{r},"
            f"'Liabilities Detail'!$E${l_first}:$E${l_last})",
            v, INR,
        )
        bk.formula(ws, f"C{r}", f"=B{r}/$B$12", v / NW, PCT)
        for col in "BC":
            cell_style(ws[f"{col}{r}"], align="right", fmt={"B": INR, "C": PCT}[col])
    bk.formula(ws, f"B{eq_total_row}", f"=SUM(B{eq_first}:B{eq_last})", E, INR)
    bk.formula(ws, f"C{eq_total_row}", f"=B{eq_total_row}/$B$12", Decimal(1), PCT)
    for col in "BC":
        cell_style(ws[f"{col}{eq_total_row}"], bold=True, bg=GREEN, align="right",
                   fmt={"B": INR, "C": PCT}[col])
    bk.formula(ws, f"B{eq_check_row}", f"=B{eq_total_row}-B12", Decimal(0), INR)
    cell_style(ws[f"B{eq_check_row}"], size=9, italic=True, align="right", fmt=INR)

    # Key ratios (find category rows)
    loans_row = asset_first + ASSET_CATEGORIES.index("Loans & Advances Given")
    cash_row = asset_first + ASSET_CATEGORIES.index("Cash & Bank Balances")
    ratio_formulas = {
        47: ("=B12/B10", NW / A, PCT),
        48: ("=B11/B12", L / NW, NUM_X),
        49: ("=B11/B10", L / A, PCT),
        50: (f"=B{loans_row}/B10", cat["Loans & Advances Given"] / A, PCT),
        51: (f"=B{cash_row}/B10", cat["Cash & Bank Balances"] / A, PCT),
    }
    for r, (f, v, fmt) in ratio_formulas.items():
        bk.formula(ws, f"B{r}", f, v, fmt)
        cell_style(ws[f"B{r}"], align="right", fmt=fmt)

    # Calculation sheet formulas
    bk.formula(wc, "B5", "='Net Worth Summary'!B10", A, INR)
    cell_style(wc["B5"], align="right", fmt=INR)
    for idx in range(len(LIABILITY_CATEGORIES)):
        r = 9 + idx
        bk.formula(wc, f"B{r}", f"='Net Worth Summary'!B{liab_first+idx}",
                   cat[LIABILITY_CATEGORIES[idx]], INR)
        cell_style(wc[f"B{r}"], align="right", fmt=INR)
    bk.formula(wc, f"B{ext_total_row}", f"=SUM(B9:B{ext_total_row-1})", L, INR)
    cell_style(wc[f"B{ext_total_row}"], bold=True, bg=GREEN, align="right", fmt=INR)
    bk.formula(wc, f"B{nw_row}", f"=B5-B{ext_total_row}", NW, INR)
    cell_style(wc[f"B{nw_row}"], bold=True, size=12, color=NAVY, bg=AMBER, align="right", fmt=INR)
    for idx in range(len(EQUITY_CATEGORIES)):
        r = eq_hdr + 1 + idx
        bk.formula(wc, f"B{r}", f"='Net Worth Summary'!B{eq_first+idx}",
                   cat[EQUITY_CATEGORIES[idx]], INR)
        cell_style(wc[f"B{r}"], align="right", fmt=INR)
    bk.formula(wc, f"B{eq_sum_row}", f"=SUM(B{eq_hdr+1}:B{eq_sum_row-1})", E, INR)
    cell_style(wc[f"B{eq_sum_row}"], bold=True, bg=GREEN, align="right", fmt=INR)
    bk.formula(wc, f"B{diff_row}", f"=B{nw_row}-B{eq_sum_row}", Decimal(0), INR)
    cell_style(wc[f"B{diff_row}"], size=9, italic=True, align="right", fmt=INR)
    capital_v = cat["Owner's Capital"]
    reserves_v = cat["Reserves & Surplus"]
    income_v = cat["Unappropriated Income"]
    net_profit_v = sum(
        (i["amount"] for g in groups if g["name"] == "CAPITAL RESERVE AND SURPLUS A/C"
         for i in g["items"]), Decimal("0")
    )
    memo_values = [
        capital_v + reserves_v,
        net_profit_v,
        income_v,
        grand["Liabilities"],
        grand["Assets"],
        Decimal(0),
        Decimal(0),
    ]
    memo_formulas = [
        f"=B{eq_hdr+1}+B{eq_hdr+2}",
        f"='Net Worth Summary'!B{eq_first+1}",
        f"=B{eq_hdr+3}",
        "='Balance Sheet (Source)'!B" + str(trow),
        "='Balance Sheet (Source)'!D" + str(trow),
        f"=B5-B{memo_start+3}",
        f"=B{ext_total_row}+B{eq_sum_row}-B{memo_start+3}",
    ]
    for idx, (f, v) in enumerate(zip(memo_formulas, memo_values)):
        r = memo_start + idx
        bk.formula(wc, f"B{r}", f, v, INR)
        cell_style(wc[f"B{r}"], size=9 if idx >= 5 else 10, align="right", fmt=INR)

    # ---------------- Charts ----------------
    pie = PieChart()
    pie.title = "Asset Mix by Category"
    pie.height, pie.width = 8.6, 14.5
    pie.add_data(
        Reference(ws, min_col=2, min_row=16, max_row=asset_last), titles_from_data=True
    )
    pie.set_categories(Reference(ws, min_col=1, min_row=asset_first, max_row=asset_last))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.legend.position = "r"
    ws.add_chart(pie, "E9")

    bar = BarChart()
    bar.type = "bar"
    bar.style = 10
    bar.title = "Liability Mix by Category"
    bar.y_axis.title = None
    bar.x_axis.title = "₹"
    bar.height, bar.width = 8.6, 14.5
    bar.add_data(
        Reference(ws, min_col=2, min_row=28, max_row=liab_last), titles_from_data=True
    )
    bar.set_categories(Reference(ws, min_col=1, min_row=liab_first, max_row=liab_last))
    bar.legend = None
    ws.add_chart(bar, "E27")

    kpi = BarChart()
    kpi.type = "col"
    kpi.style = 10
    kpi.title = "Assets vs Liabilities vs Net Worth"
    kpi.y_axis.title = "₹"
    kpi.height, kpi.width = 8.6, 14.5
    kpi.add_data(Reference(ws, min_col=2, min_row=9, max_row=12), titles_from_data=True)
    kpi.set_categories(Reference(ws, min_col=1, min_row=10, max_row=12))
    kpi.legend = None
    ws.add_chart(kpi, "E45")

    bk.wb.active = 0
    bk.wb.properties.title = f"Net Worth Report — {ENTITY}"
    bk.wb.properties.subject = f"Net worth statement as on {AS_ON} from Balance Sheet PDF"
    bk.wb.properties.creator = "generate_networth_report.py"
    bk.wb.properties.description = (
        f"Total Assets {inr(A)}; External Liabilities {inr(L)}; Net Worth {inr(NW)} INR."
    )
    bk.wb.calculation.fullCalcOnLoad = True
    bk.wb.calculation.forceFullCalc = True
    bk.wb.calculation.calcMode = "auto"
    return bk


# ---------------------------------------------------------------------------
# 5. Main
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--pdf", type=Path, default=DEFAULT_PDF)
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = ap.parse_args()

    print(f"Reading: {args.pdf}")
    groups, grand = parse_pdf(args.pdf)
    print(f"Parsed {len(groups)} groups / {sum(len(g['items']) for g in groups)} account lines.")
    print("Reconciliation:")
    checks = validate(groups, grand)
    assets, left, totals = build_records(groups)

    print(f"\n  Total Assets              : ₹ {inr(totals['assets']):>18}")
    print(f"  External Liabilities      : ₹ {inr(totals['external_liabilities']):>18}")
    print(f"  NET WORTH                 : ₹ {inr(totals['net_worth']):>18}")
    print(f"  Equity accounts (check)   : ₹ {inr(totals['equity']):>18}")
    for name in ASSET_CATEGORIES:
        print(f"    A  {name:<38} ₹ {inr(totals['cat_tot'][name]):>16}")
    for name in LIABILITY_CATEGORIES:
        print(f"    L  {name:<38} ₹ {inr(totals['cat_tot'][name]):>16}")

    bk = build_workbook(groups, grand, checks, assets, left, totals, args.pdf.name)
    bk.wb.save(args.out)
    bk.patch_cached_values(args.out)
    print(f"\nWrote: {args.out}  ({args.out.stat().st_size:,} bytes)")
    print("Sheets: " + ", ".join(bk.wb.sheetnames))


if __name__ == "__main__":
    main()
