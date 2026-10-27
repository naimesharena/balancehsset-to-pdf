# Net Worth Report from Balance Sheet PDF

This repository contains a two-page balance sheet PDF for **Vipulbhai Laljibhai Patel**
(`Vipulbhai BALANCESHEET_1-23102026161349.PDF`, "BALANCE SHEET AS ON : 10/09/2026")
and a generated Excel net worth report.

## Deliverable

**`Vipulbhai_Patel_Net_Worth_Report_10-09-2026.xlsx`** — a 6-sheet, formula-driven
workbook (with cached values so it displays correctly even in viewers that do not
recalculate formulas).

| Sheet | Contents |
| --- | --- |
| Net Worth Summary | Headline figures, asset mix, liability mix, equity composition, key ratios, 3 charts |
| Net Worth Calculation | Step-by-step Assets − Liabilities bridge + reconciliation to the equity accounts |
| Assets Detail | All 48 asset account lines with category, source group, amount, % and source page + group reconciliation |
| Liabilities Detail | All 36 liabilities/equity account lines with classification + group reconciliation |
| Balance Sheet (Source) | Faithful two-column transcript of the PDF, printed grand totals and a 21-group validation table |
| Notes & Method | Definitions, classification rules, source quirks, validation, limitations, sheet guide |

## Headline result (as on 10/09/2026, INR ₹)

| Particular | Amount |
| --- | ---: |
| Total Assets | 7,39,01,963.50 |
| Total External Liabilities | 6,08,51,626.18 |
| **Net Worth** | **1,30,50,337.32** |
| Balance Sheet grand total (both sides) | 7,39,01,963.50 |

Net worth is verified two ways: `Assets − External Liabilities` and
`Capital A/C (1,24,53,086.52) + Reserves & Surplus (1,74,850.80) + untransferred
Income A/C (4,22,400.00)`. Book capital before the untransferred income is
1,26,27,937.32 (memo item in the Calculation sheet).

## Method (short version)

* **Assets** = every balance on the right-hand ASSETS side of the PDF, grouped into
  8 asset classes (land & property, gold, vehicles, investments, deposits/FDs,
  loans & advances given, receivables/TDS, cash & bank).
* **External Liabilities** = every balance on the left-hand LIABILITIES side except
  the three owner's-equity accounts (Capital, Reserves & Surplus, Income).
* **Net Worth = Assets − External Liabilities.**
* Credit balances of asset-natured heads that appear on the liabilities side are
  reported conservatively as "Other Credit / Contra Balances". Reclassifying them
  as negative assets would not change net worth.

All figures are **book values** from the source, not market valuations. See the
"Notes & Method" sheet for the full classification rules, source quirks
(e.g. the `.37` printed without a leading zero) and disclaimer.

## Regenerate

```bash
python3 -m pip install -r requirements.txt
python3 generate_networth_report.py
```

The script re-parses the PDF with `pdfplumber`, asserts that all 21 group sub-totals
and both grand totals reconcile, classifies all 84 account lines, and rebuilds the
workbook with `openpyxl`. Optional arguments: `--pdf PATH` and `--out PATH`.
