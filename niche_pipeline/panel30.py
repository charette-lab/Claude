"""
panel30.py — loader/adapter for the 30-year history panel.

The 30-year export is a richer schema than the 10-year panel: 175 columns with
underscore names ("New_Operating_Income", "Period_End_Date") and split across
several .xlsb files (one instrument lives entirely in one file). This module
stitches the parts into a single in-memory panel and aliases the columns to the
SPACED names the rest of the code (capital.py, history.py, overearning.py)
already uses, so nothing downstream needs to change.

Usage:
    by, idx = panel30.load(["part1.xlsb", "part2.xlsb", "part3.xlsb"])
    # by[ticker] -> rows sorted oldest->newest ; idx[spaced_name] -> column index
"""
from __future__ import annotations
import pyxlsb

# spaced name the code expects  ->  underscore name in the 30-year file
ALIAS = {
    "Date": "Period_End_Date",
    "Sales": "Total_Revenue__ASR",
    "New Operating Income": "New_Operating_Income",
    "Invested Capital": "Invested_Capital",
    "Ratio": "Ratio",
    "Ratio_rolling_avg_7yr": "Ratio_rolling_avg_7yr",
    "Ratio_rolling_avg_14yr": "Ratio_rolling_avg_14yr",
    "ROICm_total - 7 years": "ROICm_total___7_years",
    "ROICm_total - 14 years": "ROICm_total___14_years",
    "ROICm_total - 21 years": "ROICm_total___21_years",
    "average_C - 7 year": "average_C___7_year",
    "Marginal Investments_rolling_sum_7yr": "Marginal_Investments_rolling_sum_7yr",
    "Marginal Investments_rolling_sum_3yr": "Marginal_Investments_rolling_sum_3yr",
    "EBITA_Margin": "EBITA_Margin",
    "EBITA_Average_Margin 10 yr": "EBITA_Average_Margin_Last_10_yr",
    "Macro Weight": "Macro_Weight",
    "current_multiple": "current_multiple",
    "d_Direktavkastning": "d_Direktavkastning",
    "Sales Growth": "Sales_Growth",
    # invested-capital components (capital.py)
    "Cash & Short Term Investments": "Cash_and_Short_Term_Investments",
    "Minority Interest - Total": "Minority_Interest___Total",
    "Debt - Long-Term - Total": "Debt___Long_Term___Total",
    "Short-Term Debt & Current Portion of Long-Term Debt":
        "Short_Term_Debt_and_Current_Portion_of_Long_Term_Debt",
    "Shareholders' Equity - Attributable to Parent ShHold - Total":
        "Shareholders'_Equity___Attributable_to_Parent_ShHold___Total",
    "Funded Status (ASR)": "Funded_Status__ASR",
    "R&D Capital Base": "RandD_Capital_Base",
    "SG&A Capital Base": "SGandA_Capital_Base",
    "R&D Growth": "RandD_Growth",
    "SG&A Growth": "SGandA_Growth",
    # physical capital / reproduction / capex (overearning.py)
    "Property Plant & Equipment - Gross - Total": "Property_Plant_and_Equipment___Gross___Total",
    "PPE - Accumulated Depreciation & Impairment - Total":
        "PPE___Accumulated_Depreciation_and_Impairment___Total",
    "Gross Reproduction Cost": "Gross_Reproduction_Cost",
    "Fixed Assets - Purchased - Cash Flow (ASR)": "Fixed_Assets___Purchased___Cash_Flow__ASR",
    # history non-recurring items (history.py)
    "Impairment - Goodwill (ASR)": "Impairment___Goodwill__ASR",
    "Impairment - Intangibles excluding Goodwill (ASR)":
        "Impairment___Intangibles_excluding_Goodwill__ASR",
    "Restructuring Charges": "Restructuring_Charges",
    # descriptive
    "Company Name": "Company_Name",
    "GICS Industry Group Name": "GICS_Industry_Group_Name",
    "Country of Headquarters": "Country_of_Headquarters",
    "Market Capitalization": "Market_Capitalization",
    # richer columns the 30-year file adds (used by the next refinements)
    "Goodwill - Net": "Goodwill___Net",
    "Capital Expenditures - Net": "Capital_Expenditures___Net___Cash_Flow",
    "Year": "Year",
}


def _read_header(path):
    wb = pyxlsb.open_workbook(path)
    ws = wb.get_sheet(wb.sheets[0])
    rows = ws.rows()
    hdr_cells = next(rows)
    ncol = max((c.c for c in hdr_cells), default=-1) + 1
    hdr = [None] * ncol
    for c in hdr_cells:
        if c.c < ncol:
            hdr[c.c] = c.v
    return hdr


def _rows(path, ncol):
    """Yield padded value-rows (length ncol) for the data rows of a part."""
    wb = pyxlsb.open_workbook(path)
    ws = wb.get_sheet(wb.sheets[0])
    rows = ws.rows()
    next(rows)                          # skip header
    for rcells in rows:
        row = [None] * ncol
        for c in rcells:
            if c.c < ncol:
                row[c.c] = c.v
        yield row


def build_idx(hdr):
    """Column index keyed by BOTH the underscore name and the spaced alias, so
    code written against either naming finds its column."""
    idx = {h: i for i, h in enumerate(hdr) if h is not None}
    for spaced, under in ALIAS.items():
        if under in idx:
            idx[spaced] = idx[under]
    return idx


def load(paths, sheet=None):
    """Stitch the parts into {ticker: [rows oldest->newest]} and an aliased idx.

    `sheet` is ignored (each part may name its sheet differently; the first sheet
    is used). Rows are grouped by Instrument across all parts and sorted by Year."""
    hdr = _read_header(paths[0])
    ncol = len(hdr)
    idx = build_idx(hdr)
    ti, yi = idx.get("Instrument"), idx.get("Year")
    by = {}
    for p in paths:
        for r in _rows(p, ncol):
            t = r[ti] if (ti is not None and ti < len(r)) else None
            if t:
                by.setdefault(t, []).append(r)
    keyd = (lambda r: (yi is not None and yi < len(r) and r[yi] is not None,
                       r[yi] if (yi is not None and yi < len(r)) else None))
    for t in by:
        by[t].sort(key=keyd)
    return by, idx
