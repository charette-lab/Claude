#!/usr/bin/env python3
"""validate_30y.py — pre-flight check for the 30-year extension data.

Run this against the new fundamentals file and the new price panel the moment they
land. It verifies the engine's required schema, reports the real date/instrument
coverage, and — most importantly — measures SURVIVORSHIP (whether delisted/failed
names are present, or only today's survivors). It does NOT run the engine; it tells
us whether the engine CAN be run honestly over 30 years.

  python3 pricing/validate_30y.py <fundamentals.xlsx|parquet> <price.parquet>[,<price2>...]
"""
import sys, warnings
import pandas as pd
warnings.filterwarnings("ignore")

# columns the full engine consumes per (instrument, period) — from build_decomposition / build_floor
REQUIRED = ["Instrument", "New Operating Income", "ROICm_total - 7 years", "average_C - 7 year",
            "Market Capitalization", "EV", "Debt - Long-Term - Total",
            "Short-Term Debt & Current Portion of Long-Term Debt", "Income Tax Rate - Instrument",
            "Sales", "GICS Industry Group Name", "Country of Headquarters"]
DATECOLS = ["Period_End_Date", "Date"]


def load_any(path):
    return pd.read_parquet(path) if path.endswith(".parquet") else pd.read_excel(path)


def check_fundamentals(path):
    print(f"\n=== FUNDAMENTALS: {path} ===")
    df = load_any(path)
    cols = set(df.columns)
    miss = [c for c in REQUIRED if c not in cols]
    dcol = next((c for c in DATECOLS if c in cols), None)
    print(f"  rows {len(df):,} | columns {len(df.columns)}")
    print(f"  date column: {dcol or 'MISSING (need Period_End_Date or Date)'}")
    print(f"  required columns MISSING ({len(miss)}): {miss if miss else 'none — schema OK'}")
    if dcol is None or "Instrument" not in cols:
        print("  -> cannot assess coverage without Instrument + a date column."); return
    df[dcol] = pd.to_datetime(df[dcol], errors="coerce"); df = df.dropna(subset=[dcol])
    df["Instrument"] = df["Instrument"].astype(str)
    yr = df[dcol].dt.year
    print(f"  span: {df[dcol].min():%Y-%m} -> {df[dcol].max():%Y-%m}  ({yr.nunique()} distinct years)")
    print(f"  instruments: {df['Instrument'].nunique():,}")
    by = df.groupby(yr)["Instrument"].nunique()
    print("  instruments per year (every ~5y):")
    for y in [c for c in by.index if c % 5 == 0 or c in (by.index.min(), by.index.max())]:
        print(f"    {y}: {by[y]}")
    # SURVIVORSHIP: names whose LAST observation is well before the panel end = delisted/dead
    end = df[dcol].max()
    last = df.groupby("Instrument")[dcol].max()
    dead = (last < end - pd.Timedelta(days=365*2))
    first = df.groupby("Instrument")[dcol].min()
    born_early = (first <= df[dcol].min() + pd.Timedelta(days=365*2))
    print(f"\n  SURVIVORSHIP CHECK (the #1 threat to a valid 30y backtest):")
    print(f"    names present near the START ({df[dcol].min():%Y}): {born_early.sum()}")
    print(f"    names that STOP >2y before the end (delisted/dead?): {dead.sum()} "
          f"({dead.mean()*100:.0f}% of universe)")
    if dead.sum() == 0:
        print("    ⚠ ZERO dead names -> this looks SURVIVOR-ONLY. A 30y backtest on today's")
        print("      survivors overstates return (never holds the names that went to zero).")
    else:
        print("    ✓ dead names present -> survivorship can be handled honestly.")


def check_prices(paths):
    print(f"\n=== PRICE PANEL: {paths} ===")
    frames = []
    for p in paths:
        try:
            frames.append(pd.read_parquet(p, columns=["Instrument", "Date", "Close Price"]))
        except Exception as e:
            print(f"  {p}: ERROR {e}"); return
    px = pd.concat(frames); px["Date"] = pd.to_datetime(px["Date"], errors="coerce")
    px["Instrument"] = px["Instrument"].astype(str)
    print(f"  rows {len(px):,} | span {px.Date.min():%Y-%m} -> {px.Date.max():%Y-%m} | instruments {px.Instrument.nunique():,}")
    return px


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); print("no args — paste the new file paths when they land."); sys.exit(0)
    check_fundamentals(sys.argv[1])
    if len(sys.argv) > 2:
        check_prices(sys.argv[2].split(","))
