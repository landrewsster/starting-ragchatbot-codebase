#!/usr/bin/env python3
"""
list_health_systems.py

List all health systems in mailed_providers_hs_reprocessed.xlsx, grouped
by city, ordered by provider count (most to least).

Output:
  - Terminal: ranked list with cities and counts
  - health_systems_ranked.xlsx:
      ranked        — one row per health system, all cities, provider count
      by_hs_city    — one row per (health system, city) pair, provider count

Usage:
    python3 list_health_systems.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
INPUT_FILE  = BASE / "mailed_providers_hs_reprocessed.xlsx"
INPUT_SHEET = "mailed_providers_hs"
OUTPUT_FILE = BASE / "health_systems_ranked.xlsx"

HS_COL      = "health_system"
HS_CITY_COL = "health_system_city"

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).strip())

# ── Load ──────────────────────────────────────────────────────────────────────
print(f"\nLoading: {INPUT_FILE.name}  sheet='{INPUT_SHEET}'")
xl = pd.ExcelFile(INPUT_FILE)
sheet = INPUT_SHEET if INPUT_SHEET in xl.sheet_names else xl.sheet_names[0]
if sheet != INPUT_SHEET:
    print(f"  WARNING: '{INPUT_SHEET}' not found, using '{sheet}'")
df = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(df)} rows")

# Find columns (case-insensitive)
def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

hs_col   = find_col(df, [HS_COL, "health_system", "HealthSystem"])
city_col = find_col(df, [HS_CITY_COL, "health_system_city", "hs_city"])

if not hs_col:
    raise SystemExit(f"ERROR: could not find health_system column. Columns: {list(df.columns)}")

print(f"  health_system col : {hs_col}")
print(f"  city col          : {city_col}")

# ── Filter to rows with a health system ───────────────────────────────────────
hs_df = df[df[hs_col].apply(norm) != ""].copy()
hs_df["_hs"]   = hs_df[hs_col].apply(norm)
hs_df["_city"] = hs_df[city_col].apply(norm) if city_col else ""

print(f"  Rows with health system: {len(hs_df)}")
print(f"  Unique health systems  : {hs_df['_hs'].nunique()}")

# ── By (health_system, city) pair ─────────────────────────────────────────────
by_hs_city = (
    hs_df.groupby(["_hs", "_city"], sort=False)
    .size()
    .reset_index(name="providers")
    .sort_values("providers", ascending=False)
    .rename(columns={"_hs": "health_system", "_city": "city"})
)

# ── By health_system — aggregate cities, sum counts ───────────────────────────
hs_totals = hs_df.groupby("_hs").size().reset_index(name="total_providers")

def collect_cities(sub):
    cities = (
        sub.groupby("_city").size()
        .reset_index(name="n")
        .sort_values("n", ascending=False)
    )
    parts = [f"{row['_city']} ({row['n']})" if row["_city"] else f"(unknown) ({row['n']})"
             for _, row in cities.iterrows()]
    return ", ".join(parts)

hs_cities = hs_df.groupby("_hs").apply(collect_cities).reset_index()
hs_cities.columns = ["_hs", "cities"]

ranked = (
    hs_totals.merge(hs_cities, on="_hs")
    .sort_values("total_providers", ascending=False)
    .rename(columns={"_hs": "health_system"})
    [["health_system", "total_providers", "cities"]]
    .reset_index(drop=True)
)
ranked.index += 1  # 1-based rank

# ── Terminal output ───────────────────────────────────────────────────────────
print(f"\nAll health systems ({len(ranked)} total), ranked by provider count:\n")
for rank, row in ranked.iterrows():
    print(f"  {rank:3d}. {row['health_system']}")
    print(f"       {row['cities']}")
    print(f"       {row['total_providers']} providers")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    ranked.to_excel(writer, sheet_name="ranked", index=True, index_label="rank")
    by_hs_city.to_excel(writer, sheet_name="by_hs_city", index=False)
    print(f"  ranked     : {len(ranked)} health systems")
    print(f"  by_hs_city : {len(by_hs_city)} health system + city combinations")

print("\nDone.")
