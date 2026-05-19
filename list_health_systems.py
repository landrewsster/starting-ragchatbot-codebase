#!/usr/bin/env python3
"""
list_health_systems.py

List all health systems represented among Round 3 providers, grouped by city,
ordered by provider count (most to least).

Because MailingList_Round3_20260519.xlsx has only name + address (no health
system column), health systems are looked up by matching each Round 3 provider's
address against mailed_providers_hs_reprocessed.xlsx.

Output:
  - Terminal: ranked list with cities and counts
  - health_systems_ranked.xlsx:
      ranked        — one row per health system, all cities, provider count
      by_hs_city    — one row per (health system, city) pair, provider count
      unmatched     — Round 3 rows whose address wasn't found in reprocessed file

Usage:
    python3 list_health_systems.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE  = BASE / "MailingList_Round3_20260519.xlsx"
HS_FILE      = BASE / "mailed_providers_hs_reprocessed.xlsx"
HS_SHEET     = "mailed_providers_hs"
OUTPUT_FILE  = BASE / "health_systems_ranked.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def zip5(s) -> str:
    return re.sub(r"\D", "", norm(s))[:5]

def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

def addr_key(addr, z):
    """Normalized address + zip5 used as join key."""
    return f"{norm(addr)}|{zip5(z)}"

# ── Load Round 3 mailing list ─────────────────────────────────────────────────
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
print(f"  {len(r3)} rows | columns: {list(r3.columns)}")

r3_addr = find_col(r3, ["primary_address_1", "Delivery Address", "address_line1"])
r3_zip  = find_col(r3, ["zip5", "ZIP+4", "zip", "Zip"])

if not r3_addr:
    raise SystemExit(f"ERROR: cannot find address column in Round 3 file.")
print(f"  addr={r3_addr}  zip={r3_zip}")

r3["_key"] = r3.apply(lambda row: addr_key(row[r3_addr], row[r3_zip] if r3_zip else ""), axis=1)

# ── Load health system source ─────────────────────────────────────────────────
print(f"\nLoading health system source: {HS_FILE.name}")
xl     = pd.ExcelFile(HS_FILE)
sheet  = HS_SHEET if HS_SHEET in xl.sheet_names else xl.sheet_names[0]
if sheet != HS_SHEET:
    print(f"  WARNING: '{HS_SHEET}' not found, using '{sheet}'")
hs_src = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(hs_src)} rows")

hs_col   = find_col(hs_src, ["health_system", "HealthSystem"])
city_col = find_col(hs_src, ["health_system_city", "hs_city"])
src_addr = find_col(hs_src, ["primary_address_1", "_addr1", "Delivery Address",
                              "address_line1", "Alternate 1 Address"])
src_zip  = find_col(hs_src, ["zip5", "ZIP+4", "_zip", "zip", "Zip"])

if not hs_col:
    raise SystemExit(f"ERROR: no health_system column in {HS_FILE.name}. Columns: {list(hs_src.columns)}")
print(f"  hs={hs_col}  city={city_col}  addr={src_addr}  zip={src_zip}")

# Build address key → (health_system, hs_city) lookup from reprocessed file.
# When multiple rows share an address, prefer the one with a non-empty health system.
addr_to_hs: dict[str, tuple[str, str]] = {}
for _, row in hs_src.iterrows():
    key = addr_key(row[src_addr] if src_addr else "", row[src_zip] if src_zip else "")
    if not key or key == "|":
        continue
    hs   = norm(row[hs_col])
    city = norm(row[city_col]) if city_col else ""
    if hs and key not in addr_to_hs:
        addr_to_hs[key] = (hs, city)

print(f"  Address → health system entries: {len(addr_to_hs)}")

# ── Match Round 3 rows to health systems ──────────────────────────────────────
print(f"\nMatching {len(r3)} Round 3 providers to health systems ...")

hs_values   = []
city_values = []
for _, row in r3.iterrows():
    result = addr_to_hs.get(row["_key"], ("", ""))
    hs_values.append(result[0])
    city_values.append(result[1])

r3["health_system"]      = hs_values
r3["health_system_city"] = city_values

matched   = (r3["health_system"] != "").sum()
unmatched = (r3["health_system"] == "").sum()
print(f"  Matched   : {matched}")
print(f"  Unmatched : {unmatched}")

# ── Build ranking from matched rows ───────────────────────────────────────────
hs_df = r3[r3["health_system"] != ""].copy()
hs_df["_hs"]   = hs_df["health_system"]
hs_df["_city"] = hs_df["health_system_city"]

by_hs_city = (
    hs_df.groupby(["_hs", "_city"], sort=False)
    .size()
    .reset_index(name="providers")
    .sort_values("providers", ascending=False)
    .rename(columns={"_hs": "health_system", "_city": "city"})
)

hs_totals = hs_df.groupby("_hs").size().reset_index(name="total_providers")

def collect_cities(sub):
    cities = (
        sub.groupby("_city").size()
        .reset_index(name="n")
        .sort_values("n", ascending=False)
    )
    parts = [
        f"{row['_city']} ({row['n']})" if row["_city"] else f"(unknown) ({row['n']})"
        for _, row in cities.iterrows()
    ]
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
ranked.index += 1

# ── Terminal output ───────────────────────────────────────────────────────────
print(f"\nAll health systems ({len(ranked)} total), ranked by provider count:\n")
for rank, row in ranked.iterrows():
    print(f"  {rank:3d}. {row['health_system']}")
    print(f"       {row['cities']}")
    print(f"       {row['total_providers']} providers")

# ── Write output ──────────────────────────────────────────────────────────────
unmatched_df = r3[r3["health_system"] == ""]

print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    ranked.to_excel(writer,       sheet_name="ranked",     index=True, index_label="rank")
    by_hs_city.to_excel(writer,   sheet_name="by_hs_city", index=False)
    unmatched_df.to_excel(writer, sheet_name="unmatched",  index=False)
    print(f"  ranked     : {len(ranked)} health systems")
    print(f"  by_hs_city : {len(by_hs_city)} health system + city combinations")
    print(f"  unmatched  : {len(unmatched_df)} Round 3 providers with no health system found")

print("\nDone.")
