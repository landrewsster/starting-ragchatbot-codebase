#!/usr/bin/env python3
# ── Self-update ───────────────────────────────────────────────────────────────
import os, sys, urllib.request
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "cross_reference_health_systems.py")
try:
    _new = urllib.request.urlopen(_URL, timeout=10).read()
    _old = open(__file__, "rb").read()
    if _new != _old:
        open(__file__, "wb").write(_new)
        print("Updated cross_reference_health_systems.py — restarting ...")
        os.execv(sys.executable, [sys.executable, __file__] + sys.argv[1:])
except Exception as _e:
    print(f"(self-update skipped: {_e})")
# ─────────────────────────────────────────────────────────────────────────────

"""
cross_reference_health_systems.py

Compare health systems in our Round 3 mailing list against the dup check list
to find which systems each list is missing.

Round 3 health systems are looked up by joining addresses against
mailed_providers_hs_reprocessed.xlsx. The dup check list supplies its
own System column.

Output: health_system_crossref.xlsx
  in_round3_not_in_dup  — systems in Round 3 but absent from dup check list
  in_dup_not_in_round3  — systems in dup check list but absent from Round 3
  matched               — systems present in both lists
  round3_unmatched      — Round 3 providers whose address had no health system

Usage:
    python3 cross_reference_health_systems.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE = BASE / "MailingList_Round3_20260519.xlsx"
HS_FILE     = BASE / "mailed_providers_hs_reprocessed.xlsx"
HS_SHEET    = "mailed_providers_hs"
DUP_FILE    = BASE / "dup check list.xlsx"
OUTPUT_FILE = BASE / "health_system_crossref.xlsx"

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
    return f"{norm(addr)}|{zip5(z)}"

# ── Load Round 3 ──────────────────────────────────────────────────────────────
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
print(f"  {len(r3)} rows")

r3_addr = find_col(r3, ["primary_address_1", "Delivery Address", "address_line1"])
r3_zip  = find_col(r3, ["zip5", "ZIP+4", "zip", "Zip"])
r3["_key"] = r3.apply(lambda row: addr_key(row[r3_addr], row[r3_zip] if r3_zip else ""), axis=1)

# ── Build address → health system lookup from reprocessed file ────────────────
print(f"\nLoading health system source: {HS_FILE.name}")
xl     = pd.ExcelFile(HS_FILE)
sheet  = HS_SHEET if HS_SHEET in xl.sheet_names else xl.sheet_names[0]
hs_src = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(hs_src)} rows")

hs_col   = find_col(hs_src, ["health_system", "HealthSystem"])
city_col = find_col(hs_src, ["health_system_city", "hs_city"])
src_addr = find_col(hs_src, ["primary_address_1", "_addr1", "Delivery Address", "address_line1"])
src_zip  = find_col(hs_src, ["zip5", "ZIP+4", "_zip", "zip", "Zip"])

addr_to_hs: dict[str, tuple[str, str]] = {}
for _, row in hs_src.iterrows():
    key = addr_key(row[src_addr] if src_addr else "", row[src_zip] if src_zip else "")
    if not key or key == "|":
        continue
    hs   = norm(row[hs_col]) if hs_col else ""
    city = norm(row[city_col]) if city_col else ""
    if hs and key not in addr_to_hs:
        addr_to_hs[key] = (hs, city)

# ── Assign health systems to Round 3 rows ────────────────────────────────────
r3["health_system"]      = r3["_key"].map(lambda k: addr_to_hs.get(k, ("", ""))[0])
r3["health_system_city"] = r3["_key"].map(lambda k: addr_to_hs.get(k, ("", ""))[1])

matched_r3   = r3[r3["health_system"] != ""]
unmatched_r3 = r3[r3["health_system"] == ""]
print(f"  Round 3 matched to health system : {len(matched_r3)}")
print(f"  Round 3 unmatched                : {len(unmatched_r3)}")

# ── Health system counts for Round 3 ─────────────────────────────────────────
r3_hs_counts = (
    matched_r3.groupby("health_system")
    .size()
    .reset_index(name="round3_providers")
    .sort_values("round3_providers", ascending=False)
)
r3_hs_set = set(r3_hs_counts["health_system"])

# ── Load dup check list ───────────────────────────────────────────────────────
print(f"\nLoading dup check list: {DUP_FILE.name}")
dup = pd.read_excel(DUP_FILE, dtype=str).fillna("")
print(f"  {len(dup)} rows")

dup_sys_col = find_col(dup, ["System", "system", "health_system"])
if not dup_sys_col:
    raise SystemExit(f"ERROR: no System column found. Columns: {list(dup.columns)}")

dup["_system"] = dup[dup_sys_col].apply(norm)
dup_hs_counts = (
    dup[dup["_system"] != ""]
    .groupby("_system")
    .size()
    .reset_index(name="dup_providers")
    .sort_values("dup_providers", ascending=False)
)
dup_hs_set = set(dup_hs_counts["_system"])

print(f"  Unique systems in dup check list : {len(dup_hs_set)}")
print(f"  Unique systems in Round 3        : {len(r3_hs_set)}")

# ── Compare ───────────────────────────────────────────────────────────────────
only_in_r3  = r3_hs_set - dup_hs_set
only_in_dup = dup_hs_set - r3_hs_set
in_both     = r3_hs_set & dup_hs_set

print(f"\n  In Round 3, missing from dup list : {len(only_in_r3)}")
print(f"  In dup list, missing from Round 3 : {len(only_in_dup)}")
print(f"  In both                           : {len(in_both)}")

# ── Build output dataframes ───────────────────────────────────────────────────
in_r3_not_dup = (
    r3_hs_counts[r3_hs_counts["health_system"].isin(only_in_r3)]
    .sort_values("round3_providers", ascending=False)
    .reset_index(drop=True)
)

in_dup_not_r3 = (
    dup_hs_counts[dup_hs_counts["_system"].isin(only_in_dup)]
    .rename(columns={"_system": "health_system"})
    .sort_values("dup_providers", ascending=False)
    .reset_index(drop=True)
)

matched_df = (
    r3_hs_counts[r3_hs_counts["health_system"].isin(in_both)]
    .merge(dup_hs_counts.rename(columns={"_system": "health_system"}),
           on="health_system", how="left")
    .sort_values("round3_providers", ascending=False)
    .reset_index(drop=True)
)

# ── Terminal summary ──────────────────────────────────────────────────────────
print(f"\nSystems in Round 3 NOT in dup check list ({len(in_r3_not_dup)}):")
for _, row in in_r3_not_dup.iterrows():
    print(f"  {row['health_system']}  ({row['round3_providers']} providers)")

print(f"\nSystems in dup check list NOT in Round 3 ({len(in_dup_not_r3)}):")
for _, row in in_dup_not_r3.iterrows():
    print(f"  {row['health_system']}  ({row['dup_providers']} providers)")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    in_r3_not_dup.to_excel(writer, sheet_name="in_round3_not_in_dup", index=False)
    in_dup_not_r3.to_excel(writer, sheet_name="in_dup_not_in_round3", index=False)
    matched_df.to_excel(   writer, sheet_name="matched",              index=False)
    unmatched_r3.to_excel( writer, sheet_name="round3_unmatched",     index=False)
    print(f"  in_round3_not_in_dup : {len(in_r3_not_dup)}")
    print(f"  in_dup_not_in_round3 : {len(in_dup_not_r3)}")
    print(f"  matched              : {len(matched_df)}")
    print(f"  round3_unmatched     : {len(unmatched_r3)}")

print("\nDone.")
