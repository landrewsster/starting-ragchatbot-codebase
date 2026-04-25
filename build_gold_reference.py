#!/usr/bin/env python3
"""
build_gold_reference.py

Merge all sheets from physician_pa_nurse and nurse combined files into a
single deduplicated gold reference list of providers.

Deduplication priority:
  1. NPI (primary key — keeps the matched tab version over npionly/mnonly)
  2. last + first + middle + zip (for mnonly rows with no NPI)

Source tab priority when same provider appears in multiple tabs:
  matched > npionly > mnonly (most complete data first)

Output:
  gold_reference_providers.xlsx  — one sheet, one row per unique provider
  gold_reference_providers.csv   — same data as CSV

Usage:
    python3 build_gold_reference.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
MAILING_BASE  = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
PROVIDER_BASE = Path.home() / "Downloads" / "CRC MDH Project" / "ProviderDataFiles"
FILES = [
    PROVIDER_BASE / "physician_pa_nurse_20260422_combined.xlsx",
    PROVIDER_BASE / "nurse_20260422_combined.xlsx",
]
OUTPUT_XLSX = MAILING_BASE / "gold_reference_providers.xlsx"

# Tab priority — lower number = higher priority (kept when deduplicating)
TAB_PRIORITY = {"matched": 0, "npionly": 1, "mnonly": 2}

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
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

# ── Load all sheets from all files ───────────────────────────────────────────
print("\nLoading provider files ...")
all_frames: list[pd.DataFrame] = []

for filepath in FILES:
    print(f"\n  {filepath.name}")
    if not filepath.exists():
        print(f"    WARNING: not found, skipping")
        continue

    sheets = pd.read_excel(filepath, sheet_name=None, dtype=str)
    for sheet_name, df in sheets.items():
        df["_source_file"] = filepath.stem
        df["_source_tab"]  = sheet_name
        df["_tab_priority"] = TAB_PRIORITY.get(sheet_name, 99)
        print(f"    Sheet '{sheet_name}': {len(df)} rows, {len(df.columns)} columns")
        all_frames.append(df)

# ── Concatenate all frames ────────────────────────────────────────────────────
print(f"\nConcatenating all sheets ...")
combined = pd.concat(all_frames, ignore_index=True, sort=False)
combined = combined.fillna("")
print(f"  Total rows before dedup: {len(combined)}")

# ── Detect key columns ────────────────────────────────────────────────────────
npi_col    = find_col(combined, ["npi", "NPI"])
last_col   = find_col(combined, ["last_name", "LastName"])
first_col  = find_col(combined, ["first_name", "FirstName"])
middle_col = find_col(combined, ["middle_name", "MiddleName"])
zip_col    = find_col(combined, ["npi_primary_zip", "zip5", "mn_zip"])

print(f"  Key cols: npi={npi_col} last={last_col} first={first_col} "
      f"middle={middle_col} zip={zip_col}")

# ── Sort so higher-priority tabs come first ───────────────────────────────────
# Within same priority, prefer physician_pa_nurse over nurse
file_priority = {f.stem: i for i, f in enumerate(FILES)}
combined["_file_priority"] = combined["_source_file"].map(
    lambda x: file_priority.get(x, 99)
)
combined.sort_values(
    ["_tab_priority", "_file_priority"],
    inplace=True, ignore_index=True
)

# ── Deduplicate by NPI ────────────────────────────────────────────────────────
print(f"\nDeduplicating ...")
deduped_rows  = []
removed_rows  = []
seen_npi  = set()
seen_name = set()
removed_npi  = 0
removed_name = 0

for _, row in combined.iterrows():
    npi = norm(row[npi_col]) if npi_col else ""

    if npi and npi not in ("nan", ""):
        if npi in seen_npi:
            removed_npi += 1
            removed_rows.append({**row, "_removed_by": "duplicate_npi"})
            continue
        seen_npi.add(npi)
    else:
        # No NPI — deduplicate by last + first + middle + zip
        last   = norm(row[last_col])   if last_col   else ""
        first  = norm(row[first_col])  if first_col  else ""
        middle = norm(row[middle_col]) if middle_col else ""
        # Coalesce zip across all possible column names (varies by source tab)
        z = ""
        for zc in ["npi_primary_zip", "zip5", "mn_zip"]:
            if zc in row.index and zip5(row[zc]):
                z = zip5(row[zc])
                break
        name_key = f"{last}|{first}|{middle}|{z}"
        if name_key in seen_name or name_key == "|||":
            removed_name += 1
            removed_rows.append({**row, "_removed_by": "duplicate_name_zip"})
            continue
        seen_name.add(name_key)

    deduped_rows.append(row)

removed_df = pd.DataFrame(removed_rows)

gold = pd.DataFrame(deduped_rows, columns=combined.columns)

# Drop internal sort columns
gold.drop(columns=["_tab_priority", "_file_priority"], inplace=True)

# ── Normalize address columns ─────────────────────────────────────────────────
# Coalesce NPI-prefixed and plain address columns so every row has data
# in consistent column names regardless of which source tab it came from.
def coalesce(row, *cols):
    for c in cols:
        v = str(row.get(c, "")).strip()
        if v and v.lower() not in ("nan", ""):
            return v
    return ""

gold["_addr1"]     = gold.apply(lambda r: coalesce(r, "npi_primary_address_1", "primary_address_1", "mn_address_1"), axis=1)
gold["_addr2"]     = gold.apply(lambda r: coalesce(r, "npi_primary_address_2", "primary_address_2", "mn_address_2"), axis=1)
gold["_city"]      = gold.apply(lambda r: coalesce(r, "npi_primary_city",      "primary_city",      "mn_city"),      axis=1)
gold["_zip"]       = gold.apply(lambda r: coalesce(r, "npi_primary_zip",       "zip5",              "mn_zip"),       axis=1)
gold["_mail_addr"] = gold.apply(lambda r: coalesce(r, "npi_mailing_address_1", "mailing_address_1"),                 axis=1)
gold["_mail_city"] = gold.apply(lambda r: coalesce(r, "npi_mailing_city",      "mailing_city"),                     axis=1)
gold["_mail_zip"]  = gold.apply(lambda r: coalesce(r, "npi_mailing_zip",       "mailing_postal_code"),               axis=1)

print(f"  Rows after dedup : {len(gold)}")
print(f"  Rows removed     : {removed_npi + removed_name}  "
      f"(by NPI: {removed_npi}, by name+zip: {removed_name})")
print(f"\nSource breakdown:")
print(gold.groupby(["_source_file", "_source_tab"]).size().to_string())

addr_filled = (gold["_addr1"] != "").sum()
print(f"\n  Rows with primary address: {addr_filled} of {len(gold)}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_XLSX}")
with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    gold.to_excel(writer, sheet_name="gold_reference", index=False)
    if not removed_df.empty:
        removed_df.drop(columns=["_tab_priority", "_file_priority"],
                        errors="ignore").to_excel(writer, sheet_name="removed_duplicates", index=False)
    print(f"  gold_reference     : {len(gold)} rows")
    print(f"  removed_duplicates : {len(removed_df)} rows")

print(f"\nGold reference: {len(gold)} unique providers")
print("Done.")
