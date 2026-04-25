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

STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$",
    re.IGNORECASE
)

def name_variants(last: str, first: str) -> set[str]:
    last  = STRIP_SUFFIXES.sub("", norm(last)).strip()
    first = STRIP_SUFFIXES.sub("", norm(first)).strip()
    first1 = first.split()[0] if first else ""
    cands = set()
    for f in {first, first1} - {""}:
        if last:
            cands.add(f"{f} {last}")
            cands.add(f"{last} {f}")
            cands.add(f"{last}, {f}")
        else:
            cands.add(f)
    if not cands and last:
        cands.add(last)
    return cands

def parse_full_name(full_name: str, first_hint: str = "") -> tuple[str, str, str]:
    """Parse 'Last, First Middle' or 'First Middle Last' → (last, first, middle)."""
    full_name = STRIP_SUFFIXES.sub("", full_name.strip()).strip()
    if not full_name:
        return "", "", ""
    if "," in full_name:
        parts = full_name.split(",", 1)
        last  = parts[0].strip()
        rest  = parts[1].strip().split()
        first  = rest[0] if rest else ""
        middle = " ".join(rest[1:]) if len(rest) > 1 else ""
        return last, first, middle
    words = full_name.split()
    if len(words) == 1:
        return words[0], "", ""
    if len(words) == 2:
        return words[1], words[0], ""
    # 3+ words: use first_hint to anchor the first name
    fh = norm(first_hint).split()[0] if first_hint else ""
    if fh and words[0].lower() == fh:
        first  = words[0]
        last   = words[-1]
        middle = " ".join(words[1:-1])
    else:
        first  = words[0]
        last   = words[-1]
        middle = " ".join(words[1:-1])
    return last, first, middle

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
            removed_rows.append({**row, "_removed_by": "duplicate_name_zip",
                                  "_dedup_key": name_key})
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

# ── Supplement: add 241498 mailing providers not already in gold ──────────────
MAIL_241498 = MAILING_BASE / "241498 0976 042026 v3.xlsx"
print(f"\nSupplement: checking {MAIL_241498.name} for providers not in gold ...")
try:
    mail_df = pd.read_excel(MAIL_241498, dtype=str).fillna("")
    print(f"  {len(mail_df)} rows in mailing file")
except FileNotFoundError:
    print(f"  WARNING: not found, skipping supplement")
    mail_df = None

if mail_df is not None:
    fn_col_m    = find_col(mail_df, ["FULL NAME", "Full Name"])
    fname_col_m = find_col(mail_df, ["First Name", "first_name", "FirstName"])
    addr_col_m  = find_col(mail_df, ["Delivery Address", "primary_address_1", "Alternate 1 Address"])
    city_col_m  = find_col(mail_df, ["City", "primary_city", "city"])
    zip_col_m   = find_col(mail_df, ["ZIP+4", "zip5", "zip", "Zip"])
    state_col_m = find_col(mail_df, ["State", "state", "St"])

    # Build name set from current gold reference for matching
    gold_name_set: set[str] = set()
    for _, row in gold.iterrows():
        l = norm(str(row.get(last_col,  ""))) if last_col  else ""
        f = norm(str(row.get(first_col, ""))) if first_col else ""
        for v in name_variants(l, f):
            gold_name_set.add(v)

    supplement_rows = []
    for _, row in mail_df.iterrows():
        fn    = norm(row.get(fn_col_m,    "")) if fn_col_m    else ""
        fname = norm(row.get(fname_col_m, "")) if fname_col_m else ""

        last, first, middle = parse_full_name(fn, fname)
        if not last:
            continue

        # Skip if already in gold
        if any(v in gold_name_set for v in name_variants(last, first)):
            continue

        addr  = str(row.get(addr_col_m,  "")).strip() if addr_col_m  else ""
        city  = str(row.get(city_col_m,  "")).strip() if city_col_m  else ""
        state = str(row.get(state_col_m, "")).strip() if state_col_m else ""
        z     = zip5(row.get(zip_col_m,  ""))          if zip_col_m  else ""

        new_row = {col: "" for col in gold.columns}
        new_row["last_name"]    = last
        new_row["first_name"]   = first
        new_row["middle_name"]  = middle
        new_row["_addr1"]       = addr
        new_row["_city"]        = city
        new_row["_zip"]         = z
        new_row["_source_file"] = "241498_mailing"
        new_row["_source_tab"]  = "1stmailing_notingold"
        supplement_rows.append(new_row)

    if supplement_rows:
        supp_df = pd.DataFrame(supplement_rows, columns=gold.columns)
        gold = pd.concat([gold, supp_df], ignore_index=True)
        print(f"  Added {len(supplement_rows)} providers from 241498 mailing")
    else:
        print(f"  All 241498 providers already in gold — nothing to add")

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
