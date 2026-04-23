#!/usr/bin/env python3
"""
check_duplicates.py

Check for duplicate entries between two mailing list files.
Matches on: name, address, or name+city.

Usage:
    python3 check_duplicates.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE      = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
FILE_A    = BASE / "241498 0976 042026 v3.xlsx"          # original mailing list
FILE_B    = BASE / "MailingListAddition20260422.xlsx"    # additions
OUTPUT    = BASE / "duplicate_check.xlsx"

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
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$"
)

def name_variants(fn: str) -> list[str]:
    fn = fn.strip()
    if not fn:
        return []
    variants = {fn}
    stripped = STRIP_SUFFIXES.sub("", fn).strip()
    if stripped and stripped != fn:
        variants.add(stripped)
    words = (stripped or fn).split()
    if len(words) >= 3:
        variants.add(f"{words[0]} {words[-1]}")
        variants.add(f"{words[-1]} {words[0]}")
        variants.add(f"{words[-1]}, {words[0]}")
    return list(variants)

def last_first_variants(last: str, first: str) -> set[str]:
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

def detect_cols(df, label):
    cols = {
        "fullname": find_col(df, ["FULL NAME", "Full Name", "full_name"]),
        "last":     find_col(df, ["last_name", "LastName", "LAST_NAME"]),
        "first":    find_col(df, ["first_name", "FirstName", "First Name", "FIRST_NAME"]),
        "addr":     find_col(df, ["Delivery Address", "npi_primary_address_1",
                                   "primary_address_1", "address_line1", "Alternate 1 Address"]),
        "city":     find_col(df, ["City", "primary_city", "city", "npi_primary_city"]),
        "zip":      find_col(df, ["ZIP+4", "zip5", "npi_primary_zip", "zip", "Zip"]),
    }
    print(f"  [{label}] fullname={cols['fullname']} last={cols['last']} "
          f"first={cols['first']} addr={cols['addr']} city={cols['city']}")
    return cols

def build_lookups(df, cols, label):
    """Return (name_set, name_city_set, addr_idx) for matching against."""
    name_set      = set()
    name_city_set = set()
    addr_idx      = {}

    for i, row in df.iterrows():
        city = norm(row[cols["city"]]) if cols["city"] else ""

        # Build name variants from FULL NAME or last+first
        if cols["fullname"]:
            variants = name_variants(norm(row[cols["fullname"]]))
        elif cols["last"] or cols["first"]:
            variants = list(last_first_variants(
                row.get(cols["last"], "") if cols["last"] else "",
                row.get(cols["first"], "") if cols["first"] else "",
            ))
        else:
            variants = []

        for v in variants:
            name_set.add(v)
            if city:
                name_city_set.add(f"{v}|{city}")

        if cols["addr"]:
            addr = norm(row[cols["addr"]])
            z    = zip5(row[cols["zip"]]) if cols["zip"] else ""
            if addr:
                addr_idx.setdefault(f"{addr}|{city}|{z}", i)

    print(f"  [{label}] {len(name_set)} name entries, {len(addr_idx)} address entries")
    return name_set, name_city_set, addr_idx

def row_matches(row, cols, name_set, name_city_set, addr_idx) -> str | None:
    """Return match method string if row matches the lookup, else None.
    Address-only match is excluded — two providers at the same clinic
    address are not duplicates of each other."""
    city = norm(row[cols["city"]]) if cols["city"] else ""

    if cols["fullname"]:
        variants = name_variants(norm(row[cols["fullname"]]))
    elif cols["last"] or cols["first"]:
        variants = list(last_first_variants(
            row.get(cols["last"], "") if cols["last"] else "",
            row.get(cols["first"], "") if cols["first"] else "",
        ))
    else:
        variants = []

    if any(v in name_set for v in variants):
        return "name"

    if city and any(f"{v}|{city}" in name_city_set for v in variants):
        return "name_city"

    return None

# ── Load files ────────────────────────────────────────────────────────────────
print(f"\nLoading File A (original): {FILE_A.name}")
try:
    df_a = pd.read_excel(FILE_A, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {FILE_A} not found")
print(f"  {len(df_a)} rows | columns: {list(df_a.columns)}")
cols_a = detect_cols(df_a, "A")

print(f"\nLoading File B (additions): {FILE_B.name}")
try:
    df_b = pd.read_excel(FILE_B, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {FILE_B} not found")
print(f"  {len(df_b)} rows | columns: {list(df_b.columns)}")
cols_b = detect_cols(df_b, "B")

# ── Build lookups ─────────────────────────────────────────────────────────────
print(f"\nBuilding lookups from File A ...")
a_names, a_name_city, a_addr = build_lookups(df_a, cols_a, "A")

print(f"Building lookups from File B ...")
b_names, b_name_city, b_addr = build_lookups(df_b, cols_b, "B")

# ── Find duplicates ───────────────────────────────────────────────────────────
print(f"\nChecking File B rows against File A ...")
b_match_methods = []
for _, row in df_b.iterrows():
    method = row_matches(row, cols_b, a_names, a_name_city, a_addr)
    b_match_methods.append(method or "unique")

df_b["duplicate_of_A"] = [m != "unique" for m in b_match_methods]
df_b["match_method"]   = b_match_methods

n_dupes = sum(df_b["duplicate_of_A"])
print(f"  Duplicates found : {n_dupes} of {len(df_b)} additions")
print(f"  Unique additions : {len(df_b) - n_dupes}")

method_counts = pd.Series(b_match_methods).value_counts()
for m, cnt in method_counts.items():
    print(f"    {m}: {cnt}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT}")
dupes   = df_b[df_b["duplicate_of_A"] == True].drop(columns=["duplicate_of_A", "match_method"])
unique  = df_b[df_b["duplicate_of_A"] == False].drop(columns=["duplicate_of_A", "match_method"])

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    unique.to_excel(writer, sheet_name="unique_additions", index=False)
    dupes.to_excel(writer, sheet_name="duplicates", index=False)
    df_b.to_excel(writer, sheet_name="all_additions_flagged", index=False)
    print(f"  unique_additions    : {len(unique)} rows")
    print(f"  duplicates          : {len(dupes)} rows")
    print(f"  all_additions_flagged: {len(df_b)} rows (with duplicate_of_A + match_method columns)")

print("\nDone.")
