#!/usr/bin/env python3
"""
check_mailing_coverage.py

Compare a combined provider file (matched / npionly / mnonly tabs) against
a mailing list to identify which providers are already on the list.

Match priority per provider row:
  1. Full name match  — tries "first last", "last first", "last, first"
                        against mailing list FULL NAME column
  2. Address match    — addr1 + city + zip5 against Delivery Address + City + ZIP+4

Usage:
    python3 check_mailing_coverage.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE          = Path.home() / "Downloads" / "CRC MDH Project"
PROVIDER_FILE = BASE / "Current Mailing Files" / "physician_pa_nurse_20260422_combined.xlsx"
MAILING_FILE  = BASE / "Current Mailing Files" / "241498 0976 042026 v3.xlsx"
OUTPUT_FILE   = BASE / "Current Mailing Files" / "mailing_coverage_check.xlsx"

TABS = ["matched", "npionly", "mnonly"]

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def zip5(s) -> str:
    v = norm(s)
    digits = re.sub(r"\D", "", v)
    return digits[:5]

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

# ── Load mailing list ────────────────────────────────────────────────────────
print(f"\nLoading mailing list: {MAILING_FILE}")
try:
    mail_df = pd.read_excel(MAILING_FILE, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: mailing list not found at {MAILING_FILE}")

print(f"  {len(mail_df)} rows")

# Detect mailing list columns
mail_fullname_col = find_col(mail_df, ["FULL NAME", "Full Name", "full_name"])
mail_first_col    = find_col(mail_df, ["First Name", "first_name", "FirstName"])
mail_addr_col     = find_col(mail_df, ["Delivery Address", "primary_address_1", "Alternate 1 Address"])
mail_city_col     = find_col(mail_df, ["City", "primary_city", "city"])
mail_zip_col      = find_col(mail_df, ["ZIP+4", "zip5", "zip", "Zip", "ZIP"])

print(f"  All columns   : {list(mail_df.columns)}")
print(f"  Full name col : {mail_fullname_col}")
print(f"  First name col: {mail_first_col}")
print(f"  Address col   : {mail_addr_col}")
print(f"  City col      : {mail_city_col}")
print(f"  Zip col       : {mail_zip_col}")
if mail_fullname_col:
    samples = mail_df[mail_fullname_col].dropna().head(5).tolist()
    print(f"  Sample FULL NAME values: {samples}")
if mail_addr_col:
    samples = mail_df[mail_addr_col].dropna().head(3).tolist()
    print(f"  Sample address values  : {samples}")

STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$"
)

def fullname_variants(fn: str) -> list[str]:
    """Return all normalized forms of a full name to index."""
    fn = fn.strip()
    if not fn:
        return []
    variants = {fn}
    # Strip trailing suffix (Jr, MD, etc.)
    stripped = STRIP_SUFFIXES.sub("", fn).strip()
    if stripped and stripped != fn:
        variants.add(stripped)
    # If 3+ words, also add first-word + last-word (drops middle name/initial)
    words = stripped.split() if stripped else fn.split()
    if len(words) >= 3:
        fl = f"{words[0]} {words[-1]}"
        variants.add(fl)
        variants.add(f"{words[-1]} {words[0]}")   # last first order too
        variants.add(f"{words[-1]}, {words[0]}")
    return list(variants)

# Build lookup sets from mailing list
fullname_set  = set()   # normalized full name variants
name_city_set = set()   # "fullname_variant|city" for name+city fallback
addr_idx      = {}      # "addr|city|zip" → row index

for i, row in mail_df.iterrows():
    city = norm(row[mail_city_col]) if mail_city_col else ""

    # Full name — index all variants, and all variants paired with city
    if mail_fullname_col:
        fn = norm(row[mail_fullname_col])
        for v in fullname_variants(fn):
            fullname_set.add(v)
            if city:
                name_city_set.add(f"{v}|{city}")

    # Address (addr + city + zip)
    if mail_addr_col:
        addr = norm(row[mail_addr_col])
        z    = zip5(row[mail_zip_col]) if mail_zip_col else ""
        key  = f"{addr}|{city}|{z}"
        if addr:
            addr_idx.setdefault(key, i)

print(f"\n  Full name lookup  : {len(fullname_set)} entries")
print(f"  Name+city lookup  : {len(name_city_set)} entries")
print(f"  Address lookup    : {len(addr_idx)} entries")

# ── Provider column detection ─────────────────────────────────────────────────
def detect_provider_cols(df: pd.DataFrame, tab: str) -> dict:
    cols = {
        "last":     find_col(df, ["last_name", "LastName"]),
        "first":    find_col(df, ["first_name", "FirstName"]),
        "middle":   find_col(df, ["middle_name", "MiddleName"]),
        # NPI-side address (matched / npionly tabs)
        "addr":     find_col(df, ["npi_primary_address_1", "primary_address_1"]),
        "city":     find_col(df, ["npi_primary_city", "primary_city"]),
        "zip":      find_col(df, ["npi_primary_zip", "zip5"]),
        # MN-list address (matched / mnonly tabs)
        "mn_addr":  find_col(df, ["mn_address_1", "address_line1", "primary_address_1"]),
        "mn_city":  find_col(df, ["mn_city", "primary_city"]),
        "mn_zip":   find_col(df, ["mn_zip", "zip5"]),
    }
    # Avoid duplicating the same column across NPI/MN slots
    if cols["mn_addr"] == cols["addr"]:
        cols["mn_addr"] = None
    print(f"  [{tab}] last={cols['last']} first={cols['first']} "
          f"addr={cols['addr']} mn_addr={cols['mn_addr']}")
    return cols

# ── Match function ───────────────────────────────────────────────────────────
def match_row(row, pc: dict) -> tuple[bool, str]:
    """Return (found_in_mailing, match_method)."""

    last  = STRIP_SUFFIXES.sub("", norm(row.get(pc["last"],  ""))).strip() if pc["last"]  else ""
    first = STRIP_SUFFIXES.sub("", norm(row.get(pc["first"], ""))).strip() if pc["first"] else ""
    first1 = first.split()[0] if first else ""   # just first token of first name

    # 1. Name match — try all plausible orderings against mailing FULL NAME
    if last or first:
        candidates = set()
        for f in {first, first1} - {""}:
            if last:
                candidates.add(f"{f} {last}")
                candidates.add(f"{last} {f}")
                candidates.add(f"{last}, {f}")
            else:
                candidates.add(f)
        if not candidates and last:
            candidates.add(last)

        if candidates & fullname_set:
            return True, "name"

    # 2. Address match — try NPI-side address, then MN-list address
    prov_cities = set()
    for addr_col, city_col, zip_col in [
        (pc["addr"],    pc["city"],    pc["zip"]),
        (pc["mn_addr"], pc["mn_city"], pc["mn_zip"]),
    ]:
        if addr_col:
            addr = norm(row.get(addr_col, ""))
            city = norm(row.get(city_col, "")) if city_col else ""
            z    = zip5(row.get(zip_col,  "")) if zip_col  else ""
            if city:
                prov_cities.add(city)
            key = f"{addr}|{city}|{z}"
            if addr and key in addr_idx:
                return True, "address"

    # 3. Name + city fallback — catches address format differences
    if (last or first) and prov_cities:
        for city in prov_cities:
            for c in candidates:
                if f"{c}|{city}" in name_city_set:
                    return True, "name_city"

    return False, "not_found"

# ── Process each tab ─────────────────────────────────────────────────────────
results: dict[str, pd.DataFrame] = {}

for tab in TABS:
    print(f"\n{'─'*55}")
    print(f"Tab: {tab}")
    try:
        df = pd.read_excel(PROVIDER_FILE, sheet_name=tab, dtype=str)
    except Exception as e:
        print(f"  ERROR loading tab '{tab}': {e}")
        continue

    print(f"  {len(df)} rows")
    pc = detect_provider_cols(df, tab)

    found_flags   = []
    match_methods = []

    for _, row in df.iterrows():
        found, method = match_row(row, pc)
        found_flags.append(found)
        match_methods.append(method)

    df["in_mailing_list"] = found_flags
    df["match_method"]    = match_methods

    n_found = sum(found_flags)
    n_miss  = len(found_flags) - n_found
    print(f"\n  In mailing list : {n_found}")
    print(f"  NOT in mailing  : {n_miss}")

    method_counts = pd.Series(match_methods).value_counts()
    for method, cnt in method_counts.items():
        print(f"    {method}: {cnt}")

    results[tab] = df

# ── Write output ─────────────────────────────────────────────────────────────
print(f"\n{'─'*55}")
print(f"Writing: {OUTPUT_FILE}")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for tab, df in results.items():
        df.to_excel(writer, sheet_name=tab, index=False)

        not_in = df[df["in_mailing_list"] == False].copy()
        not_in.drop(columns=["in_mailing_list", "match_method"], inplace=True)
        sheet = f"{tab}_missing"
        not_in.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  {tab}: {len(df)} total | {len(not_in)} missing → '{sheet}' tab")

print("\nDone.")
