#!/usr/bin/env python3
"""
check_mailing_coverage.py

Bidirectional comparison between a combined provider file and a mailing list.

Output tabs:
  matched / npionly / mnonly   — all provider rows, with in_mailing_list + match_method
  matched_not_mailed           — providers NOT on the mailing list (missing from mailing)
  npionly_not_mailed
  mnonly_not_mailed
  mailing_not_in_providers     — mailing list rows NOT found in any provider tab (missing from provider file)

Match priority (both directions):
  1. Full name  (first last / last first / last, first — middle names/suffixes stripped)
  2. Full address (addr1 + city + zip5)
  3. Name + city  (catches address format differences)

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
OUTPUT_FILE   = BASE / "Current Mailing Files" / "mailing_coverage_check_physicians.xlsx"

TABS = ["matched", "npionly", "mnonly"]

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def zip5(s) -> str:
    digits = re.sub(r"\D", "", norm(s))
    return digits[:5]

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
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
    """All normalized forms of a full name string."""
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

def provider_name_candidates(last: str, first: str) -> set[str]:
    """All name strings to try when matching a provider against a mailing full name."""
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

# ── Load mailing list ────────────────────────────────────────────────────────
print(f"\nLoading mailing list: {MAILING_FILE}")
try:
    mail_df = pd.read_excel(MAILING_FILE, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: mailing list not found at {MAILING_FILE}")

print(f"  {len(mail_df)} rows")

mail_fullname_col = find_col(mail_df, ["FULL NAME", "Full Name", "full_name"])
mail_first_col    = find_col(mail_df, ["First Name", "first_name", "FirstName"])
mail_addr_col     = find_col(mail_df, ["Delivery Address", "primary_address_1", "Alternate 1 Address"])
mail_city_col     = find_col(mail_df, ["City", "primary_city", "city"])
mail_zip_col      = find_col(mail_df, ["ZIP+4", "zip5", "zip", "Zip", "ZIP"])

print(f"  Columns       : {list(mail_df.columns)}")
print(f"  Full name col : {mail_fullname_col}")
print(f"  Address col   : {mail_addr_col}  City: {mail_city_col}  Zip: {mail_zip_col}")
if mail_fullname_col:
    print(f"  Sample names  : {mail_df[mail_fullname_col].dropna().head(5).tolist()}")

# ── Build mailing-list lookups (provider → mailing direction) ─────────────────
mail_fullname_set  = set()
mail_name_city_set = set()
mail_addr_idx      = {}

for i, row in mail_df.iterrows():
    city = norm(row[mail_city_col]) if mail_city_col else ""
    if mail_fullname_col:
        fn = norm(row[mail_fullname_col])
        for v in name_variants(fn):
            mail_fullname_set.add(v)
            if city:
                mail_name_city_set.add(f"{v}|{city}")
    if mail_addr_col:
        addr = norm(row[mail_addr_col])
        z    = zip5(row[mail_zip_col]) if mail_zip_col else ""
        if addr:
            mail_addr_idx.setdefault(f"{addr}|{city}|{z}", i)

print(f"\n  Name lookup      : {len(mail_fullname_set)} entries")
print(f"  Name+city lookup : {len(mail_name_city_set)} entries")
print(f"  Address lookup   : {len(mail_addr_idx)} entries")

# ── Provider column detection ─────────────────────────────────────────────────
def detect_provider_cols(df: pd.DataFrame, tab: str) -> dict:
    cols = {
        "last":    find_col(df, ["last_name", "LastName"]),
        "first":   find_col(df, ["first_name", "FirstName"]),
        "addr":    find_col(df, ["npi_primary_address_1", "primary_address_1"]),
        "city":    find_col(df, ["npi_primary_city", "primary_city"]),
        "zip":     find_col(df, ["npi_primary_zip", "zip5"]),
        "mn_addr": find_col(df, ["mn_address_1", "address_line1", "primary_address_1"]),
        "mn_city": find_col(df, ["mn_city", "primary_city"]),
        "mn_zip":  find_col(df, ["mn_zip", "zip5"]),
    }
    if cols["mn_addr"] == cols["addr"]:
        cols["mn_addr"] = None
    return cols

# ── Match: provider row → mailing list ───────────────────────────────────────
def provider_in_mailing(row, pc: dict) -> tuple[bool, str]:
    last  = norm(row.get(pc["last"],  "")) if pc["last"]  else ""
    first = norm(row.get(pc["first"], "")) if pc["first"] else ""
    cands = provider_name_candidates(last, first)

    if cands & mail_fullname_set:
        return True, "name"

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
            if addr and f"{addr}|{city}|{z}" in mail_addr_idx:
                return True, "address"

    if cands and prov_cities:
        for city in prov_cities:
            for c in cands:
                if f"{c}|{city}" in mail_name_city_set:
                    return True, "name_city"

    return False, "not_found"

# ── Process each provider tab ─────────────────────────────────────────────────
results: dict[str, pd.DataFrame] = {}

# Also build a combined provider lookup for the reverse check
prov_fullname_set  = set()
prov_name_city_set = set()
prov_addr_idx      = {}

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
        found, method = provider_in_mailing(row, pc)
        found_flags.append(found)
        match_methods.append(method)

        # Accumulate provider-side lookup for reverse check
        last  = norm(row.get(pc["last"],  "")) if pc["last"]  else ""
        first = norm(row.get(pc["first"], "")) if pc["first"] else ""
        cands = provider_name_candidates(last, first)
        for c in cands:
            prov_fullname_set.add(c)
        for addr_col, city_col, zip_col in [
            (pc["addr"],    pc["city"],    pc["zip"]),
            (pc["mn_addr"], pc["mn_city"], pc["mn_zip"]),
        ]:
            if addr_col:
                addr = norm(row.get(addr_col, ""))
                city = norm(row.get(city_col, "")) if city_col else ""
                z    = zip5(row.get(zip_col,  "")) if zip_col  else ""
                if city:
                    for c in cands:
                        prov_name_city_set.add(f"{c}|{city}")
                if addr:
                    prov_addr_idx.setdefault(f"{addr}|{city}|{z}", True)

    df["in_mailing_list"] = found_flags
    df["match_method"]    = match_methods

    n_found = sum(found_flags)
    n_miss  = len(found_flags) - n_found
    print(f"  In mailing list : {n_found}  |  NOT in mailing : {n_miss}")
    for method, cnt in pd.Series(match_methods).value_counts().items():
        print(f"    {method}: {cnt}")

    results[tab] = df

# ── Reverse check: mailing list → provider file ───────────────────────────────
print(f"\n{'─'*55}")
print("Reverse check: mailing list → provider file")

mail_found = []
mail_methods = []

for _, row in mail_df.iterrows():
    fn   = norm(row[mail_fullname_col]) if mail_fullname_col else ""
    city = norm(row[mail_city_col])     if mail_city_col     else ""
    addr = norm(row[mail_addr_col])     if mail_addr_col     else ""
    z    = zip5(row[mail_zip_col])      if mail_zip_col      else ""

    variants = name_variants(fn)

    found, method = False, "not_found"
    if any(v in prov_fullname_set for v in variants):
        found, method = True, "name"
    elif addr and f"{addr}|{city}|{z}" in prov_addr_idx:
        found, method = True, "address"
    elif city and any(f"{v}|{city}" in prov_name_city_set for v in variants):
        found, method = True, "name_city"

    mail_found.append(found)
    mail_methods.append(method)

mail_df["in_provider_file"] = mail_found
mail_df["match_method"]     = mail_methods

n_found = sum(mail_found)
n_miss  = len(mail_found) - n_found
print(f"  In provider file  : {n_found}  |  NOT in provider file : {n_miss}")
for method, cnt in pd.Series(mail_methods).value_counts().items():
    print(f"    {method}: {cnt}")

# ── Write output ─────────────────────────────────────────────────────────────
print(f"\n{'─'*55}")
print(f"Writing: {OUTPUT_FILE}")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for tab, df in results.items():
        df.to_excel(writer, sheet_name=tab, index=False)
        not_mailed = df[df["in_mailing_list"] == False].drop(
            columns=["in_mailing_list", "match_method"]
        )
        sheet = f"{tab}_not_mailed"
        not_mailed.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  {tab}: {len(df)} total | {len(not_mailed)} not on mailing list → '{sheet}'")

    # Mailing list rows not found in any provider tab
    not_in_providers = mail_df[mail_df["in_provider_file"] == False].drop(
        columns=["in_provider_file", "match_method"]
    )
    not_in_providers.to_excel(writer, sheet_name="mailing_not_in_providers", index=False)
    print(f"  mailing list: {len(mail_df)} total | {len(not_in_providers)} not in provider file → 'mailing_not_in_providers'")

print("\nDone.")
