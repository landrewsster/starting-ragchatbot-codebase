#!/usr/bin/env python3
"""
check_mailing_coverage.py

Compare a combined provider file (matched / npionly / mnonly tabs) against
a mailing list to identify which providers are already on the list.

Match priority per provider row:
  1. NPI  (if NPI present in both files)
  2. Exact normalized last + first name
  3. Normalized address (addr1 + city + zip5)

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
MAILING_FILE  = BASE / "Current Mailing Files" / "241498 0976 041326 updated.xlsx"
OUTPUT_FILE   = BASE / "Current Mailing Files" / "mailing_coverage_check.xlsx"

TABS = ["matched", "npionly", "mnonly"]

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

def detect_cols(df: pd.DataFrame, label: str) -> dict:
    cols = {
        "npi":   find_col(df, ["npi", "NPI"]),
        "last":  find_col(df, ["last_name", "LastName", "LAST_NAME"]),
        "first": find_col(df, ["first_name", "FirstName", "FIRST_NAME"]),
        # address candidates — try NPI-prefixed first (wide pivot output), then plain
        "addr":  find_col(df, ["npi_primary_address_1", "primary_address_1",
                                "address_line1", "AddressLine1"]),
        "city":  find_col(df, ["npi_primary_city", "primary_city", "city", "City"]),
        "zip":   find_col(df, ["npi_primary_zip", "zip5", "zip", "Zip", "ZIP"]),
        # fallback MN-list address columns (present in matched / mnonly wide output)
        "mn_addr": find_col(df, ["address_line1", "AddressLine1"]),
        "mn_city": find_col(df, ["city", "City"]),
        "mn_zip":  find_col(df, ["zip", "Zip", "ZIP"]),
    }
    print(f"  [{label}] detected columns:")
    for k, v in cols.items():
        if v:
            print(f"    {k}: {v}")
    return cols

# ── Load mailing list ────────────────────────────────────────────────────────
print(f"\nLoading mailing list: {MAILING_FILE}")
try:
    mail_df = pd.read_excel(MAILING_FILE, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: mailing list not found at {MAILING_FILE}")

print(f"  {len(mail_df)} rows | columns: {list(mail_df.columns)}")
mc = detect_cols(mail_df, "mailing list")

# Build lookup indices from mailing list
npi_idx  = {}   # norm_npi  → mailing row index
name_idx = {}   # "last|first" → mailing row index
addr_idx = {}   # "addr1|city|zip" → mailing row index

for i, row in mail_df.iterrows():
    # NPI
    if mc["npi"]:
        v = norm(row[mc["npi"]])
        if v and v not in ("nan", ""):
            npi_idx.setdefault(v, i)

    # Name
    if mc["last"] and mc["first"]:
        key = norm(row[mc["last"]]) + "|" + norm(row[mc["first"]])
        if key != "|":
            name_idx.setdefault(key, i)

    # Address
    if mc["addr"]:
        addr = norm(row[mc["addr"]])
        city = norm(row[mc["city"]]) if mc["city"] else ""
        zip_ = norm(row[mc["zip"]])[:5] if mc["zip"] else ""
        key  = f"{addr}|{city}|{zip_}"
        if addr:
            addr_idx.setdefault(key, i)

print(f"\n  NPI lookup: {len(npi_idx)} entries")
print(f"  Name lookup: {len(name_idx)} entries")
print(f"  Address lookup: {len(addr_idx)} entries")

# ── Match function ───────────────────────────────────────────────────────────
def match_row(row, pc: dict) -> tuple[bool, str]:
    """Return (found_in_mailing, match_method)."""

    # 1. NPI
    if pc["npi"] and npi_idx:
        v = norm(row.get(pc["npi"], ""))
        if v and v in npi_idx:
            return True, "npi"

    # 2. Exact name
    if pc["last"] and pc["first"] and name_idx:
        key = norm(row.get(pc["last"], "")) + "|" + norm(row.get(pc["first"], ""))
        if key != "|" and key in name_idx:
            return True, "name_exact"

    # 3. Address (try NPI-side address, then MN-list address)
    for addr_col, city_col, zip_col in [
        (pc["addr"],    pc["city"],    pc["zip"]),
        (pc["mn_addr"], pc["mn_city"], pc["mn_zip"]),
    ]:
        if addr_col and addr_idx:
            addr = norm(row.get(addr_col, ""))
            city = norm(row.get(city_col, "")) if city_col else ""
            zip_ = norm(row.get(zip_col,  ""))[:5] if zip_col else ""
            key  = f"{addr}|{city}|{zip_}"
            if addr and key in addr_idx:
                return True, "address"

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

    print(f"  {len(df)} rows | columns: {list(df.columns)}")
    pc = detect_cols(df, tab)

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
        # Full tab with in_mailing_list and match_method columns added
        df.to_excel(writer, sheet_name=tab, index=False)

        # Filtered tab: only those NOT in the mailing list
        not_in = df[df["in_mailing_list"] == False].copy()
        not_in.drop(columns=["in_mailing_list", "match_method"], inplace=True)
        sheet = f"{tab}_missing"
        not_in.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  {tab}: {len(df)} total | {len(not_in)} missing from mailing list → '{sheet}' tab")

print("\nDone.")
