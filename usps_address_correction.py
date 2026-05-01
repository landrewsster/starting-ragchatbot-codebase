#!/usr/bin/env python3
"""
usps_address_correction.py

Apply USPS-corrected addresses from a presort corrections CSV to a mailing
list Excel file. Matches on normalized first + last name.

Input files (edit paths below):
  MAILING_FILE  — Excel mailing list to update
  USPS_CSV      — USPS presort corrections CSV (e.g. "241498 0976 041326.csv")

Output:
  A new Excel file with '_usps_corrected' appended to the name.
  Adds an 'address_updated' column ('Y' where address was changed).

Usage:
    python3 usps_address_correction.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project"
MAILING_FILE = BASE / "ProviderDataFiles" / "Mailing List 2 for Printing Services_columns.xlsx"
USPS_CSV     = BASE / "Current Mailing Files" / "241498 0976 041326.csv"
OUTPUT_FILE  = BASE / "ProviderDataFiles" / "Mailing List 2 for Printing Services_columns_usps_corrected.xlsx"

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

def load_csv(path):
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    sys.exit(f"ERROR: could not decode {path}")

# ── Load mailing list ────────────────────────────────────────────────────────
print(f"\nLoading mailing list: {MAILING_FILE}")
mail_df = pd.read_excel(MAILING_FILE, dtype=str)
print(f"  {len(mail_df)} rows | columns: {list(mail_df.columns)}")

# Detect name columns
mail_last_col  = find_col(mail_df, ["last_name",  "LastName",  "LAST_NAME"])
mail_first_col = find_col(mail_df, ["first_name", "FirstName", "First Name", "FIRST_NAME"])
mail_addr1_col = find_col(mail_df, ["primary_address_1", "address_line1", "Alternate 1 Address"])
mail_addr2_col = find_col(mail_df, ["primary_address_2", "address_line2", "Alternate 2 Address"])
mail_city_col  = find_col(mail_df, ["primary_city", "city", "City"])
mail_state_col = find_col(mail_df, ["state", "State", "St"])
mail_zip_col   = find_col(mail_df, ["zip5", "zip", "Zip", "ZIP+4"])

print(f"  Name cols  : last={mail_last_col} first={mail_first_col}")
print(f"  Addr cols  : addr1={mail_addr1_col} addr2={mail_addr2_col} city={mail_city_col}")

if not mail_last_col or not mail_first_col:
    sys.exit("ERROR: could not find last_name / first_name columns in mailing list")

# Build name key for each mailing row
mail_df["_name_key"] = (
    mail_df[mail_last_col].apply(norm) + " " + mail_df[mail_first_col].apply(norm)
)
name_to_idx = {k: i for i, k in mail_df["_name_key"].items() if k.strip()}

# ── Load USPS corrections ────────────────────────────────────────────────────
print(f"\nLoading USPS corrections: {USPS_CSV}")
usps_df = load_csv(USPS_CSV)
print(f"  {len(usps_df)} rows | columns: {list(usps_df.columns)}")

usps_last_col  = find_col(usps_df, ["last_name",  "LastName",  "Last Name"])
usps_first_col = find_col(usps_df, ["first_name", "FirstName", "First Name"])
usps_addr1_col = find_col(usps_df, ["Alternate 1 Address", "primary_address_1", "address_1"])
usps_addr2_col = find_col(usps_df, ["Alternate 2 Address", "primary_address_2", "address_2"])
usps_city_col  = find_col(usps_df, ["City", "primary_city", "city"])
usps_state_col = find_col(usps_df, ["St", "State", "state"])
usps_zip_col   = find_col(usps_df, ["ZIP+4", "zip5", "zip"])

print(f"  Name cols : last={usps_last_col} first={usps_first_col}")
print(f"  Addr cols : addr1={usps_addr1_col} city={usps_city_col}")

if not usps_last_col or not usps_first_col:
    sys.exit("ERROR: could not find last_name / first_name columns in USPS file")

# ── Apply corrections ────────────────────────────────────────────────────────
print(f"\nApplying corrections ...")
updated = 0
not_found = 0

mail_df["address_updated"] = ""

for _, row in usps_df.iterrows():
    key = norm(row[usps_last_col]) + " " + norm(row[usps_first_col])
    if key not in name_to_idx:
        not_found += 1
        continue

    idx = name_to_idx[key]
    changed = False

    if usps_addr1_col and mail_addr1_col:
        new_val = str(row[usps_addr1_col]).strip() if not pd.isna(row[usps_addr1_col]) else ""
        if new_val and norm(new_val) != norm(mail_df.at[idx, mail_addr1_col]):
            mail_df.at[idx, mail_addr1_col] = new_val
            changed = True

    if usps_addr2_col and mail_addr2_col:
        new_val = str(row[usps_addr2_col]).strip() if not pd.isna(row[usps_addr2_col]) else ""
        if new_val:
            mail_df.at[idx, mail_addr2_col] = new_val
            changed = True

    if usps_city_col and mail_city_col:
        new_val = str(row[usps_city_col]).strip() if not pd.isna(row[usps_city_col]) else ""
        if new_val and norm(new_val) != norm(mail_df.at[idx, mail_city_col]):
            mail_df.at[idx, mail_city_col] = new_val
            changed = True

    if usps_state_col and mail_state_col:
        new_val = str(row[usps_state_col]).strip() if not pd.isna(row[usps_state_col]) else ""
        if new_val:
            mail_df.at[idx, mail_state_col] = new_val

    if usps_zip_col and mail_zip_col:
        new_val = str(row[usps_zip_col]).strip() if not pd.isna(row[usps_zip_col]) else ""
        if new_val:
            mail_df.at[idx, mail_zip_col] = new_val

    if changed:
        mail_df.at[idx, "address_updated"] = "Y"
        updated += 1

print(f"  Addresses updated : {updated}")
print(f"  Not found in mailing list : {not_found}")

# ── Write output ─────────────────────────────────────────────────────────────
mail_df.drop(columns=["_name_key"], inplace=True)
print(f"\nWriting: {OUTPUT_FILE}")
mail_df.to_excel(OUTPUT_FILE, index=False)
print("Done.")
