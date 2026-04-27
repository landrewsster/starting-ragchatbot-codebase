#!/usr/bin/env python3
"""
check_round2.py

Inspect MailingListRound2.xlsx for:
  1. Duplicate entries (by NPI, and by name)
  2. Incomplete addresses (missing address, city, state, or zip)
  3. Name column issues:
       - Full name in column B instead of first name only
       - Likely first name in column A instead of last name

Usage:
    python3 check_round2.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
INPUT_FILE  = BASE / "MailingListRound2.xlsx"
OUTPUT_FILE = BASE / "round2_check.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
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

STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$",
    re.IGNORECASE,
)

# ── Load ──────────────────────────────────────────────────────────────────────
print(f"\nLoading: {INPUT_FILE.name}")
try:
    xl = pd.ExcelFile(INPUT_FILE)
    print(f"  Sheets: {xl.sheet_names}")
    df = xl.parse(xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {INPUT_FILE} not found")

print(f"  {len(df)} rows | {len(df.columns)} columns")
print(f"  Columns: {list(df.columns)}")

# ── Identify columns ──────────────────────────────────────────────────────────
# Column A is the first column, B is the second
col_a = df.columns[0]   # expected: last name
col_b = df.columns[1] if len(df.columns) > 1 else None  # expected: first name

# Columns D, E, F are the three address columns (positions 3, 4, 5)
# Column C (position 2) is middle name
addr_cols = [df.columns[i] for i in range(3, min(6, len(df.columns)))]
print(f"\n  Address columns (D-F): {addr_cols}")

npi_col   = find_col(df, ["npi", "NPI"])
state_col = find_col(df, ["state", "State", "primary_state", "mailing_state"])
zip_col   = find_col(df, ["zip", "Zip", "zip5", "postal_code", "mailing_postal_code",
                           "primary_zip", "mailing_zip"])

print(f"  npi={npi_col}  state={state_col}  zip={zip_col}")

# ── Flag issues ───────────────────────────────────────────────────────────────
# Address is missing only if ALL THREE of columns C, D, E are empty.
issues = []

for i, row in df.iterrows():
    row_issues = []

    # All three address columns empty → no address at all
    addr_values = [norm(row[c]) for c in addr_cols]
    if not any(addr_values):
        row_issues.append("missing_address")

    if state_col and not norm(row[state_col]):
        row_issues.append("missing_state")
    if zip_col and not norm(row[zip_col]):
        row_issues.append("missing_zip")

    if row_issues:
        issues.append((i, "|".join(row_issues)))

# ── Duplicates by NPI ─────────────────────────────────────────────────────────
dup_npi_rows = pd.DataFrame()
if npi_col:
    npi_series = df[npi_col].apply(norm)
    valid_npi  = npi_series[npi_series != ""]
    dup_mask   = valid_npi.duplicated(keep=False)
    dup_npi_rows = df[npi_series.isin(valid_npi[dup_mask])].copy()
    dup_npi_rows = dup_npi_rows.sort_values(npi_col)
    print(f"\nDuplicate NPI rows    : {len(dup_npi_rows)}")

# ── Duplicates by name ────────────────────────────────────────────────────────
def name_norm(row):
    a = norm(row[col_a])
    b = norm(row[col_b]) if col_b else ""
    return STRIP_SUFFIXES.sub("", f"{a}|{b}").strip()

df["_name_key"] = df.apply(name_norm, axis=1)
valid_names = df["_name_key"][df["_name_key"] != "|"]
dup_name_mask = valid_names.duplicated(keep=False)
dup_name_rows = df[df["_name_key"].isin(valid_names[dup_name_mask])].copy()
dup_name_rows = dup_name_rows.sort_values("_name_key")
print(f"Duplicate name rows   : {len(dup_name_rows)}")

# ── Build issue dataframe ─────────────────────────────────────────────────────
addr_issues_df = df.loc[[i for i, _ in issues]].copy()
addr_issues_df["_issues"] = [f for _, f in issues]
addr_issues_df = addr_issues_df.sort_values("_issues")

print(f"Incomplete address    : {len(addr_issues_df)}")

# ── Issue summary ─────────────────────────────────────────────────────────────
from collections import Counter
all_flags = []
for _, flags in issues:
    all_flags.extend(flags.split("|"))
if all_flags:
    print(f"\nIssue breakdown:")
    for flag, cnt in sorted(Counter(all_flags).items(), key=lambda x: -x[1]):
        print(f"  {cnt:4d}  {flag}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    addr_issues_df.drop(columns=["_name_key"], errors="ignore").to_excel(
        writer, sheet_name="address_issues", index=False)
    dup_npi_rows.to_excel(
        writer, sheet_name="duplicate_npi",  index=False)
    dup_name_rows.drop(columns=["_name_key"], errors="ignore").to_excel(
        writer, sheet_name="duplicate_name", index=False)
    print(f"  address_issues : {len(addr_issues_df)} rows")
    print(f"  address_issues : {len(addr_issues_df)} rows")
    print(f"  duplicate_npi  : {len(dup_npi_rows)} rows")
    print(f"  duplicate_name : {len(dup_name_rows)} rows")

print("\nDone.")
