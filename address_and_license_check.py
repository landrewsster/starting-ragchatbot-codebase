#!/usr/bin/env python3
# ── Self-update ───────────────────────────────────────────────────────────────
import os, sys, urllib.request
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "address_and_license_check.py")
try:
    _new = urllib.request.urlopen(_URL, timeout=10).read()
    _old = open(__file__, "rb").read()
    if _new != _old:
        open(__file__, "wb").write(_new)
        print("Updated address_and_license_check.py — restarting ...")
        os.execv(sys.executable, [sys.executable, __file__] + sys.argv[1:])
except Exception as _e:
    print(f"(self-update skipped: {_e})")
# ─────────────────────────────────────────────────────────────────────────────

"""
address_and_license_check.py

Two tasks:

1. ADDRESS COMPARISON (matched providers)
   Re-runs the ManualSearch vs Round 3 name match, then for each matched
   pair compares the address from both files. Flags as same, different,
   or missing on one side.

2. LICENSE BOARD CHECK (not-in-Round-3 providers)
   Name-matches the unmatched manual search providers against the MN
   Physician and PA list (MN State Licensing Board) to see who appears
   in the state licensing file.

Output: address_and_license_check.xlsx
  address_same       — matched providers with the same address
  address_different  — matched providers with differing addresses
  in_license_board   — not-in-Round-3 providers found in MN license file
  not_in_license     — not-in-Round-3 providers not found in MN license file

Usage:
    python3 address_and_license_check.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE  = BASE / "MailingList_Round3_20260519.xlsx"
MANUAL_FILE  = BASE / "ManualSearch_05172026.xlsx"
LICENSE_FILE = BASE / "MN State Licensing Board" / "MN Physician and PA list March 2026.xlsx"
OUTPUT_FILE  = BASE / "address_and_license_check.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────
STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq|fnp|crna|crnp)\.?\s*$",
    re.IGNORECASE,
)
ADDR_ABBREVS = [
    (r"\bstreet\b", "st"), (r"\bavenue\b", "ave"), (r"\bboulevard\b", "blvd"),
    (r"\bdrive\b", "dr"), (r"\bcourt\b", "ct"), (r"\bcircle\b", "cir"),
    (r"\blane\b", "ln"), (r"\broad\b", "rd"), (r"\bplace\b", "pl"),
    (r"\bparkway\b", "pkwy"), (r"\bhighway\b", "hwy"),
    (r"\bsuite\b", "ste"), (r"\bfloor\b", "fl"), (r"\bbuilding\b", "bldg"),
    (r"#\s*(?=\d)", "ste "), (r"\bnorth\b", "n"), (r"\bsouth\b", "s"),
    (r"\beast\b", "e"), (r"\bwest\b", "w"),
    (r"[,\.]", " "),
]

def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def norm_addr(s) -> str:
    s = norm(s)
    for pattern, repl in ADDR_ABBREVS:
        s = re.sub(pattern, repl, s)
    return re.sub(r"\s+", " ", s).strip()

def zip5(s) -> str:
    return re.sub(r"\D", "", norm(s))[:5]

def clean_name(s) -> str:
    s = norm(s)
    s = STRIP_SUFFIXES.sub("", s).strip()
    return re.sub(r"[,.]", "", s).strip()

def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

def make_key(last: str, first: str) -> str:
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    return f"{l}|{f1}"

def lf_keys(last: str, first: str) -> set[str]:
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    keys = set()
    if l and f:
        keys.add(f"{l}|{f}")
    if l and f1 and f1 != f:
        keys.add(f"{l}|{f1}")
    return keys

def make_all_keys(col_a: str, col_b: str) -> set[str]:
    keys = set()
    keys |= lf_keys(col_a, col_b)
    b = clean_name(col_b)
    if " " in b:
        words = b.split()
        keys |= lf_keys(words[-1], words[0])
        keys |= lf_keys(words[-1], col_a)
    return keys

# ── Load Round 3 ──────────────────────────────────────────────────────────────
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
print(f"  {len(r3)} rows | columns: {list(r3.columns)}")

r3_last  = find_col(r3, ["last_name", "LastName", "Last Name"]) or r3.columns[0]
r3_first = find_col(r3, ["first_name", "FirstName", "First Name"]) or r3.columns[1]
r3_mid   = find_col(r3, ["middle_name", "middle", "MiddleName"]) or (r3.columns[2] if len(r3.columns) > 2 else None)
r3_addr  = find_col(r3, ["primary_address_1", "Delivery Address", "address_line1"])
r3_city  = find_col(r3, ["primary_city", "City", "city"])
r3_zip   = find_col(r3, ["zip5", "ZIP+4", "zip", "Zip"])
print(f"  last={r3_last!r}  first={r3_first!r}  mid={r3_mid!r}")
print(f"  addr={r3_addr!r}  city={r3_city!r}  zip={r3_zip!r}")

r3_key_to_idx: dict[str, int] = {}
for idx, row in r3.iterrows():
    for key in make_all_keys(row[r3_last], row[r3_first]):
        r3_key_to_idx.setdefault(key, idx)

# ── Load manual search file ───────────────────────────────────────────────────
print(f"\nLoading manual search file: {MANUAL_FILE.name}")
xl = pd.ExcelFile(MANUAL_FILE)
sheet = "ProviderList_COMBINED" if "ProviderList_COMBINED" in xl.sheet_names else xl.sheet_names[0]
print(f"  Using sheet: {sheet!r}")
manual = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(manual)} rows | columns: {list(manual.columns)}")

s_last  = find_col(manual, ["Last_Name",  "last_name",  "LastName",  "Last Name"])  or manual.columns[0]
s_first = find_col(manual, ["First_Name", "first_name", "FirstName", "First Name"]) or manual.columns[1]
s_mid   = find_col(manual, ["middle_name", "middle", "MiddleName", "middle_initial"]) or (manual.columns[1] if len(manual.columns) > 1 else None)
s_addr  = find_col(manual, ["primary_address_1", "address", "Address", "Clinic_Address", "address_line1"])
s_city  = find_col(manual, ["primary_city", "city", "City", "Clinic_City"])
s_zip   = find_col(manual, ["zip5", "zip", "Zip", "Clinic_Zip", "ZIP"])
print(f"  last={s_last!r}  first={s_first!r}  mid={s_mid!r}")
print(f"  addr={s_addr!r}  city={s_city!r}  zip={s_zip!r}")

# Deduplicate by name
manual["_norm_key"] = manual.apply(lambda row: make_key(row[s_last], row[s_first]), axis=1)
before = len(manual)
manual = manual.drop_duplicates(subset=["_norm_key"]).reset_index(drop=True)
print(f"  Duplicates removed: {before - len(manual)}  ({len(manual)} unique providers remain)")

# ── Match and compare addresses ───────────────────────────────────────────────
print(f"\nMatching and comparing addresses ...")

addr_rows = []
unmatched_rows = []

for _, row in manual.iterrows():
    keys      = make_all_keys(row[s_last], row[s_first])
    found_idx = next((r3_key_to_idx[k] for k in keys if k in r3_key_to_idx), None)

    manual_name = f"{row[s_last]}, {row[s_first]} {row[s_mid] if s_mid else ''}".strip()

    if found_idx is not None:
        r3_row   = r3.iloc[found_idx]
        r3_name  = f"{r3_row[r3_last]}, {r3_row[r3_first]} {r3_row[r3_mid] if r3_mid else ''}".strip()

        m_addr_raw  = row[s_addr]  if s_addr  else ""
        m_city_raw  = row[s_city]  if s_city  else ""
        m_zip_raw   = row[s_zip]   if s_zip   else ""
        r3_addr_raw = r3_row[r3_addr] if r3_addr else ""
        r3_city_raw = r3_row[r3_city] if r3_city else ""
        r3_zip_raw  = r3_row[r3_zip]  if r3_zip  else ""

        m_addr_norm  = f"{norm_addr(m_addr_raw)}|{norm(m_city_raw)}|{zip5(m_zip_raw)}"
        r3_addr_norm = f"{norm_addr(r3_addr_raw)}|{norm(r3_city_raw)}|{zip5(r3_zip_raw)}"
        same = (m_addr_norm == r3_addr_norm)

        addr_rows.append({
            "manual_name":      manual_name,
            "r3_name_match":    r3_name,
            "address_match":    "same" if same else "different",
            "manual_address":   m_addr_raw,
            "manual_city":      m_city_raw,
            "manual_zip":       m_zip_raw,
            "r3_address":       r3_addr_raw,
            "r3_city":          r3_city_raw,
            "r3_zip":           r3_zip_raw,
        })
    else:
        unmatched_rows.append(row)

addr_df      = pd.DataFrame(addr_rows)
unmatched_df = pd.DataFrame(unmatched_rows).drop(columns=["_norm_key"], errors="ignore")

n_same = (addr_df["address_match"] == "same").sum()     if len(addr_df) else 0
n_diff = (addr_df["address_match"] == "different").sum() if len(addr_df) else 0
print(f"  Matched providers       : {len(addr_df)}")
print(f"    Same address          : {n_same}")
print(f"    Different address     : {n_diff}")
print(f"  Not in Round 3         : {len(unmatched_rows)}")

# ── Load MN licensing board file ──────────────────────────────────────────────
print(f"\nLoading MN licensing board: {LICENSE_FILE.name}")
try:
    lic_xl    = pd.ExcelFile(LICENSE_FILE)
    print(f"  Sheets: {lic_xl.sheet_names}")
    lic_df    = lic_xl.parse(lic_xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {LICENSE_FILE} not found")
print(f"  {len(lic_df)} rows | columns: {list(lic_df.columns)}")

l_last  = find_col(lic_df, ["last_name", "LastName", "Last Name", "last", "LAST"])
l_first = find_col(lic_df, ["first_name", "FirstName", "First Name", "first", "FIRST"])
if not l_last:
    l_last  = lic_df.columns[0]
    print(f"  WARNING: using column A ({l_last!r}) as last name")
if not l_first:
    l_first = lic_df.columns[1]
    print(f"  WARNING: using column B ({l_first!r}) as first name")
print(f"  last={l_last!r}  first={l_first!r}")

lic_key_to_idx: dict[str, int] = {}
for idx, row in lic_df.iterrows():
    for key in make_all_keys(row[l_last], row[l_first]):
        lic_key_to_idx.setdefault(key, idx)

print(f"  Unique name keys in license file: {len(lic_key_to_idx)}")

# ── Match not-in-Round-3 against licensing board ──────────────────────────────
print(f"\nChecking {len(unmatched_rows)} not-in-Round-3 providers against license board ...")

in_license     = []
not_in_license = []

for _, row in unmatched_df.iterrows():
    keys      = make_all_keys(row[s_last], row[s_first] if s_first else "")
    found_idx = next((lic_key_to_idx[k] for k in keys if k in lic_key_to_idx), None)

    if found_idx is not None:
        lic_row  = lic_df.iloc[found_idx]
        lic_name = f"{lic_row[l_last]}, {lic_row[l_first]}".strip(", ")
        row2     = row.copy()
        row2["license_name_match"] = lic_name
        in_license.append(row2)
    else:
        not_in_license.append(row)

in_lic_df     = pd.DataFrame(in_license).reset_index(drop=True)
not_in_lic_df = pd.DataFrame(not_in_license).reset_index(drop=True)
print(f"  Found in license board  : {len(in_lic_df)}")
print(f"  Not in license board    : {len(not_in_lic_df)}")

# ── Terminal summary ──────────────────────────────────────────────────────────
if len(addr_df):
    print(f"\nProviders with DIFFERENT addresses ({n_diff}):")
    for _, row in addr_df[addr_df["address_match"] == "different"].iterrows():
        print(f"  {row['manual_name']}")
        print(f"    Manual : {row['manual_address']}, {row['manual_city']} {row['manual_zip']}")
        print(f"    Round3 : {row['r3_address']}, {row['r3_city']} {row['r3_zip']}")

print(f"\nNot-in-Round-3 found in license board ({len(in_lic_df)}):")
for _, row in in_lic_df.iterrows():
    print(f"  {row[s_last]}, {row[s_first]}  →  {row['license_name_match']}")

# ── Write output ──────────────────────────────────────────────────────────────
addr_same = addr_df[addr_df["address_match"] == "same"].reset_index(drop=True)   if len(addr_df) else pd.DataFrame()
addr_diff = addr_df[addr_df["address_match"] == "different"].reset_index(drop=True) if len(addr_df) else pd.DataFrame()

print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    addr_same.to_excel(    writer, sheet_name="address_same",      index=False)
    addr_diff.to_excel(    writer, sheet_name="address_different",  index=False)
    in_lic_df.to_excel(    writer, sheet_name="in_license_board",   index=False)
    not_in_lic_df.to_excel(writer, sheet_name="not_in_license",     index=False)
    print(f"  address_same      : {len(addr_same)}")
    print(f"  address_different : {len(addr_diff)}")
    print(f"  in_license_board  : {len(in_lic_df)}")
    print(f"  not_in_license    : {len(not_in_lic_df)}")

print("\nDone.")
