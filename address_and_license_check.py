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

1. ADDRESS COMPARISON (matched.csv)
   For each provider in matched.csv, look them up in the master mailing
   list by name and compare addresses side-by-side.

2. LICENSE BOARD CHECK (not_in_master.csv)
   Name-match providers in not_in_master.csv against the MN Physician
   and PA list (MN State Licensing Board).

Output: address_and_license_check.xlsx
  address_same       — matched providers with the same address
  address_different  — matched providers with differing addresses
  in_license_board   — not-in-master providers found in MN license file
  not_in_license     — not-in-master providers not found in license file

Usage:
    python3 address_and_license_check.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
MATCHED_FILE = BASE / "matched.csv"
NOTMATCH_FILE= BASE / "not_in_master.csv"
MASTER_FILE  = BASE / "MasterMailingList_multiple_address_check_20260504.xlsx - all.csv"
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

# ── Load matched.csv ──────────────────────────────────────────────────────────
print(f"\nLoading matched file: {MATCHED_FILE.name}")
try:
    matched = pd.read_csv(MATCHED_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {MATCHED_FILE} not found")
print(f"  {len(matched)} rows | columns: {list(matched.columns)}")

m_last  = find_col(matched, ["Last_Name", "last_name", "LastName", "Last Name"]) or matched.columns[0]
m_first = find_col(matched, ["First_Name", "first_name", "FirstName", "First Name"]) or matched.columns[1]
m_mid   = find_col(matched, ["middle_name", "middle", "middle_initial", "manual_middle_initial"])
m_addr  = find_col(matched, ["primary_address_1", "Clinic_Address", "address", "Address", "address_line1"])
m_city  = find_col(matched, ["primary_city", "Clinic_City", "city", "City"])
m_zip   = find_col(matched, ["zip5", "Clinic_Zip", "zip", "Zip"])
print(f"  last={m_last!r}  first={m_first!r}  mid={m_mid!r}")
print(f"  addr={m_addr!r}  city={m_city!r}  zip={m_zip!r}")

# ── Load not_in_master.csv ────────────────────────────────────────────────────
print(f"\nLoading not-in-master file: {NOTMATCH_FILE.name}")
try:
    not_matched = pd.read_csv(NOTMATCH_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {NOTMATCH_FILE} not found")
print(f"  {len(not_matched)} rows | columns: {list(not_matched.columns)}")

nm_last  = find_col(not_matched, ["Last_Name", "last_name", "LastName", "Last Name"]) or not_matched.columns[0]
nm_first = find_col(not_matched, ["First_Name", "first_name", "FirstName", "First Name"]) or not_matched.columns[1]
print(f"  last={nm_last!r}  first={nm_first!r}")

# ── Load master mailing list (CSV) for address lookup ─────────────────────────
print(f"\nLoading master mailing list: {MASTER_FILE.name}")
try:
    master = pd.read_csv(MASTER_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {MASTER_FILE} not found")
print(f"  {len(master)} rows | columns: {list(master.columns)}")

ma_last  = find_col(master, ["last_name", "LastName", "Last Name"]) or master.columns[0]
ma_first = find_col(master, ["first_name", "FirstName", "First Name"]) or master.columns[1]
ma_addr  = find_col(master, ["primary_address_1", "Delivery Address", "address_line1"])
ma_city  = find_col(master, ["primary_city", "City", "city"])
ma_zip   = find_col(master, ["zip5", "zip", "Zip"])
print(f"  last={ma_last!r}  first={ma_first!r}")
print(f"  addr={ma_addr!r}  city={ma_city!r}  zip={ma_zip!r}")

# Build name → row index lookup from master
master_key_to_idx: dict[str, int] = {}
for idx, row in master.iterrows():
    for key in make_all_keys(row[ma_last], row[ma_first]):
        master_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(master_key_to_idx)}")

# ── Compare addresses for matched providers ───────────────────────────────────
print(f"\nComparing addresses for {len(matched)} matched providers ...")

addr_rows = []
for _, row in matched.iterrows():
    keys      = make_all_keys(row[m_last], row[m_first])
    found_idx = next((master_key_to_idx[k] for k in keys if k in master_key_to_idx), None)

    manual_name = f"{row[m_last]}, {row[m_first]}"
    if m_mid and row.get(m_mid, ""):
        manual_name += f" {row[m_mid]}"

    m_addr_raw  = row[m_addr]  if m_addr  else ""
    m_city_raw  = row[m_city]  if m_city  else ""
    m_zip_raw   = row[m_zip]   if m_zip   else ""

    if found_idx is not None:
        ma_row       = master.iloc[found_idx]
        master_name  = f"{ma_row[ma_last]}, {ma_row[ma_first]}"
        ma_addr_raw  = ma_row[ma_addr] if ma_addr else ""
        ma_city_raw  = ma_row[ma_city] if ma_city else ""
        ma_zip_raw   = ma_row[ma_zip]  if ma_zip  else ""

        m_norm  = f"{norm_addr(m_addr_raw)}|{norm(m_city_raw)}|{zip5(m_zip_raw)}"
        ma_norm = f"{norm_addr(ma_addr_raw)}|{norm(ma_city_raw)}|{zip5(ma_zip_raw)}"
        same    = (m_norm == ma_norm)

        addr_rows.append({
            "manual_name":      manual_name,
            "master_name":      master_name,
            "address_match":    "same" if same else "different",
            "manual_address":   m_addr_raw,
            "manual_city":      m_city_raw,
            "manual_zip":       m_zip_raw,
            "master_address":   ma_addr_raw,
            "master_city":      ma_city_raw,
            "master_zip":       ma_zip_raw,
        })
    else:
        addr_rows.append({
            "manual_name":      manual_name,
            "master_name":      "(not found in master)",
            "address_match":    "not_found",
            "manual_address":   m_addr_raw,
            "manual_city":      m_city_raw,
            "manual_zip":       m_zip_raw,
            "master_address":   "",
            "master_city":      "",
            "master_zip":       "",
        })

addr_df  = pd.DataFrame(addr_rows)
n_same   = (addr_df["address_match"] == "same").sum()
n_diff   = (addr_df["address_match"] == "different").sum()
n_nf     = (addr_df["address_match"] == "not_found").sum()
print(f"  Same address      : {n_same}")
print(f"  Different address : {n_diff}")
print(f"  Not found in master: {n_nf}")

# ── Load MN licensing board ───────────────────────────────────────────────────
print(f"\nLoading MN licensing board: {LICENSE_FILE.name}")
try:
    lic_xl = pd.ExcelFile(LICENSE_FILE)
    print(f"  Sheets: {lic_xl.sheet_names}")
    lic_df = lic_xl.parse(lic_xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {LICENSE_FILE} not found")
print(f"  {len(lic_df)} rows | columns: {list(lic_df.columns)}")

l_last  = find_col(lic_df, ["last_name", "LastName", "Last Name", "last", "LAST"]) or lic_df.columns[0]
l_first = find_col(lic_df, ["first_name", "FirstName", "First Name", "first", "FIRST"]) or lic_df.columns[1]
print(f"  last={l_last!r}  first={l_first!r}")

lic_key_to_idx: dict[str, int] = {}
for idx, row in lic_df.iterrows():
    for key in make_all_keys(row[l_last], row[l_first]):
        lic_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(lic_key_to_idx)}")

# ── Match not_in_master against licensing board ───────────────────────────────
print(f"\nChecking {len(not_matched)} not-in-master providers against license board ...")

in_license     = []
not_in_license = []

for _, row in not_matched.iterrows():
    keys      = make_all_keys(row[nm_last], row[nm_first])
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
print(f"  Found in license board : {len(in_lic_df)}")
print(f"  Not in license board   : {len(not_in_lic_df)}")

# ── Terminal summary ──────────────────────────────────────────────────────────
print(f"\nProviders with DIFFERENT addresses ({n_diff}):")
for _, row in addr_df[addr_df["address_match"] == "different"].iterrows():
    print(f"  {row['manual_name']}")
    print(f"    Manual : {row['manual_address']}, {row['manual_city']} {row['manual_zip']}")
    print(f"    Master : {row['master_address']}, {row['master_city']} {row['master_zip']}")

print(f"\nNot-in-master found in license board ({len(in_lic_df)}):")
for _, row in in_lic_df.iterrows():
    print(f"  {row[nm_last]}, {row[nm_first]}  →  {row['license_name_match']}")

# ── Write output ──────────────────────────────────────────────────────────────
addr_same = addr_df[addr_df["address_match"] == "same"].reset_index(drop=True)
addr_diff = addr_df[addr_df["address_match"].isin(["different", "not_found"])].reset_index(drop=True)

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
