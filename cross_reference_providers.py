#!/usr/bin/env python3
# ── Self-update ───────────────────────────────────────────────────────────────
import os, sys, urllib.request
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "cross_reference_providers.py")
try:
    _new = urllib.request.urlopen(_URL, timeout=10).read()
    _old = open(__file__, "rb").read()
    if _new != _old:
        open(__file__, "wb").write(_new)
        print("Updated cross_reference_providers.py — restarting ...")
        os.execv(sys.executable, [sys.executable, __file__] + sys.argv[1:])
except Exception as _e:
    print(f"(self-update skipped: {_e})")
# ─────────────────────────────────────────────────────────────────────────────

"""
cross_reference_providers.py

Name-match providers between ManualSearch_05172026.xlsx and the master
mailing list (MasterMailingList_multiple_address_check_20260504 - all.csv)
to find providers in the manual search file that are NOT in the master list.

Matching is done by normalized (last, first) name key with suffix stripping
(MD, DO, NP, etc.) and first-token first-name matching.

Output: provider_crossref.xlsx
  not_in_master   — manual search providers with no master list match
  matched         — providers found in both lists

Usage:
    python3 cross_reference_providers.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
MASTER_FILE  = BASE / "MailingList_Round3_20260519.xlsx"
MANUAL_FILE  = BASE / "ManualSearch_05172026.xlsx"
OUTPUT_FILE  = BASE / "provider_round3_crossref.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────
STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq|fnp|crna|crnp)\.?\s*$",
    re.IGNORECASE,
)

def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def clean_name(s) -> str:
    s = norm(s)
    s = STRIP_SUFFIXES.sub("", s).strip()
    s = re.sub(r"[,.]", "", s).strip()
    return s

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
    """Keys from explicit last + first values. Requires both to be non-empty."""
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
    """
    If col B has no space: col A = last name, col B = first name.
    If col B has a space:  col B = full name (first [middle] last),
                           col A is a repeat of the first name — ignore it.
    For 3+ word full names, also tries last two words as a compound last name
    to handle names like 'Kaeley Whiting Allen' where last name = 'Whiting Allen'.
    """
    b = clean_name(col_b)
    if " " in b:
        words = b.split()
        first = words[0]
        keys  = lf_keys(words[-1], first)           # single last word
        if len(words) >= 3:
            compound = " ".join(words[-2:])
            keys |= lf_keys(compound, first)         # last two words as last name
        return keys
    else:
        return lf_keys(col_a, col_b)                 # standard: col A=last, col B=first

# ── Load master mailing list (CSV) ────────────────────────────────────────────
print(f"\nLoading Round 3 mailing list: {MASTER_FILE.name}")
try:
    master = pd.read_excel(MASTER_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {MASTER_FILE} not found")
print(f"  {len(master)} rows")
print(f"  Columns: {list(master.columns)}")

m_last  = find_col(master, ["last_name",  "LastName",  "Last Name",  "LAST NAME",  "last"])
m_first = find_col(master, ["first_name", "FirstName", "First Name", "FIRST NAME", "first"])

if not m_last and len(master.columns) >= 1:
    m_last = master.columns[0]
    print(f"  WARNING: using column A ({m_last!r}) as last name")
if not m_first and len(master.columns) >= 2:
    m_first = master.columns[1]
    print(f"  WARNING: using column B ({m_first!r}) as first name")

m_mid = find_col(master, ["middle_name", "middle", "MiddleName", "Middle Name", "middle_initial"])
if not m_mid and len(master.columns) >= 3:
    m_mid = master.columns[2]
    print(f"  WARNING: using column C ({m_mid!r}) as middle name")
print(f"  last={m_last!r}  first={m_first!r}  middle={m_mid!r}")

# Build lookup: key → row index in master
master_key_to_idx: dict[str, int] = {}
for idx, row in master.iterrows():
    col_a = row[m_last]  if m_last  else ""
    col_b = row[m_first] if m_first else ""
    for key in make_all_keys(col_a, col_b):
        master_key_to_idx.setdefault(key, idx)

print(f"  Unique name keys in master list: {len(master_key_to_idx)}")

# ── Load manual search file ───────────────────────────────────────────────────
print(f"\nLoading manual search file: {MANUAL_FILE.name}")
xl = pd.ExcelFile(MANUAL_FILE)
print(f"  Sheets: {xl.sheet_names}")
sheet = "ProviderList_COMBINED" if "ProviderList_COMBINED" in xl.sheet_names else xl.sheet_names[0]
print(f"  Using sheet: {sheet!r}")
manual = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(manual)} rows")
print(f"  Columns: {list(manual.columns)}")

s_last  = find_col(manual, ["Last_Name",  "last_name",  "LastName",  "Last Name",  "last",  "LAST",  "surname"])
s_first = find_col(manual, ["First_Name", "first_name", "FirstName", "First Name", "first", "FIRST", "given_name"])

if not s_last:
    print(f"  WARNING: no last name column found — using column A ({manual.columns[0]!r})")
    s_last = manual.columns[0]
if not s_first:
    print(f"  WARNING: no first name column found — using column B ({manual.columns[1]!r})")
    s_first = manual.columns[1]

s_mid = find_col(manual, ["middle_name", "middle", "MiddleName", "Middle Name", "middle_initial"])
if not s_mid and len(manual.columns) >= 2:
    s_mid = manual.columns[1]   # col B = middle initial in manual search
    print(f"  Using column B ({s_mid!r}) as middle initial")
print(f"  last={s_last!r}  first={s_first!r}  middle={s_mid!r}")

# ── Deduplicate manual search file by name ────────────────────────────────────
manual["_norm_key"] = manual.apply(
    lambda row: make_key(row[s_last], row[s_first] if s_first else ""), axis=1
)
before = len(manual)
dups_removed = manual[manual.duplicated(subset=["_norm_key"], keep="first")].copy()
manual = manual.drop_duplicates(subset=["_norm_key"]).reset_index(drop=True)
after  = len(manual)
print(f"  Duplicates removed: {before - after}  ({after} unique providers remain)")

# ── Match manual search providers against master list ─────────────────────────
print(f"\nMatching {after} manual search providers against master list ...")

matched_flags  = []
matched_names  = []

for _, row in manual.iterrows():
    keys      = make_all_keys(row[s_last], row[s_first] if s_first else "")
    found_idx = next((master_key_to_idx[k] for k in keys if k in master_key_to_idx), None)

    if found_idx is not None:
        matched_flags.append(True)
        m_row   = master.iloc[found_idx]
        m_l     = m_row[m_last]  if m_last  else ""
        m_f     = m_row[m_first] if m_first else ""
        m_m     = m_row[m_mid]   if m_mid   else ""
        m_name  = f"{m_l}, {m_f} {m_m}".strip().strip(",").strip()
        matched_names.append(m_name)
    else:
        matched_flags.append(False)
        matched_names.append("")

manual["_in_master"]         = matched_flags
manual["_master_name_match"] = matched_names
manual["_manual_middle"]     = manual[s_mid].apply(norm) if s_mid else ""

n_matched   = sum(matched_flags)
n_not_found = after - n_matched
print(f"  Matched to master list : {n_matched}")
print(f"  NOT in master list     : {n_not_found}")

# ── Build output dataframes ───────────────────────────────────────────────────
not_in_master = (
    manual[~manual["_in_master"]]
    .drop(columns=["_in_master", "_master_name_match", "_manual_middle", "_norm_key"])
    .reset_index(drop=True)
)

matched_df = (
    manual[manual["_in_master"]]
    .rename(columns={
        "_master_name_match": "master_name_match",
        "_manual_middle":     "manual_middle_initial",
    })
    .drop(columns=["_in_master", "_norm_key"])
    .reset_index(drop=True)
)

# ── Terminal summary ──────────────────────────────────────────────────────────
sys_col  = find_col(manual, ["System", "system", "health_system", "HealthSystem"])
city_col = find_col(manual, ["Clinic_City", "clinic_city", "City", "city"])

print(f"\nManual search providers NOT in master list ({len(not_in_master)}):")
for _, row in not_in_master.iterrows():
    sys_val  = row.get(sys_col,  "") if sys_col  else ""
    city_val = row.get(city_col, "") if city_col else ""
    print(f"  {row[s_last]}, {row[s_first]}  |  {sys_val}  |  {city_val}")

# ── Write output ──────────────────────────────────────────────────────────────
dups_out = dups_removed.drop(columns=["_norm_key"], errors="ignore").reset_index(drop=True)

print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    not_in_master.to_excel(writer, sheet_name="not_in_master",     index=False)
    matched_df.to_excel(   writer, sheet_name="matched",            index=False)
    dups_out.to_excel(     writer, sheet_name="removed_duplicates", index=False)
    print(f"  not_in_master      : {len(not_in_master)} providers")
    print(f"  matched            : {len(matched_df)} providers")
    print(f"  removed_duplicates : {len(dups_out)} providers")

print("\nDone.")
