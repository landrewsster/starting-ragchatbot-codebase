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

Name-match individual providers between ManualSearch_05172026.xlsx and the
Round 3 mailing list to find:
  - Providers in the manual search file NOT in Round 3
  - Providers in Round 3 NOT in the manual search file
  - Providers matched in both

Matching is done by normalized (last, first) name key.  Suffix stripping
(MD, DO, NP, etc.) and whitespace normalization are applied to both sides
before comparison.

Output: provider_crossref.xlsx
  not_in_round3   — manual search providers with no Round 3 match
  not_in_dup      — Round 3 providers not in manual search file
  matched         — providers found in both lists

Usage:
    python3 cross_reference_providers.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE = BASE / "MailingList_Round3_20260519.xlsx"
DUP_FILE    = BASE / "ManualSearch_05172026.xlsx"
OUTPUT_FILE = BASE / "provider_crossref.xlsx"

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
    """Canonical matching key: 'lastname|firstname_initial'."""
    l = clean_name(last)
    f = clean_name(first)
    f1 = f.split()[0] if f else ""   # first token only (handles middle names)
    return f"{l}|{f1}"

def make_key_full(last: str, first: str) -> set[str]:
    """Return a set of keys to try: strict (last|first) and loose (last|initial)."""
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    keys = set()
    if l and f:
        keys.add(f"{l}|{f}")    # exact first name
    if l and f1:
        keys.add(f"{l}|{f1}")   # first initial / first token
    if l:
        keys.add(f"{l}|")       # last name only fallback
    return keys

# ── Load Round 3 ──────────────────────────────────────────────────────────────
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
print(f"  {len(r3)} rows")
print(f"  Columns: {list(r3.columns)}")

r3_last  = find_col(r3, ["last_name", "LastName", "Last Name", "LAST NAME"])
r3_first = find_col(r3, ["first_name", "FirstName", "First Name", "FIRST NAME"])

# Positional fallback if named columns not found
if not r3_last and len(r3.columns) >= 1:
    r3_last = r3.columns[0]
    print(f"  WARNING: using column A ({r3_last!r}) as last name")
if not r3_first and len(r3.columns) >= 2:
    r3_first = r3.columns[1]
    print(f"  WARNING: using column B ({r3_first!r}) as first name")

print(f"  last={r3_last!r}  first={r3_first!r}")

# Build lookup: key → list of original r3 row indices
r3_key_to_rows: dict[str, list[int]] = {}
for idx, row in r3.iterrows():
    for key in make_key_full(row[r3_last] if r3_last else "",
                              row[r3_first] if r3_first else ""):
        r3_key_to_rows.setdefault(key, []).append(idx)

print(f"  Unique name keys in Round 3: {len(r3_key_to_rows)}")

# ── Load manual search file ───────────────────────────────────────────────────
print(f"\nLoading manual search file: {DUP_FILE.name}")
xl_dup = pd.ExcelFile(DUP_FILE)
print(f"  Sheets: {xl_dup.sheet_names}")
dup = xl_dup.parse(xl_dup.sheet_names[0], dtype=str).fillna("")
print(f"  {len(dup)} rows")
print(f"  Columns: {list(dup.columns)}")

dup_last  = find_col(dup, ["Last_Name",  "last_name",  "LastName",  "Last Name",
                            "last",       "LAST",       "surname",   "Surname"])
dup_first = find_col(dup, ["First_Name", "first_name", "FirstName", "First Name",
                            "first",      "FIRST",      "given_name"])

# Positional fallback — print columns and let user know
if not dup_last:
    print(f"  WARNING: no last name column found — using column A ({dup.columns[0]!r})")
    dup_last = dup.columns[0]
if not dup_first:
    print(f"  WARNING: no first name column found — using column B ({dup.columns[1]!r})")
    dup_first = dup.columns[1]

print(f"  last={dup_last!r}  first={dup_first!r}")

# ── Deduplicate manual search file by name ────────────────────────────────────
dup["_norm_key"] = dup.apply(
    lambda row: make_key(row[dup_last], row[dup_first]), axis=1
)
before = len(dup)
dup = dup.drop_duplicates(subset=["_norm_key"]).reset_index(drop=True)
after = len(dup)
print(f"  Duplicates removed from manual search file: {before - after}  ({after} unique providers remain)")

# ── Match dup list against Round 3 ───────────────────────────────────────────
print(f"\nMatching {after} manual search providers against Round 3 ...")

dup_matched_flags = []   # True if found in Round 3
dup_r3_name       = []   # The matched Round 3 name (for review)

for _, row in dup.iterrows():
    last  = row[dup_last]
    first = row[dup_first]
    keys  = make_key_full(last, first)

    found_idx = None
    for key in keys:
        if key in r3_key_to_rows:
            found_idx = r3_key_to_rows[key][0]
            break

    if found_idx is not None:
        dup_matched_flags.append(True)
        r3_row = r3.iloc[found_idx]
        r3_name = f"{r3_row[r3_last] if r3_last else ''}, {r3_row[r3_first] if r3_first else ''}".strip(", ")
        dup_r3_name.append(r3_name)
    else:
        dup_matched_flags.append(False)
        dup_r3_name.append("")

dup["_in_round3"]     = dup_matched_flags
dup["_r3_name_match"] = dup_r3_name

n_matched     = sum(dup_matched_flags)
n_not_in_r3   = len(dup) - n_matched
print(f"  Dup list matched to Round 3 : {n_matched}")
print(f"  Dup list NOT in Round 3     : {n_not_in_r3}")

# ── Match Round 3 against dup list ───────────────────────────────────────────
# Build dup key set for quick lookup
dup_keys: set[str] = set()
for _, row in dup.iterrows():
    dup_keys.update(make_key_full(row[dup_last], row[dup_first]))

r3_matched_flags = []
for _, row in r3.iterrows():
    keys = make_key_full(row[r3_last] if r3_last else "",
                          row[r3_first] if r3_first else "")
    r3_matched_flags.append(bool(keys & dup_keys))

r3["_in_dup"] = r3_matched_flags
n_r3_not_dup  = sum(1 for f in r3_matched_flags if not f)
print(f"  Round 3 NOT in dup list     : {n_r3_not_dup}")
print(f"  Round 3 in dup list         : {sum(r3_matched_flags)}")

# ── Build output dataframes ───────────────────────────────────────────────────
not_in_r3 = (
    dup[~dup["_in_round3"]]
    .drop(columns=["_in_round3", "_r3_name_match", "_norm_key"])
    .reset_index(drop=True)
)

matched_dup = (
    dup[dup["_in_round3"]]
    .rename(columns={"_r3_name_match": "round3_name_match"})
    .drop(columns=["_in_round3", "_norm_key"])
    .reset_index(drop=True)
)

not_in_dup = (
    r3[~r3["_in_dup"]]
    .drop(columns=["_in_dup"])
    .reset_index(drop=True)
)

# ── Terminal summary for not_in_r3 ────────────────────────────────────────────
sys_col  = find_col(dup, ["System", "system", "health_system", "HealthSystem"])
city_col = find_col(dup, ["Clinic_City", "clinic_city", "City", "city"])

print(f"\nManual search providers NOT in Round 3 ({len(not_in_r3)}):")
for _, row in not_in_r3.iterrows():
    sys_val  = row.get(sys_col,  "") if sys_col  else ""
    city_val = row.get(city_col, "") if city_col else ""
    print(f"  {row[dup_last]}, {row[dup_first]}  |  {sys_val}  |  {city_val}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    not_in_r3.to_excel( writer, sheet_name="not_in_round3", index=False)
    matched_dup.to_excel(writer, sheet_name="matched",       index=False)
    not_in_dup.to_excel( writer, sheet_name="not_in_dup",    index=False)
    print(f"  not_in_round3 : {len(not_in_r3)} dup list providers")
    print(f"  matched       : {len(matched_dup)} providers in both")
    print(f"  not_in_dup    : {len(not_in_dup)} Round 3 providers not in dup list")

print("\nDone.")
