#!/usr/bin/env python3
"""
compare_gold_vs_reprocessed.py

Compare gold_reference_providers (gold_reference sheet) against
mailed_providers_hs_reprocessed (mailed_providers_hs sheet) to identify
which gold reference providers were not included in the mailed list.

Output: compare_gold_vs_reprocessed.xlsx
  not_in_mailed   — gold providers absent from the mailed list (~670 expected)
  in_mailed       — gold providers found in the mailed list, with health system
  summary         — counts broken down by source tab

Usage:
    python3 compare_gold_vs_reprocessed.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE         = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
GOLD_FILE    = BASE / "gold_reference_providers.xlsx"
GOLD_SHEET   = "gold_reference"
MAILED_FILE  = BASE / "mailed_providers_hs_reprocessed.xlsx"
MAILED_SHEET = "mailed_providers_hs"
OUTPUT_FILE  = BASE / "compare_gold_vs_reprocessed.xlsx"

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

def clean_name(s) -> str:
    s = STRIP_SUFFIXES.sub("", norm(s)).strip()
    return re.sub(r"[,.]", "", s).strip()

def name_key(last: str, first: str) -> set[str]:
    last  = clean_name(last)
    first = clean_name(first)
    first1 = first.split()[0] if first else ""
    keys = set()
    for f in {first, first1} - {""}:
        if last:
            keys.add(f"{last}|{f}")
            keys.add(f"{f}|{last}")
        else:
            keys.add(f)
    if not keys and last:
        keys.add(last)
    return keys

def fullname_keys(fn: str) -> set[str]:
    """Keys from a 'First Last' or 'Last, First' full-name string."""
    fn = clean_name(fn)
    if not fn:
        return set()
    keys = set()
    words = fn.split()
    if len(words) >= 2:
        last  = words[-1]
        first = words[0]
        first1 = first.split()[0]
        for f in {first, first1} - {""}:
            keys.add(f"{last}|{f}")
            keys.add(f"{f}|{last}")
    elif fn:
        keys.add(fn)
    return keys

# ── Load gold reference ────────────────────────────────────────────────────────
print(f"\nLoading gold reference: {GOLD_FILE.name}  sheet='{GOLD_SHEET}'")
xl = pd.ExcelFile(GOLD_FILE)
sheet = GOLD_SHEET if GOLD_SHEET in xl.sheet_names else xl.sheet_names[0]
if sheet != GOLD_SHEET:
    print(f"  WARNING: '{GOLD_SHEET}' not found, using '{sheet}'")
gold = xl.parse(sheet, dtype=str).fillna("")
print(f"  {len(gold)} rows | columns: {list(gold.columns)}")

g_npi   = find_col(gold, ["npi", "NPI"])
g_last  = find_col(gold, ["last_name", "LastName", "Last Name"])
g_first = find_col(gold, ["first_name", "FirstName", "First Name"])
g_src   = find_col(gold, ["_source_tab", "source_tab", "source"])
print(f"  npi={g_npi}  last={g_last}  first={g_first}  source={g_src}")

# ── Load mailed providers ──────────────────────────────────────────────────────
print(f"\nLoading mailed providers: {MAILED_FILE.name}  sheet='{MAILED_SHEET}'")
mailed = pd.read_excel(MAILED_FILE, sheet_name=MAILED_SHEET, dtype=str).fillna("")
print(f"  {len(mailed)} rows | columns: {list(mailed.columns[:15])} ...")

m_npi      = find_col(mailed, ["npi", "NPI"])
m_last     = find_col(mailed, ["last_name", "LastName", "Last Name", "_check_last"])
m_first    = find_col(mailed, ["first_name", "FirstName", "First Name", "_check_first"])
m_fullname = find_col(mailed, ["_check_fullname", "FULL NAME", "Full Name", "full_name"])
m_hs       = find_col(mailed, ["health_system"])
m_city     = find_col(mailed, ["health_system_city"])
print(f"  npi={m_npi}  last={m_last}  first={m_first}  fullname={m_fullname}  hs={m_hs}")

# ── Build lookup sets from mailed list ────────────────────────────────────────
mailed_npi_set:  set[str] = set()
mailed_name_set: set[str] = set()

# Also build NPI → health_system/city for the "in_mailed" join
npi_to_hs:   dict[str, str] = {}
npi_to_city: dict[str, str] = {}

for _, row in mailed.iterrows():
    npi  = norm(row[m_npi])      if m_npi      else ""
    last = norm(row[m_last])     if m_last     else ""
    frst = norm(row[m_first])    if m_first    else ""
    fn   = norm(row[m_fullname]) if m_fullname else ""
    hs   = str(row[m_hs]).strip()   if m_hs   else ""
    city = str(row[m_city]).strip() if m_city else ""

    if npi:
        mailed_npi_set.add(npi)
        if hs and npi not in npi_to_hs:
            npi_to_hs[npi]   = hs
            npi_to_city[npi] = city
    for k in name_key(last, frst):
        mailed_name_set.add(k)
    for k in fullname_keys(fn):
        mailed_name_set.add(k)

print(f"\n  Mailed NPIs : {len(mailed_npi_set)}")
print(f"  Mailed names: {len(mailed_name_set)}")

# ── Match each gold provider against the mailed list ──────────────────────────
match_flags  = []
match_methods = []
matched_hs   = []
matched_city = []

for _, row in gold.iterrows():
    npi  = norm(row[g_npi])   if g_npi   else ""
    last = norm(row[g_last])  if g_last  else ""
    frst = norm(row[g_first]) if g_first else ""

    if npi and npi in mailed_npi_set:
        match_flags.append(True)
        match_methods.append("npi")
        matched_hs.append(npi_to_hs.get(npi, ""))
        matched_city.append(npi_to_city.get(npi, ""))
        continue

    found_name = any(k in mailed_name_set for k in name_key(last, frst))
    if found_name:
        match_flags.append(True)
        match_methods.append("name")
        matched_hs.append("")
        matched_city.append("")
    else:
        match_flags.append(False)
        match_methods.append("not_mailed")
        matched_hs.append("")
        matched_city.append("")

gold["_in_mailed"]      = match_flags
gold["_match_method"]   = match_methods
gold["_mailed_hs"]      = matched_hs
gold["_mailed_hs_city"] = matched_city

not_mailed = gold[~gold["_in_mailed"]].copy()
in_mailed  = gold[gold["_in_mailed"]].copy()

print(f"\nResults:")
print(f"  In mailed list  : {len(in_mailed)}")
print(f"  NOT in mailed   : {len(not_mailed)}")

if g_src:
    print(f"\n  Not-mailed breakdown by source tab:")
    for src, cnt in not_mailed[g_src].value_counts().items():
        print(f"    {cnt:4d}  {src}")

# ── Build summary ──────────────────────────────────────────────────────────────
summary_rows = []
if g_src:
    for src in gold[g_src].unique():
        sub = gold[gold[g_src] == src]
        summary_rows.append({
            "source_tab":    src,
            "total":         len(sub),
            "in_mailed":     sub["_in_mailed"].sum(),
            "not_in_mailed": (~sub["_in_mailed"]).sum(),
        })
summary_df = pd.DataFrame(summary_rows).sort_values("not_in_mailed", ascending=False)

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    not_mailed.to_excel(writer, sheet_name="not_in_mailed", index=False)
    in_mailed.to_excel(writer,  sheet_name="in_mailed",     index=False)
    summary_df.to_excel(writer, sheet_name="summary",        index=False)
    print(f"  not_in_mailed : {len(not_mailed)} rows")
    print(f"  in_mailed     : {len(in_mailed)} rows")
    print(f"  summary       : {len(summary_df)} source tabs")

print("\nDone.")
