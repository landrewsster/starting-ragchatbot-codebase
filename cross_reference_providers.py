#!/usr/bin/env python3
# ── Self-update ───────────────────────────────────────────────────────────────
import os, sys, urllib.request, ssl
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "cross_reference_providers.py")
try:
    _ctx = ssl._create_unverified_context()
    _new = urllib.request.urlopen(_URL, timeout=10, context=_ctx).read()
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

# Optional command-line arg: manual file name (just the filename, not full path)
# Usage: python3 cross_reference_providers.py ManualSearch_Part2_05212026.xlsx
if len(sys.argv) > 1:
    MANUAL_FILE = BASE / sys.argv[1]
    stem        = Path(sys.argv[1]).stem
    OUTPUT_FILE = BASE / f"provider_round3_crossref_{stem}.xlsx"
else:
    MANUAL_FILE = BASE / "clean_manual_list_NS.xlsx"
    OUTPUT_FILE = BASE / "provider_round3_crossref.xlsx"

print(f"Manual file : {MANUAL_FILE.name}")
print(f"Output file : {OUTPUT_FILE.name}")

# ── Helpers ───────────────────────────────────────────────────────────────────
STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq|fnp|crna|crnp)\.?\s*$",
    re.IGNORECASE,
)

def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def zip5(s) -> str:
    """Return the first 5 digits of a ZIP code, or empty string."""
    z = re.sub(r"\D", "", norm(s))
    return z[:5] if z else ""

def clean_name(s) -> str:
    s = norm(s)
    s = STRIP_SUFFIXES.sub("", s).strip()
    s = re.sub(r"[,.]", "", s).strip()
    return s

def find_col(df, candidates):
    low      = {c.lower(): c for c in df.columns}
    low_norm = {c.lower().replace(" ", "_"): c for c in df.columns}  # spaces → underscores
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
        if c.lower().replace(" ", "_") in low_norm:   # e.g. "Clinic_Address" finds "Clinic Address"
            return low_norm[c.lower().replace(" ", "_")]
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
# Case-insensitive sheet match for "ProviderList_Combined" / "ProviderList_COMBINED"
_sheet_lower = {s.lower(): s for s in xl.sheet_names}
sheet = _sheet_lower.get("providerlist_combined", xl.sheet_names[0])
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
s_addr = find_col(manual, ["primary_address_1", "Clinic_Address", "address", "Address", "address_line1"])
s_city = find_col(manual, ["primary_city", "Clinic_City", "city", "City"])
s_zip  = find_col(manual, ["zip5", "Clinic_Zip", "zip", "Zip"])
print(f"  last={s_last!r}  first={s_first!r}  middle={s_mid!r}")
print(f"  addr={s_addr!r}  city={s_city!r}  zip={s_zip!r}")

# ── Reshape to wide: one row per provider, multiple address columns ────────────
manual["_norm_key"] = manual.apply(
    lambda row: make_key(row[s_last], row[s_first] if s_first else ""), axis=1
)
before = len(manual)

addr_cols_to_skip = {c for c in [s_addr, s_city, s_zip] if c}
non_addr_cols     = [c for c in manual.columns if c not in addr_cols_to_skip]

wide_rows        = []
dropped_dup_addr = []   # same provider, exact same non-empty address repeated

for key, group in manual.groupby("_norm_key", sort=False):
    base = group.iloc[0][non_addr_cols].to_dict()
    seen = []
    for _, r in group.iterrows():
        a = norm(r[s_addr]) if s_addr else ""
        c = norm(r[s_city]) if s_city else ""
        z = zip5(r[s_zip])  if s_zip  else ""
        raw = (r[s_addr] if s_addr else "", r[s_city] if s_city else "", r[s_zip] if s_zip else "")
        if a and (a, c, z) in seen:
            # Exact duplicate non-empty address — log and skip
            dropped_dup_addr.append({s_last: r[s_last], s_first: r[s_first],
                                     "dup_address": raw[0], "dup_city": raw[1], "dup_zip": raw[2],
                                     "norm_addr": a, "norm_city": c, "norm_zip": z,
                                     "all_seen": " || ".join(f"{s[0]}|{s[1]}|{s[2]}" for s in seen)})
        elif (a, c, z) not in seen:
            # New address slot (including empty-address rows — kept so they're visible)
            seen.append((a, c, z))
            n = len(seen)
            base[f"address_{n}"] = raw[0]
            base[f"city_{n}"]    = raw[1]
            base[f"zip_{n}"]     = raw[2]
        # else: duplicate empty address (provider listed N times all with no address) — skip silently
    base["num_addresses"] = len(seen) if seen else 1
    wide_rows.append(base)

manual = pd.DataFrame(wide_rows).reset_index(drop=True)
after     = len(manual)
collapsed = before - after
print(f"  Unique providers   : {after}  (from {before} rows)")
print(f"  Rows collapsed     : {collapsed}")
if dropped_dup_addr:
    print(f"    {len(dropped_dup_addr)} dropped — duplicate address for same provider:")
    for d in dropped_dup_addr:
        print(f"      {d[s_last]}, {d[s_first]}  |  {d['dup_address']}, {d['dup_city']} {d['dup_zip']}")
if not dropped_dup_addr and collapsed > 0:
    print(f"    (all {collapsed} collapses produced new address slots → see address_N columns)")

# ── Match manual search providers against master list ─────────────────────────
print(f"\nMatching {after} manual search providers against master list ...")

matched_flags  = []
matched_names  = []
match_quality  = []   # "exact", "initial_match", or "suspect_middle"

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

        # Middle initial check
        s_mi = norm(row[s_mid])[0] if s_mid and norm(row[s_mid]) else ""
        m_mi = norm(m_m)[0]        if m_m                        else ""
        if s_mi and m_mi:
            quality = "exact" if s_mi == m_mi else "suspect_middle"
        elif not s_mi and not m_mi:
            quality = "no_middle_either_side"
        else:
            quality = "one_side_missing_middle"
        match_quality.append(quality)
    else:
        matched_flags.append(False)
        matched_names.append("")
        match_quality.append("")

manual["_in_master"]         = matched_flags
manual["_master_name_match"] = matched_names
manual["_match_quality"]     = match_quality
manual["_manual_middle"]     = manual[s_mid].apply(norm) if s_mid else ""

n_matched   = sum(matched_flags)
n_not_found = after - n_matched
print(f"  Matched to master list    : {n_matched}")
print(f"    exact                   : {match_quality.count('exact')}")
print(f"    no_middle_either_side   : {match_quality.count('no_middle_either_side')}")
print(f"    one_side_missing_middle : {match_quality.count('one_side_missing_middle')}")
print(f"    suspect_middle          : {match_quality.count('suspect_middle')}")
print(f"  NOT in master list        : {n_not_found}")
print(f"\nRow count check: {n_matched} matched + {n_not_found} not matched = {n_matched + n_not_found} unique providers")

# ── Build output dataframes ───────────────────────────────────────────────────
not_in_master = (
    manual[~manual["_in_master"]]
    .drop(columns=["_in_master", "_master_name_match", "_manual_middle", "_match_quality", "_norm_key"])
    .reset_index(drop=True)
)

matched_df = (
    manual[manual["_in_master"]]
    .rename(columns={
        "_master_name_match": "master_name_match",
        "_manual_middle":     "manual_middle_initial",
        "_match_quality":     "match_quality",
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
internal_cols = ["_in_master", "_master_name_match", "_manual_middle", "_match_quality", "_norm_key"]

# Providers with more than one address (address_2 column is populated)
addr_cols_present = [c for c in manual.columns if c.startswith("address_") or c.startswith("city_") or c.startswith("zip_")]
print(f"\n  address columns present: {addr_cols_present}")

if "address_2" in manual.columns:
    has_second_addr = manual["address_2"].apply(
        lambda x: pd.notna(x) and str(x).strip() not in ("", "nan")
    )
    multi_addr_providers = (
        manual[has_second_addr]
        .drop(columns=[c for c in internal_cols if c in manual.columns], errors="ignore")
        .reset_index(drop=True)
    )
    print(f"  Providers with 2+ addresses: {len(multi_addr_providers)}")
else:
    multi_addr_providers = pd.DataFrame(columns=[s_last, s_first])
    print(f"  No address_2 column found — no providers with 2+ addresses")

dropped_df = pd.DataFrame(dropped_dup_addr) if dropped_dup_addr else pd.DataFrame()

print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    not_in_master.to_excel(       writer, sheet_name="not_in_master",        index=False)
    matched_df.to_excel(          writer, sheet_name="matched",              index=False)
    multi_addr_providers.to_excel(writer, sheet_name="multiple_address_rows", index=False)
    if not dropped_df.empty:
        dropped_df.to_excel(      writer, sheet_name="dropped_rows",         index=False)
    print(f"  not_in_master         : {len(not_in_master)} providers")
    print(f"  matched               : {len(matched_df)} providers")
    print(f"  multiple_address_rows : {len(multi_addr_providers)} providers with 2+ addresses")
    if not dropped_df.empty:
        print(f"  dropped_rows          : {len(dropped_df)} rows collapsed (dup or empty address)")

# ── Row count summary (at end for easy reference) ─────────────────────────────
def _nonempty(col, df):
    return df[col].apply(lambda x: pd.notna(x) and str(x).strip() not in ("", "nan")).sum() if col in df.columns else 0

_has2 = _nonempty("address_2", manual)
_has3 = _nonempty("address_3", manual)
_has4 = _nonempty("address_4", manual)
_has5 = _nonempty("address_5", manual)
_exactly2  = _has2 - _has3
_exactly3  = _has3 - _has4
_exactly4  = _has4 - _has5
_exactly5p = _has5
_extra_multi = _exactly2 * 1 + _exactly3 * 2 + _exactly4 * 3 + _exactly5p * 4
_extra_drop  = len(dropped_dup_addr)
_total       = after + _extra_multi + _extra_drop

print(f"""
══ Row count reconciliation ══════════════════════════════════════
  {after} unique providers in output
  + {_extra_multi} extra rows from multi-address providers:""")
if _exactly2:  print(f"      {_exactly2} provider(s) with exactly 2 addresses → +{_exactly2} row(s)")
if _exactly3:  print(f"      {_exactly3} provider(s) with exactly 3 addresses → +{_exactly3 * 2} row(s)")
if _exactly4:  print(f"      {_exactly4} provider(s) with exactly 4 addresses → +{_exactly4 * 3} row(s)")
if _exactly5p: print(f"      {_exactly5p} provider(s) with 5+ addresses        → +{_exactly5p * 4} row(s)")
if not _has2:  print(f"      (none)")
if _extra_drop:
    print(f"  + {_extra_drop} dropped rows (exact duplicate address) → see 'dropped_rows' sheet")
print(f"  ─────────────────────────────────────────────────────────────")
print(f"  Total accounted for: {_total}  (manual file rows: {before})")
if _total == before:
    print(f"  ✓ Counts reconcile perfectly")
else:
    print(f"  ✗ Still {abs(before - _total)} row(s) unaccounted for — check terminal output above")
print(f"═════════════════════════════════════════════════════════════════")

print("\nDone.")
