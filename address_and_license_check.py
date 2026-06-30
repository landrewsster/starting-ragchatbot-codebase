#!/usr/bin/env python3
# ── Self-update ─────────────────────────────────────────────────────────────────────────────
import os, sys, urllib.request, ssl
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "address_and_license_check.py")
try:
    _ctx = ssl._create_unverified_context()
    _new = urllib.request.urlopen(_URL, timeout=10, context=_ctx).read()
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

1. ADDRESS COMPARISON (matched sheet from provider_round3_crossref.xlsx)
   For each provider in the matched sheet, look them up in the Round 3
   mailing list and multiple_addresses.csv and compare addresses.

2. LICENSE BOARD CHECK (not_in_master sheet from provider_round3_crossref.xlsx)
   Name-match providers against the MN Physician and PA list.

Output: address_and_license_check.xlsx
  r3_addr_same        — matched providers with same address as Round 3
  r3_addr_different   — matched providers with different address from Round 3
  multi_addr_same     — matched providers with same address as multiple_addresses
  multi_addr_different— matched providers with different address from multiple_addresses
  in_license_board    — not-in-master providers found in MN license file
  not_in_license      — not-in-master providers not found in license file

Usage:
    python3 address_and_license_check.py
"""

import re
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────────────
BASE            = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE     = BASE / "MailingList_Round3_20260519.xlsx"
MULTI_ADDR_FILE = BASE / "multiple_addresses.csv"
LICENSE_FILE    = BASE / "MN State Licensing Board" / "MN Physician and PA list March 2026.xlsx"

# Optional command-line arg: crossref file name (just the filename, not full path)
# Usage: python3 address_and_license_check.py provider_round3_crossref_ManualSearch_Part2_05212026.xlsx
if len(sys.argv) > 1:
    CROSSREF_FILE = BASE / sys.argv[1]
    stem          = Path(sys.argv[1]).stem
    OUTPUT_FILE   = BASE / f"address_and_license_check_{stem}.xlsx"
else:
    CROSSREF_FILE = BASE / "Manual_List_Complete_063026.xlsx"
    OUTPUT_FILE   = BASE / "address_and_license_check.xlsx"

print(f"Crossref file: {CROSSREF_FILE.name}")
print(f"Output file  : {OUTPUT_FILE.name}")

# ── Helpers ─────────────────────────────────────────────────────────────────────────────
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
# Ordinal words → numeric form so "First" == "1st", "Second" == "2nd", etc.
ORDINAL_MAP = [
    (r"\bfirst\b",    "1st"),  (r"\bsecond\b",  "2nd"),
    (r"\bthird\b",    "3rd"),  (r"\bfourth\b",  "4th"),
    (r"\bfifth\b",    "5th"),  (r"\bsixth\b",   "6th"),
    (r"\bseventh\b",  "7th"),  (r"\beighth\b",  "8th"),
    (r"\bninth\b",    "9th"),  (r"\btenth\b",   "10th"),
    (r"\beleventh\b", "11th"), (r"\btwelfth\b", "12th"),
]

def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

# Matches a house number followed immediately by a directional, then street name:
# e.g. "1026 w 7th st" → group(1)="1026" group(2)="w" group(3)="7th st"
# Reorder to post-directional form: "1026 7th st w"
_PRE_DIR_RE = re.compile(
    r'^(\d+\w*)\s+(n|s|e|w|ne|nw|se|sw)\s+(.+)$'
)

def norm_addr(s) -> str:
    s = norm(s)
    for pattern, repl in ORDINAL_MAP:   # "First" → "1st" etc. before other subs
        s = re.sub(pattern, repl, s)
    for pattern, repl in ADDR_ABBREVS:
        s = re.sub(pattern, repl, s)
    s = re.sub(r"[.,#\-]", " ", s)      # strip punctuation that causes false mismatches
    s = re.sub(r"\s+", " ", s).strip()
    # Normalize pre-directional to post-directional:
    # "1026 W 7th St" and "1804 SW Trott Ave" → "1026 7th st w" / "1804 trott ave sw"
    m = _PRE_DIR_RE.match(s)
    if m:
        s = f"{m.group(1)} {m.group(3)} {m.group(2)}"
    return s

def zip5(s) -> str:
    return re.sub(r"\D", "", norm(s))[:5]

CITY_MAP = [
    (r"\bsaint\b", "st"),   # Saint Paul → St Paul
    (r"\bmount\b", "mt"),   # Mount Pleasant → Mt Pleasant
    (r"\bfort\b",  "ft"),   # Fort Snelling → Ft Snelling
]

def norm_city(s) -> str:
    s = norm(s)
    s = re.sub(r"\.", "", s)          # strip periods: "St." → "St"
    for pattern, repl in CITY_MAP:
        s = re.sub(pattern, repl, s)
    return re.sub(r"\s+", " ", s).strip()

FUZZY_THRESHOLD = 0.82   # similarity ratio to flag as "possible_match"

def addr_similarity(a1: str, a2: str) -> float:
    """Return 0–1 similarity between two normalised address strings."""
    return SequenceMatcher(None, norm_addr(a1), norm_addr(a2)).ratio()

def compare_address(manual_addrs: list, ref_addr: str, ref_city: str, ref_zip: str) -> str:
    """
    Compare a list of (addr, city, zip) tuples against a single reference address.
    Returns: 'same', 'possible_match', or 'different'.

    Matching logic (in order):
      1. Exact match after full normalisation (case, punctuation, abbreviations, ordinals, city aliases)
         Suite/unit numbers must match exactly — missing or different suite → not same.
      2. Fuzzy similarity >= FUZZY_THRESHOLD → 'possible_match' for manual review
    """
    if not ref_addr:
        return "different"
    ref_norm = f"{norm_addr(ref_addr)}|{norm_city(ref_city)}|{zip5(ref_zip)}"
    best_sim = 0.0
    for a, c, z in manual_addrs:
        m_norm = f"{norm_addr(a)}|{norm_city(c)}|{zip5(z)}"
        if m_norm == ref_norm:
            return "same"
        sim = SequenceMatcher(None, m_norm, ref_norm).ratio()
        if sim > best_sim:
            best_sim = sim
    if best_sim >= FUZZY_THRESHOLD:
        return "possible_match"
    return "different"

def clean_name(s) -> str:
    s = norm(s)
    s = STRIP_SUFFIXES.sub("", s).strip()
    return re.sub(r"[,.]", "", s).strip()

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

def lf_keys(last: str, first: str) -> set:
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    fi = f1[0] if f1 else ""        # first initial only
    keys = set()
    if l and f:
        keys.add(f"{l}|{f}")
    if l and f1 and f1 != f:
        keys.add(f"{l}|{f1}")
    if l and fi:
        keys.add(f"{l}|~{fi}")      # prefix ~ marks initial-only keys
    return keys

def make_all_keys(col_a: str, col_b: str) -> set:
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

# ── Load crossref file ─────────────────────────────────────────────────────────────────────────────
# Supports two formats:
#   Old format: separate 'matched' and 'not_in_master' sheets
#   New format: single sheet — pivots multiple address rows into wide format,
#               then auto-classifies using Round 3 name lookup
print(f"\nLoading cross-reference file: {CROSSREF_FILE.name}")
try:
    _xl = pd.ExcelFile(CROSSREF_FILE)
except FileNotFoundError:
    raise SystemExit(f"ERROR: {CROSSREF_FILE} not found")
print(f"  Sheets: {_xl.sheet_names}")

_auto_classify = False   # will be True for single-sheet format
_all_df        = None    # holds combined df until R3 lookup is ready
_single_sheet  = None    # name of the single sheet (for reporting)

if "matched" in _xl.sheet_names:
    # ── Old two-sheet format ────────────────────────────────────────────────────────────────────
    matched = _xl.parse("matched", dtype=str).fillna("")
    print(f"  matched      : {len(matched)} rows | columns: {list(matched.columns)}")

    if "not_in_master" not in _xl.sheet_names:
        raise SystemExit("ERROR: 'not_in_master' sheet not found in cross-reference file")
    not_matched = _xl.parse("not_in_master", dtype=str).fillna("")
    print(f"  not_in_master: {len(not_matched)} rows | columns: {list(not_matched.columns)}")

    m_last  = find_col(matched, ["Last_Name", "last_name", "LastName", "Last Name"]) or matched.columns[0]
    m_first = find_col(matched, ["First_Name", "first_name", "FirstName", "First Name"]) or matched.columns[1]
    m_mid   = find_col(matched, ["middle_name", "middle", "middle_initial", "manual_middle_initial"])

    # Deduplicate matched by name
    _m_key  = matched[m_last].apply(clean_name) + "|" + matched[m_first].apply(clean_name)
    _m_dups = _m_key.duplicated(keep="first")
    n_matched_dups = int(_m_dups.sum())
    _dups_df = matched[_m_dups].copy()
    _dups_df.insert(0, "source_sheet", "matched")
    if n_matched_dups:
        print(f"  *** {n_matched_dups} duplicate(s) removed from matched:")
        for _, row in matched[_m_dups].iterrows():
            print(f"      {row[m_last]}, {row[m_first]}")
        matched = matched[~_m_dups].reset_index(drop=True)
    else:
        print(f"  No duplicates in matched sheet")

    nm_last  = find_col(not_matched, ["Last_Name", "last_name", "LastName", "Last Name"]) or not_matched.columns[0]
    nm_first = find_col(not_matched, ["First_Name", "first_name", "FirstName", "First Name"]) or not_matched.columns[1]
    _nm_col2 = not_matched.columns[2] if len(not_matched.columns) > 2 else None
    nm_third = _nm_col2 if _nm_col2 and _nm_col2 not in (nm_last, nm_first) else None

    # Deduplicate not_matched by name
    _nm_key  = not_matched[nm_last].apply(clean_name) + "|" + not_matched[nm_first].apply(clean_name)
    _nm_dups = _nm_key.duplicated(keep="first")
    n_not_matched_dups = int(_nm_dups.sum())
    _nm_dups_df = not_matched[_nm_dups].copy()
    _nm_dups_df.insert(0, "source_sheet", "not_in_master")
    _dups_df = pd.concat([_dups_df, _nm_dups_df], ignore_index=True)
    if n_not_matched_dups:
        print(f"  *** {n_not_matched_dups} duplicate(s) removed from not_in_master:")
        for _, row in not_matched[_nm_dups].iterrows():
            print(f"      {row[nm_last]}, {row[nm_first]}")
        not_matched = not_matched[~_nm_dups].reset_index(drop=True)
    else:
        print(f"  No duplicates in not_in_master sheet")
    _merged_df = pd.DataFrame()  # not used in two-sheet path

else:
    # ── New single-sheet format ────────────────────────────────────────────────────────────────────────
    # Pivot: collapse multiple rows per provider into one row with numbered address
    # columns (address_1/city_1/zip_1, address_2/city_2/zip_2, …) so every
    # address is available for comparison.
    _single_sheet = _xl.sheet_names[0]
    _all_df_raw   = _xl.parse(_single_sheet, dtype=str).fillna("")
    print(f"  Single sheet '{_single_sheet}': {len(_all_df_raw)} rows | columns: {list(_all_df_raw.columns)}")

    m_last  = find_col(_all_df_raw, ["Last_Name", "last_name", "LastName", "Last Name"]) or _all_df_raw.columns[0]
    m_first = find_col(_all_df_raw, ["First_Name", "first_name", "FirstName", "First Name"]) or _all_df_raw.columns[1]
    m_mid   = find_col(_all_df_raw, ["Middle_Name", "middle_name", "middle", "middle_initial"])
    nm_last, nm_first = m_last, m_first
    _nm_col2 = _all_df_raw.columns[2] if len(_all_df_raw.columns) > 2 else None
    nm_third = _nm_col2 if _nm_col2 and _nm_col2 not in (nm_last, nm_first) else None

    # Detect address columns before pivot
    _s_addr   = find_col(_all_df_raw, ["Clinic_Address", "primary_address_1", "address", "Address"])
    _s_city   = find_col(_all_df_raw, ["Clinic_City", "primary_city", "city", "City"])
    _s_state  = find_col(_all_df_raw, ["Clinic_State", "state", "State"])
    _s_zip    = find_col(_all_df_raw, ["Clinic_Zip", "zip5", "zip", "Zip"])
    _s_clinic = find_col(_all_df_raw, ["Clinic_name_location", "clinic_name", "location"])
    _s_addr_cols = [c for c in [_s_addr, _s_city, _s_state, _s_zip, _s_clinic] if c]
    print(f"  Address cols detected: addr={_s_addr!r} city={_s_city!r} state={_s_state!r} zip={_s_zip!r}")

    # Build name key and collect groups in first-occurrence order
    _all_df_raw["_name_key"] = (
        _all_df_raw[m_last].apply(clean_name) + "|" + _all_df_raw[m_first].apply(clean_name)
    )
    _seen_keys    = list(dict.fromkeys(_all_df_raw["_name_key"]))
    _pivoted_rows = []
    _merged_records = []
    n_collapsed   = 0   # extra rows folded into the first (total_raw - unique_providers)

    for key in _seen_keys:
        group = _all_df_raw[_all_df_raw["_name_key"] == key]

        # Start from the first row; remove original address cols (replaced by numbered ones)
        first = group.iloc[0].to_dict()
        for c in _s_addr_cols:
            first.pop(c, None)
        first.pop("_name_key", None)

        # Collect unique addresses across all rows for this provider
        seen_sigs = set()
        addr_num  = 0
        for _, sub_row in group.iterrows():
            a  = sub_row.get(_s_addr,   "") if _s_addr   else ""
            c  = sub_row.get(_s_city,   "") if _s_city   else ""
            st = sub_row.get(_s_state,  "") if _s_state  else ""
            z  = sub_row.get(_s_zip,    "") if _s_zip    else ""
            cl = sub_row.get(_s_clinic, "") if _s_clinic else ""
            if not norm(a):
                continue
            sig = f"{norm_addr(a)}|{norm_city(c)}|{zip5(z)}"
            if sig not in seen_sigs:
                seen_sigs.add(sig)
                addr_num += 1
                first[f"address_{addr_num}"] = a
                first[f"city_{addr_num}"]    = c
                first[f"state_{addr_num}"]   = st
                first[f"zip_{addr_num}"]     = z
                if _s_clinic:
                    first[f"clinic_name_{addr_num}"] = cl

        if addr_num == 0:  # provider had no address in any row
            first["address_1"] = first["city_1"] = first["state_1"] = first["zip_1"] = ""

        extra = len(group) - 1
        n_collapsed += extra
        if len(group) > 1:
            note = ("true duplicate (same address)" if addr_num <= 1
                    else f"merged {addr_num} unique addresses")
            _merged_records.append({
                "name":             f"{group.iloc[0][m_last]}, {group.iloc[0][m_first]}",
                "rows_in_source":   len(group),
                "unique_addresses": addr_num,
                "note":             note,
            })
        _pivoted_rows.append(first)

    _all_df    = pd.DataFrame(_pivoted_rows).fillna("").reset_index(drop=True)
    _merged_df = (pd.DataFrame(_merged_records) if _merged_records
                  else pd.DataFrame(columns=["name", "rows_in_source", "unique_addresses", "note"]))
    print(f"  {len(_all_df)} unique providers after pivot ({n_collapsed} extra rows collapsed)")
    if _merged_records:
        _show = _merged_records[:20]
        for rec in _show:
            print(f"    {rec['name']}: {rec['rows_in_source']} rows → "
                  f"{rec['unique_addresses']} addr ({rec['note']})")
        if len(_merged_records) > 20:
            print(f"    ... and {len(_merged_records)-20} more (see merged_addresses sheet)")

    _auto_classify     = True
    n_matched_dups     = n_collapsed
    n_not_matched_dups = 0
    matched = not_matched = pd.DataFrame()  # placeholders until R3 lookup is ready
    _dups_df = pd.DataFrame()              # not used in single-sheet path

# Detect address columns (same logic applies for both formats)
_addr_src = matched if not _auto_classify else _all_df
m_addr_sets = []
i = 1
while True:
    a = find_col(_addr_src, [f"address_{i}", f"primary_address_{i}"])
    c = find_col(_addr_src, [f"city_{i}"])
    z = find_col(_addr_src, [f"zip_{i}"])
    if a:
        m_addr_sets.append((a, c, z))
        i += 1
    else:
        break
if not m_addr_sets:
    a = find_col(_addr_src, ["primary_address_1", "Clinic_Address", "address", "Address", "address_line1"])
    c = find_col(_addr_src, ["primary_city", "Clinic_City", "city", "City"])
    z = find_col(_addr_src, ["zip5", "Clinic_Zip", "zip", "Zip"])
    if a:
        m_addr_sets.append((a, c, z))

print(f"  name cols : last={m_last!r}  first={m_first!r}  mid={m_mid!r}")
print(f"  addr sets : {len(m_addr_sets)} — {[(a,c,z) for a,c,z in m_addr_sets]}")

# ── Load Round 3 for address lookup ──────────────────────────────────────────────────────────────────
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
try:
    r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {ROUND3_FILE} not found")
print(f"  {len(r3)} rows | columns: {list(r3.columns)}")

r3_last  = find_col(r3, ["last_name", "LastName", "Last Name"]) or r3.columns[0]
r3_first = find_col(r3, ["first_name", "FirstName", "First Name"]) or r3.columns[1]
# Address may appear in any of columns D/E/F — collect all three, use first non-empty at runtime
r3_addr_cols = [c for c in [
    find_col(r3, ["primary_address_1", "address_line1"]),
    find_col(r3, ["primary_address_2", "address_line2"]),
    find_col(r3, ["address_line3", "address_line_3"]),
] if c]
r3_city  = find_col(r3, ["primary_city", "City", "city"])
r3_zip   = find_col(r3, ["zip5", "ZIP+4", "zip", "Zip"])
print(f"  last={r3_last!r}  first={r3_first!r}")
print(f"  addr_cols={r3_addr_cols}  city={r3_city!r}  zip={r3_zip!r}")

r3_key_to_idx = {}
for idx, row in r3.iterrows():
    for key in make_all_keys(row[r3_last], row[r3_first]):
        r3_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(r3_key_to_idx)}")

# ── Auto-classify single-sheet providers using Round 3 lookup ────────────────────────────────────
if _auto_classify:
    _in_r3 = _all_df.apply(
        lambda row: any(k in r3_key_to_idx
                        for k in make_all_keys(row[m_last], row[m_first])),
        axis=1
    )
    matched     = _all_df[_in_r3].reset_index(drop=True)
    not_matched = _all_df[~_in_r3].reset_index(drop=True)
    print(f"\n  Auto-classified {len(_all_df)} providers using Round 3 name lookup:")
    print(f"    found in Round 3 → matched      : {len(matched)}")
    print(f"    not in Round 3  → not_in_master : {len(not_matched)}")

# ── Compare addresses for matched providers (manual vs Round 3) ─────────────────────────────
print(f"\nComparing addresses for {len(matched)} matched providers ...")

addr_rows = []
for _, row in matched.iterrows():
    keys      = make_all_keys(row[m_last], row[m_first])
    found_idx = next((r3_key_to_idx[k] for k in keys if k in r3_key_to_idx), None)

    manual_name = f"{row[m_last]}, {row[m_first]}"
    if m_mid and row.get(m_mid, ""):
        manual_name += f" {row[m_mid]}"

    # Collect all manual addresses (wide format: address_1, address_2, ...)
    manual_addrs = []
    for a_col, c_col, z_col in m_addr_sets:
        a = row.get(a_col, "") if a_col else ""
        c = row.get(c_col, "") if c_col else ""
        z = row.get(z_col, "") if z_col else ""
        if norm(a):
            manual_addrs.append((a, c, z))

    if found_idx is not None:
        r3_row  = r3.iloc[found_idx]
        r3_name = f"{r3_row[r3_last]}, {r3_row[r3_first]}"
        r3_city_raw = r3_row[r3_city] if r3_city else ""
        r3_zip_raw  = r3_row[r3_zip]  if r3_zip  else ""

        # Use first non-empty address column from D/E/F
        r3_addr_raw = next(
            (r3_row[c] for c in r3_addr_cols if norm(r3_row[c])), ""
        )

        match = compare_address(manual_addrs, r3_addr_raw, r3_city_raw, r3_zip_raw)

        r3_norm_key = f"{norm_addr(r3_addr_raw)}|{norm_city(r3_city_raw)}|{zip5(r3_zip_raw)}"
        out = {"manual_name": manual_name, "r3_name": r3_name,
               "address_match": match,
               "r3_address": r3_addr_raw, "r3_city": r3_city_raw, "r3_zip": r3_zip_raw,
               "_debug_r3_norm": r3_norm_key}
        for i, (a, c, z) in enumerate(manual_addrs, start=1):
            out[f"manual_address_{i}"] = a
            out[f"manual_city_{i}"]    = c
            out[f"manual_zip_{i}"]     = z
            out[f"_debug_manual_norm_{i}"] = f"{norm_addr(a)}|{norm_city(c)}|{zip5(z)}"
        addr_rows.append(out)
    else:
        out = {"manual_name": manual_name, "r3_name": "(not found in Round 3)",
               "address_match": "not_found", "r3_address": "", "r3_city": "", "r3_zip": ""}
        for i, (a, c, z) in enumerate(manual_addrs, start=1):
            out[f"manual_address_{i}"] = a
            out[f"manual_city_{i}"]    = c
            out[f"manual_zip_{i}"]     = z
        addr_rows.append(out)

addr_df = pd.DataFrame(addr_rows)
n_same    = (addr_df["address_match"] == "same").sum()
n_possible= (addr_df["address_match"] == "possible_match").sum()
n_diff    = (addr_df["address_match"] == "different").sum()
n_nf      = (addr_df["address_match"] == "not_found").sum()
print(f"  Same address        : {n_same}")
print(f"  Possible match      : {n_possible}  (similar but not identical — review manually)")
print(f"  Different address   : {n_diff}")
print(f"  Not found in Round 3: {n_nf}")

# ── Load multiple_addresses.csv for second address comparison ─────────────────────────────
print(f"\nLoading multiple addresses file: {MULTI_ADDR_FILE.name}")
try:
    multi = pd.read_csv(MULTI_ADDR_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {MULTI_ADDR_FILE} not found")
print(f"  {len(multi)} rows | columns: {list(multi.columns)}")

ma_last  = find_col(multi, ["last_name", "LastName", "Last Name"]) or multi.columns[0]
ma_first = find_col(multi, ["first_name", "FirstName", "First Name"]) or multi.columns[1]
# Address 1: multiple line components (cols C-F), single city/zip
ma_addr1_lines = [c for c in [
    find_col(multi, ["mail_address"]),
    find_col(multi, ["mail_address.1"]),
    find_col(multi, ["mail_address.2"]),
    find_col(multi, ["mail_address.3"]),
] if c]
ma_city1 = find_col(multi, ["mail_city", "primary_city", "City", "city"])
ma_zip1  = find_col(multi, ["mail_zip", "zip5", "zip", "Zip"])
# Address 2: col I with its own city/zip
ma_addr2 = find_col(multi, ["addr_2"])
ma_city2 = find_col(multi, ["city_2"])
ma_zip2  = find_col(multi, ["zip_2"])
# Address 3: col L with its own city/zip
ma_addr3 = find_col(multi, ["addr_3"])
ma_city3 = find_col(multi, ["city_3"])
ma_zip3  = find_col(multi, ["zip_3"])
print(f"  last={ma_last!r}  first={ma_first!r}")
print(f"  addr1_lines={ma_addr1_lines}  city1={ma_city1!r}  zip1={ma_zip1!r}")
print(f"  addr2={ma_addr2!r}  city2={ma_city2!r}  zip2={ma_zip2!r}")
print(f"  addr3={ma_addr3!r}  city3={ma_city3!r}  zip3={ma_zip3!r}")

multi_key_to_idx = {}
for idx, row in multi.iterrows():
    for key in make_all_keys(row[ma_last], row[ma_first]):
        multi_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(multi_key_to_idx)}")

# ── Compare addresses against multiple_addresses (secondary check for R3 non-matches) ──
# Only run on providers whose R3 address was not confirmed same —
# avoids double-counting and uses multi file as a second source to resolve R3 discrepancies.
r3_non_same_names = set(
    addr_df[addr_df["address_match"] != "same"]["manual_name"]
)
print(f"\nComparing R3 non-same providers ({len(r3_non_same_names)}) against multiple_addresses file ...")

multi_addr_rows = []
for _, row in matched.iterrows():
    manual_name = f"{row[m_last]}, {row[m_first]}"
    if m_mid and row.get(m_mid, ""):
        manual_name += f" {row[m_mid]}"

    # Skip providers already confirmed same in R3
    if manual_name not in r3_non_same_names:
        continue

    keys      = make_all_keys(row[m_last], row[m_first])
    found_idx = next((multi_key_to_idx[k] for k in keys if k in multi_key_to_idx), None)

    # Collect all manual addresses (wide format: address_1, address_2, ...)
    manual_addrs = []
    for a_col, c_col, z_col in m_addr_sets:
        a = row.get(a_col, "") if a_col else ""
        c = row.get(c_col, "") if c_col else ""
        z = row.get(z_col, "") if z_col else ""
        if norm(a):
            manual_addrs.append((a, c, z))

    if found_idx is not None:
        ma_row  = multi.iloc[found_idx]
        ma_name = f"{ma_row[ma_last]}, {ma_row[ma_first]}"

        # Build all reference addresses for this provider
        ref_addrs = []
        a1 = next((ma_row[c] for c in ma_addr1_lines if norm(ma_row[c])), "")
        if a1:
            ref_addrs.append((a1,
                              ma_row[ma_city1] if ma_city1 else "",
                              ma_row[ma_zip1]  if ma_zip1  else ""))
        if ma_addr2 and norm(ma_row[ma_addr2]):
            ref_addrs.append((ma_row[ma_addr2],
                              ma_row[ma_city2] if ma_city2 else "",
                              ma_row[ma_zip2]  if ma_zip2  else ""))
        if ma_addr3 and norm(ma_row[ma_addr3]):
            ref_addrs.append((ma_row[ma_addr3],
                              ma_row[ma_city3] if ma_city3 else "",
                              ma_row[ma_zip3]  if ma_zip3  else ""))

        # Compare manual addresses against ALL reference addresses, take best result
        match = "different"
        for ref_a, ref_c, ref_z in ref_addrs:
            result = compare_address(manual_addrs, ref_a, ref_c, ref_z)
            if result == "same":
                match = "same"
                break
            if result == "possible_match":
                match = "possible_match"

        out = {"manual_name": manual_name, "multi_name": ma_name, "address_match": match}
        for i, (a, c, z) in enumerate(manual_addrs, start=1):
            out[f"manual_address_{i}"] = a
            out[f"manual_city_{i}"]    = c
            out[f"manual_zip_{i}"]     = z
        for i, (a, c, z) in enumerate(ref_addrs, start=1):
            out[f"multi_address_{i}"] = a
            out[f"multi_city_{i}"]    = c
            out[f"multi_zip_{i}"]     = z
        multi_addr_rows.append(out)
    else:
        out = {"manual_name": manual_name, "multi_name": "(not found in multiple_addresses)",
               "address_match": "not_found"}
        for i, (a, c, z) in enumerate(manual_addrs, start=1):
            out[f"manual_address_{i}"] = a
            out[f"manual_city_{i}"]    = c
            out[f"manual_zip_{i}"]     = z
        multi_addr_rows.append(out)

multi_df    = pd.DataFrame(multi_addr_rows)
mn_same     = (multi_df["address_match"] == "same").sum()
mn_possible = (multi_df["address_match"] == "possible_match").sum()
mn_diff     = (multi_df["address_match"] == "different").sum()
mn_nf       = (multi_df["address_match"] == "not_found").sum()
print(f"  Same address             : {mn_same}")
print(f"  Possible match           : {mn_possible}  (similar but not identical — review manually)")
print(f"  Different address        : {mn_diff}")
print(f"  Not found in multi-addr  : {mn_nf}")

# ── Load MN licensing board ──────────────────────────────────────────────────────────────────────────
print(f"\nLoading MN licensing board: {LICENSE_FILE.name}")
try:
    lic_xl = pd.ExcelFile(LICENSE_FILE)
    print(f"  Sheets: {lic_xl.sheet_names}")
    lic_df = lic_xl.parse(lic_xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {LICENSE_FILE} not found")
print(f"  {len(lic_df)} rows | columns: {list(lic_df.columns)}")

l_last      = find_col(lic_df, ["last_name", "LastName", "Last Name"]) or lic_df.columns[0]
l_first     = find_col(lic_df, ["first_name", "FirstName", "First Name"]) or lic_df.columns[1]
# Specialty boards = col J (index 9), Certification = col K (index 10)
l_specialty = lic_df.columns[9]  if len(lic_df.columns) > 9  else None
l_cert      = lic_df.columns[10] if len(lic_df.columns) > 10 else None
l_email     = lic_df.columns[11] if len(lic_df.columns) > 11 else None
print(f"  last={l_last!r}  first={l_first!r}")
print(f"  specialty_boards col (J): {l_specialty!r}")
print(f"  certification col    (K): {l_cert!r}")
print(f"  email col            (L): {l_email!r}")
print(f"  sample last values : {lic_df[l_last].dropna().head(3).tolist()}")
print(f"  sample first values: {lic_df[l_first].dropna().head(3).tolist()}")

lic_key_to_idx = {}
for idx, row in lic_df.iterrows():
    for key in make_all_keys(row[l_last], row[l_first]):
        lic_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(lic_key_to_idx)}")

# Collect debug info for Baker — printed at end
_debug_baker_lic_keys  = [k for k in lic_key_to_idx if "baker" in k]
_debug_baker_lic_rows  = lic_df[lic_df[l_last].str.lower().str.contains("baker", na=False)
                                | lic_df[l_first].str.lower().str.contains("baker", na=False)]
_debug_baker_nm_keys   = []

# ── Match not_in_master against licensing board ───────────────────────────────────────────────
print(f"\nChecking {len(not_matched)} not-in-master providers against license board ...")

in_license     = []
not_in_license = []

for _, row in not_matched.iterrows():
    # Try col_a + col_b and also col_a + col_c, since col_b may be a middle
    # initial and col_c the actual first or last name (mixed format)
    keys = make_all_keys(row[nm_last], row[nm_first])
    if nm_third:
        keys |= make_all_keys(row[nm_last],  row[nm_third])
        keys |= make_all_keys(row[nm_third], row[nm_last])
        keys |= make_all_keys(row[nm_third], row[nm_first])

    # Collect Baker debug info
    if "baker" in norm(row[nm_last]) or "baker" in norm(row[nm_first]) or \
       (nm_third and "baker" in norm(row[nm_third])):
        _debug_baker_nm_keys.append(sorted(keys))

    # Prefer exact keys first, fall back to initial (~) keys
    exact_keys   = {k for k in keys if "|~" not in k}
    initial_keys = {k for k in keys if "|~" in k}
    found_idx    = next((lic_key_to_idx[k] for k in exact_keys   if k in lic_key_to_idx), None)
    match_type   = "exact"
    if found_idx is None:
        found_idx  = next((lic_key_to_idx[k] for k in initial_keys if k in lic_key_to_idx), None)
        match_type = "initial_match"

    if found_idx is not None:
        lic_row   = lic_df.iloc[found_idx]
        lic_name  = f"{lic_row[l_last]}, {lic_row[l_first]}".strip(", ")
        specialty = lic_row[l_specialty] if l_specialty else ""
        cert      = lic_row[l_cert]      if l_cert      else ""
        email     = lic_row[l_email]     if l_email     else ""
        row2      = row.copy()
        row2["license_name_match"]       = lic_name
        row2["match_type"]               = match_type
        row2["specialty_boards"]         = specialty
        row2["certification"]            = cert
        row2["email"]                    = email
        row2["specialty_boards_missing"] = "missing" if not norm(specialty) else ""
        row2["certification_missing"]    = "missing" if not norm(cert)      else ""
        in_license.append(row2)
    else:
        row2 = row.copy()
        row2["specialty_boards"]         = ""
        row2["certification"]            = ""
        row2["email"]                    = ""
        row2["specialty_boards_missing"] = ""
        row2["certification_missing"]    = ""
        not_in_license.append(row2)

in_lic_df     = pd.DataFrame(in_license).reset_index(drop=True)
not_in_lic_df = pd.DataFrame(not_in_license).reset_index(drop=True)
print(f"  Found in license board : {len(in_lic_df)}")
print(f"  Not in license board   : {len(not_in_lic_df)}")
print(f"  Total covered          : {len(in_lic_df) + len(not_in_lic_df)}  (should equal {len(not_matched)})")

# Specialty completeness for providers found in license board
if not in_lic_df.empty and "specialty_boards" in in_lic_df.columns:
    n_spec_present = in_lic_df["specialty_boards"].apply(norm).ne("").sum()
    n_spec_missing = len(in_lic_df) - n_spec_present
    print(f"\n  Specialty boards (in_license_board):")
    print(f"    Has specialty listed : {n_spec_present}")
    print(f"    Blank / missing      : {n_spec_missing}")
else:
    n_spec_present = 0
    n_spec_missing = 0

# ── Also check matched providers against licensing board ────────────────────────────────────
print(f"\nChecking {len(matched)} matched providers against license board ...")

matched_in_license     = []
matched_not_in_license = []

for _, row in matched.iterrows():
    keys         = make_all_keys(row[m_last], row[m_first])
    exact_keys   = {k for k in keys if "|~" not in k}
    initial_keys = {k for k in keys if "|~" in k}
    found_idx    = next((lic_key_to_idx[k] for k in exact_keys   if k in lic_key_to_idx), None)
    match_type   = "exact"
    if found_idx is None:
        found_idx  = next((lic_key_to_idx[k] for k in initial_keys if k in lic_key_to_idx), None)
        match_type = "initial_match"

    if found_idx is not None:
        lic_row   = lic_df.iloc[found_idx]
        lic_name  = f"{lic_row[l_last]}, {lic_row[l_first]}".strip(", ")
        specialty = lic_row[l_specialty] if l_specialty else ""
        cert      = lic_row[l_cert]      if l_cert      else ""
        email     = lic_row[l_email]     if l_email     else ""
        row2      = row.copy()
        row2["license_name_match"]       = lic_name
        row2["match_type"]               = match_type
        row2["specialty_boards"]         = specialty
        row2["certification"]            = cert
        row2["email"]                    = email
        row2["specialty_boards_missing"] = "missing" if not norm(specialty) else ""
        row2["certification_missing"]    = "missing" if not norm(cert)      else ""
        matched_in_license.append(row2)
    else:
        row2 = row.copy()
        row2["specialty_boards"]         = ""
        row2["certification"]            = ""
        row2["email"]                    = ""
        row2["specialty_boards_missing"] = ""
        row2["certification_missing"]    = ""
        matched_not_in_license.append(row2)

m_in_lic_df     = pd.DataFrame(matched_in_license).reset_index(drop=True)
m_not_in_lic_df = pd.DataFrame(matched_not_in_license).reset_index(drop=True)
print(f"  Found in license board : {len(m_in_lic_df)}")
print(f"  Not in license board   : {len(m_not_in_lic_df)}")

if not m_in_lic_df.empty and "specialty_boards" in m_in_lic_df.columns:
    n_m_spec_present = m_in_lic_df["specialty_boards"].apply(norm).ne("").sum()
    n_m_spec_missing = len(m_in_lic_df) - n_m_spec_present
    print(f"\n  Specialty boards (matched_in_license):")
    print(f"    Has specialty listed : {n_m_spec_present}")
    print(f"    Blank / missing      : {n_m_spec_missing}")
else:
    n_m_spec_present = 0
    n_m_spec_missing = 0

# ── Terminal summary ────────────────────────────────────────────────────────────────────────────
print(f"\nProviders with POSSIBLE address match ({n_possible}) — review manually:")
for _, row in addr_df[addr_df["address_match"] == "possible_match"].iterrows():
    print(f"  {row['manual_name']}")
    i = 1
    while f"manual_address_{i}" in row.index and row[f"manual_address_{i}"]:
        print(f"    Manual{i}: {row[f'manual_address_{i}']}, {row.get(f'manual_city_{i}','')} {row.get(f'manual_zip_{i}','')}")
        i += 1
    print(f"    Round3 : {row['r3_address']}, {row['r3_city']} {row['r3_zip']}")

print(f"\nNot-in-master found in license board ({len(in_lic_df)}):")
for _, row in in_lic_df.iterrows():
    print(f"  {row[nm_last]}, {row[nm_first]}  →  {row['license_name_match']}")

# ── Build output subsets and summary ─────────────────────────────────────────────────────────────────
r3_same        = addr_df[addr_df["address_match"] == "same"].reset_index(drop=True)
r3_possible    = addr_df[addr_df["address_match"] == "possible_match"].reset_index(drop=True)
r3_diff        = addr_df[addr_df["address_match"].isin(["different", "not_found"])].reset_index(drop=True)
multi_same     = multi_df[multi_df["address_match"] == "same"].reset_index(drop=True)
multi_possible = multi_df[multi_df["address_match"] == "possible_match"].reset_index(drop=True)
multi_diff     = multi_df[multi_df["address_match"] == "different"].reset_index(drop=True)
multi_notfound = multi_df[multi_df["address_match"] == "not_found"].reset_index(drop=True)

n_matched_provs   = len(r3_same) + len(r3_possible) + len(r3_diff)
n_multi_checked   = len(multi_same) + len(multi_possible) + len(multi_diff) + len(multi_notfound)
n_unmatched_provs = len(in_lic_df) + len(not_in_lic_df)
n_grand_total     = n_matched_provs + n_unmatched_provs

_ok   = lambda n, exp: "OK" if n == exp else f"*** MISMATCH (got {n}, expected {exp}) ***"
_note = lambda n, exp: f"should = {exp} ({_ok(n, exp)})"

# INPUT section varies by format
if _auto_classify:
    _input_rows = [
        {"section": "INPUT",
         "category": f"source sheet '{_single_sheet}' (raw rows before pivot)",
         "count": len(_all_df) + n_matched_dups,
         "note": f"{n_matched_dups} extra rows collapsed → {len(_all_df)} unique providers"},
        {"section": "",
         "category": "  — extra rows merged (multiple addresses or true dups)",
         "count": n_matched_dups,
         "note": "see merged_addresses sheet for details"},
        {"section": "",
         "category": "found in Round 3 (auto-classified → matched)",
         "count": len(matched),
         "note": "name matched Round 3 mailing list"},
        {"section": "",
         "category": "not in Round 3 (auto-classified → not_in_master)",
         "count": len(not_matched),
         "note": "name did not match Round 3"},
        {"section": "",
         "category": "TOTAL unique providers",
         "count": len(matched) + len(not_matched),
         "note": ""},
    ]
else:
    _input_rows = [
        {"section": "INPUT",
         "category": "matched sheet (crossref input)",
         "count": len(matched),
         "note": f"rows after removing {n_matched_dups} duplicate(s)"},
        {"section": "",
         "category": "  — duplicates removed from matched",
         "count": n_matched_dups,
         "note": "same name appeared more than once in matched sheet"},
        {"section": "",
         "category": "not_in_master sheet (crossref input)",
         "count": len(not_matched),
         "note": f"rows after removing {n_not_matched_dups} duplicate(s)"},
        {"section": "",
         "category": "  — duplicates removed from not_in_master",
         "count": n_not_matched_dups,
         "note": "same name appeared more than once in not_in_master sheet"},
        {"section": "",
         "category": "TOTAL input (unique providers)",
         "count": len(matched) + len(not_matched),
         "note": "should equal total unique providers in manual list"},
    ]

summary_rows = _input_rows + [
    {"section": "", "category": "", "count": "", "note": ""},
    {"section": "ADDRESS CHECK (matched providers vs Round 3)",
     "category": "r3_addr_same",
     "count": len(r3_same),
     "note": "same address as Round 3 mailing list"},
    {"section": "",
     "category": "r3_addr_possible",
     "count": len(r3_possible),
     "note": "similar but not identical — review manually"},
    {"section": "",
     "category": "r3_addr_different",
     "count": len(r3_diff),
     "note": "different or not found in Round 3"},
    {"section": "",
     "category": "TOTAL",
     "count": n_matched_provs,
     "note": _note(n_matched_provs, len(matched))},
    {"section": "", "category": "", "count": "", "note": ""},
    {"section": "ADDRESS CHECK 2 (r3 non-same subset vs multiple_addresses.csv)",
     "category": "multi_addr_same",
     "count": len(multi_same),
     "note": "confirmed same via multiple_addresses.csv"},
    {"section": "",
     "category": "multi_addr_possible",
     "count": len(multi_possible),
     "note": "similar — review manually"},
    {"section": "",
     "category": "multi_addr_different",
     "count": len(multi_diff),
     "note": "found in file but address differs"},
    {"section": "",
     "category": "multi_addr_not_found",
     "count": len(multi_notfound),
     "note": "not found in multiple_addresses.csv"},
    {"section": "",
     "category": "TOTAL checked in multi file",
     "count": n_multi_checked,
     "note": f"subset of {len(r3_possible) + len(r3_diff)} r3 non-same providers (NOT additive with r3 totals above)"},
    {"section": "", "category": "", "count": "", "note": ""},
    {"section": "LICENSE CHECK — matched providers (Round 3) vs MN license board",
     "category": "matched_in_license",
     "count": len(m_in_lic_df),
     "note": "matched (R3) providers found in MN Physician/PA license file"},
    {"section": "",
     "category": "  — has specialty listed",
     "count": n_m_spec_present,
     "note": "specialty_boards column is non-blank"},
    {"section": "",
     "category": "  — specialty blank/missing",
     "count": n_m_spec_missing,
     "note": "specialty_boards column is blank"},
    {"section": "",
     "category": "matched_not_in_license",
     "count": len(m_not_in_lic_df),
     "note": "matched (R3) providers not found in license file"},
    {"section": "",
     "category": "TOTAL",
     "count": len(m_in_lic_df) + len(m_not_in_lic_df),
     "note": _note(len(m_in_lic_df) + len(m_not_in_lic_df), len(matched))},
    {"section": "", "category": "", "count": "", "note": ""},
    {"section": "LICENSE CHECK — not_in_master providers vs MN license board",
     "category": "in_license_board",
     "count": len(in_lic_df),
     "note": "not-in-master providers found in MN Physician/PA license file"},
    {"section": "",
     "category": "  — has specialty listed",
     "count": n_spec_present,
     "note": "specialty_boards column is non-blank"},
    {"section": "",
     "category": "  — specialty blank/missing",
     "count": n_spec_missing,
     "note": "specialty_boards column is blank"},
    {"section": "",
     "category": "not_in_license",
     "count": len(not_in_lic_df),
     "note": "not-in-master providers not found in license file"},
    {"section": "",
     "category": "TOTAL",
     "count": n_unmatched_provs,
     "note": _note(n_unmatched_provs, len(not_matched))},
    {"section": "", "category": "", "count": "", "note": ""},
    {"section": "GRAND TOTAL",
     "category": "matched providers + not_in_master providers",
     "count": n_grand_total,
     "note": "should equal total unique providers in manual list"},
]
summary_df = pd.DataFrame(summary_rows)

# ── Write output ───────────────────────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    summary_df.to_excel(     writer, sheet_name="summary",               index=False)
    if _auto_classify:
        _merged_df.to_excel( writer, sheet_name="merged_addresses",      index=False)
    else:
        _dups_df.to_excel(   writer, sheet_name="duplicates_removed",    index=False)
    r3_same.to_excel(        writer, sheet_name="r3_addr_same",          index=False)
    r3_possible.to_excel(    writer, sheet_name="r3_addr_possible",      index=False)
    r3_diff.to_excel(        writer, sheet_name="r3_addr_different",     index=False)
    multi_same.to_excel(     writer, sheet_name="multi_addr_same",       index=False)
    multi_possible.to_excel( writer, sheet_name="multi_addr_possible",   index=False)
    multi_diff.to_excel(     writer, sheet_name="multi_addr_different",  index=False)
    multi_notfound.to_excel( writer, sheet_name="multi_addr_not_found",  index=False)
    m_in_lic_df.to_excel(    writer, sheet_name="matched_in_license",    index=False)
    m_not_in_lic_df.to_excel(writer, sheet_name="matched_not_in_license",index=False)
    in_lic_df.to_excel(      writer, sheet_name="in_license_board",      index=False)
    not_in_lic_df.to_excel(  writer, sheet_name="not_in_license",        index=False)
    print(f"  summary              : written")
    if _auto_classify:
        print(f"  merged_addresses     : {len(_merged_df)} providers with multiple rows")
    else:
        print(f"  duplicates_removed   : {len(_dups_df)}")
    print(f"  r3_addr_same         : {len(r3_same)}")
    print(f"  r3_addr_possible     : {len(r3_possible)}  ← review manually")
    print(f"  r3_addr_different    : {len(r3_diff)}")
    print(f"  multi_addr_same      : {len(multi_same)}")
    print(f"  multi_addr_possible  : {len(multi_possible)}  ← review manually")
    print(f"  multi_addr_different : {len(multi_diff)}  ← found in multi file but address differs")
    print(f"  multi_addr_not_found : {len(multi_notfound)}  ← not in multiple_addresses file at all")
    print(f"  matched_in_license   : {len(m_in_lic_df)}  (specialty: {n_m_spec_present} listed, {n_m_spec_missing} blank)")
    print(f"  matched_not_in_lic   : {len(m_not_in_lic_df)}")
    print(f"  in_license_board     : {len(in_lic_df)}  (specialty: {n_spec_present} listed, {n_spec_missing} blank)")
    print(f"  not_in_license       : {len(not_in_lic_df)}")

print(f"""
Row count summary:
  INPUT:
    {len(matched)} rows in 'matched' sheet
    {len(not_matched)} rows in 'not_in_master' sheet
    {len(matched) + len(not_matched)} total

  ADDRESS CHECK (matched providers vs Round 3):
    {len(r3_same)} same  |  {len(r3_possible)} possible  |  {len(r3_diff)} different/not_found
    Total = {n_matched_provs}  {_ok(n_matched_provs, len(matched))}

  ADDRESS CHECK 2 (r3 non-same subset vs multiple_addresses.csv):
    {len(multi_same)} same  |  {len(multi_possible)} possible  |  {len(multi_diff)} different  |  {len(multi_notfound)} not in file
    Checked {n_multi_checked} of {len(r3_possible) + len(r3_diff)} r3 non-same providers
    NOTE: multi_* sheets are a SUBSET of r3 non-same — do not add to r3 totals

  LICENSE CHECK (matched vs MN license board):
    {len(m_in_lic_df)} found  |  {len(m_not_in_lic_df)} not found
    Specialty: {n_m_spec_present} listed, {n_m_spec_missing} blank

  LICENSE CHECK (not_in_master vs MN license board):
    {len(in_lic_df)} found  |  {len(not_in_lic_df)} not found
    Total = {n_unmatched_provs}  {_ok(n_unmatched_provs, len(not_matched))}
    Specialty: {n_spec_present} listed, {n_spec_missing} blank

  GRAND TOTAL: {n_matched_provs} matched + {n_unmatched_provs} not-matched = {n_grand_total}
""")

print("Done.")

# ── DEBUG: Baker name matching ────────────────────────────────────────────────────────────────────────
print("\n── DEBUG: Baker ────────────────────────────────────────────────────────────────────────")
print(f"  License file columns detected: last={l_last!r}  first={l_first!r}")
print(f"  License keys containing 'baker': {_debug_baker_lic_keys}")
if not _debug_baker_lic_rows.empty:
    print(f"  Raw baker rows in license file (detected columns):")
    for _, r in _debug_baker_lic_rows.iterrows():
        print(f"    {l_last}={r[l_last]!r}  {l_first}={r[l_first]!r}")
else:
    print(f"  'baker' NOT found in detected columns {l_last!r}/{l_first!r}")

# Search every column in the license file for 'baker'
print(f"  Searching ALL license file columns for 'baker':")
for col in lic_df.columns:
    hits = lic_df[lic_df[col].astype(str).str.lower().str.contains("baker", na=False)]
    if not hits.empty:
        print(f"    Column {col!r}: {len(hits)} rows — sample: {hits[col].iloc[0]!r}")

print(f"  Keys generated for Baker in not_in_master:")
for keys in _debug_baker_nm_keys:
    print(f"    {keys}")
