#!/usr/bin/env python3
# ── Self-update ───────────────────────────────────────────────────────────────
import os, sys, urllib.request, ssl
_URL = ("https://raw.githubusercontent.com/landrewsster/"
        "starting-ragchatbot-codebase/claude/clean-sort-merge-npi-GngZZ/"
        "email_list.py")
try:
    _ctx = ssl._create_unverified_context()
    _new = urllib.request.urlopen(_URL, timeout=10, context=_ctx).read()
    _old = open(__file__, "rb").read()
    if _new != _old:
        open(__file__, "wb").write(_new)
        print("Updated email_list.py — restarting ...")
        os.execv(sys.executable, [sys.executable, __file__] + sys.argv[1:])
except Exception as _e:
    print(f"(self-update skipped: {_e})")
# ─────────────────────────────────────────────────────────────────────────────

"""
email_list.py

Combines all unique providers from:
  - Round 3 mailing list
  - Part 1 crossref (matched + not_in_master)
  - Part 2 crossref (matched + not_in_master)

Deduplicates across all sources, looks each provider up in the
MN Physician and PA list March 2026, and outputs those with an
email address (col L).

Output: email_list.xlsx
  with_email — providers found in license board who have an email
  no_email   — providers found in license board with no email

Usage:
    python3 email_list.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE           = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
ROUND3_FILE    = BASE / "MailingList_Round3_20260519.xlsx"
CROSSREF1_FILE = BASE / "provider_round3_crossref.xlsx"
CROSSREF2_FILE = BASE / "provider_round3_crossref_ManualSearch_Part2_05212026.xlsx"
LICENSE_FILE   = BASE / "MN State Licensing Board" / "MN Physician and PA list March 2026.xlsx"
OUTPUT_FILE    = BASE / "email_list.xlsx"

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
    return re.sub(r"[,.]", "", s).strip()

def find_col(df, candidates):
    low      = {c.lower(): c for c in df.columns}
    low_norm = {c.lower().replace(" ", "_"): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
        if c.lower().replace(" ", "_") in low_norm:
            return low_norm[c.lower().replace(" ", "_")]
    return None

def lf_keys(last: str, first: str) -> set[str]:
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    fi = f1[0] if f1 else ""
    keys = set()
    if l and f:
        keys.add(f"{l}|{f}")
    if l and f1 and f1 != f:
        keys.add(f"{l}|{f1}")
    if l and fi:
        keys.add(f"{l}|~{fi}")
    return keys

def make_all_keys(col_a: str, col_b: str) -> set[str]:
    b = clean_name(col_b)
    if " " in b:
        words = b.split()
        first = words[0]
        keys  = lf_keys(words[-1], first)
        if len(words) >= 3:
            keys |= lf_keys(" ".join(words[-2:]), first)
        return keys
    return lf_keys(col_a, col_b)

def make_key(last: str, first: str) -> str:
    l  = clean_name(last)
    f  = clean_name(first)
    f1 = f.split()[0] if f else ""
    return f"{l}|{f1}"

# ── Load sources ──────────────────────────────────────────────────────────────
seen_keys     = set()
all_providers = []

def _add_providers(df, last_col, first_col, source_label):
    if df is None or df.empty or not last_col or not first_col:
        return 0
    added = 0
    for _, row in df.iterrows():
        key = make_key(norm(row.get(last_col, "")), norm(row.get(first_col, "")))
        if key and key != "|" and key not in seen_keys:
            seen_keys.add(key)
            all_providers.append({
                "last_name":  row.get(last_col, ""),
                "first_name": row.get(first_col, ""),
                "source":     source_label,
                "_key":       key,
            })
            added += 1
    return added

# 1. Round 3 mailing list
print(f"\nLoading Round 3: {ROUND3_FILE.name}")
try:
    r3 = pd.read_excel(ROUND3_FILE, dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {ROUND3_FILE} not found")
r3_last  = find_col(r3, ["last_name", "LastName", "Last Name"]) or r3.columns[0]
r3_first = find_col(r3, ["first_name", "FirstName", "First Name"]) or r3.columns[1]
n = _add_providers(r3, r3_last, r3_first, "Round3")
print(f"  {len(r3)} rows — added {n} unique providers")

# 2 & 3. Both crossref files — only not_in_master providers.
# Matched providers are already in Round 3 by definition; loading them
# separately would cause duplicates because the manual file may spell
# the same name differently (e.g. "Liz" vs "Elizabeth").
for xref_path, part_label in [
    (CROSSREF1_FILE, "Part1"),
    (CROSSREF2_FILE, "Part2"),
]:
    print(f"\nLoading {part_label} crossref: {xref_path.name}")
    if not xref_path.exists():
        print(f"  WARNING: {xref_path.name} not found — skipping")
        continue
    xl = pd.ExcelFile(xref_path)
    if "not_in_master" not in xl.sheet_names:
        print(f"  WARNING: 'not_in_master' sheet not found — skipping")
        continue
    df = xl.parse("not_in_master", dtype=str).fillna("")
    last_col  = find_col(df, ["Last_Name", "last_name", "LastName", "Last Name"])
    first_col = find_col(df, ["First_Name", "first_name", "FirstName", "First Name"])
    n = _add_providers(df, last_col, first_col, f"{part_label}_not_in_master")
    print(f"  not_in_master: {len(df)} rows — {n} new unique providers added")

combined = pd.DataFrame(all_providers)
print(f"\nTotal unique providers across all sources: {len(combined)}")

# ── Load MN license board ─────────────────────────────────────────────────────
print(f"\nLoading MN license board: {LICENSE_FILE.name}")
try:
    lic_xl = pd.ExcelFile(LICENSE_FILE)
    lic_df = lic_xl.parse(lic_xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {LICENSE_FILE} not found")
print(f"  {len(lic_df)} rows | columns: {list(lic_df.columns)}")

l_last      = find_col(lic_df, ["last_name", "LastName", "Last Name"]) or lic_df.columns[0]
l_first     = find_col(lic_df, ["first_name", "FirstName", "First Name"]) or lic_df.columns[1]
l_specialty = lic_df.columns[9]  if len(lic_df.columns) > 9  else None
l_cert      = lic_df.columns[10] if len(lic_df.columns) > 10 else None
l_email     = lic_df.columns[11] if len(lic_df.columns) > 11 else None
print(f"  last={l_last!r}  first={l_first!r}")
print(f"  specialty col (J): {l_specialty!r}")
print(f"  cert col      (K): {l_cert!r}")
print(f"  email col     (L): {l_email!r}")

# Build license lookup
lic_key_to_idx: dict[str, int] = {}
for idx, row in lic_df.iterrows():
    for key in make_all_keys(row[l_last], row[l_first]):
        lic_key_to_idx.setdefault(key, idx)
print(f"  Unique name keys: {len(lic_key_to_idx)}")

# ── Look up each provider in license board ────────────────────────────────────
print(f"\nLooking up {len(combined)} providers in MN license board ...")

with_email = []
no_email   = []

for _, prov in combined.iterrows():
    keys = lf_keys(prov["last_name"], prov["first_name"])
    exact_keys   = {k for k in keys if "|~" not in k}
    initial_keys = {k for k in keys if "|~" in k}
    found_idx = next((lic_key_to_idx[k] for k in exact_keys   if k in lic_key_to_idx), None)
    if found_idx is None:
        found_idx = next((lic_key_to_idx[k] for k in initial_keys if k in lic_key_to_idx), None)

    if found_idx is not None:
        lic_row   = lic_df.iloc[found_idx]
        lic_name  = f"{lic_row[l_last]}, {lic_row[l_first]}".strip(", ")
        specialty = lic_row[l_specialty] if l_specialty else ""
        cert      = lic_row[l_cert]      if l_cert      else ""
        email     = lic_row[l_email]     if l_email     else ""
        out = {
            "last_name":                prov["last_name"],
            "first_name":               prov["first_name"],
            "source":                   prov["source"],
            "license_name_match":       lic_name,
            "specialty_boards":         specialty,
            "certification":            cert,
            "email":                    email,
            "specialty_boards_missing": "missing" if not norm(specialty) else "",
            "certification_missing":    "missing" if not norm(cert)      else "",
        }
        if norm(email):
            with_email.append(out)
        else:
            no_email.append(out)
    # Providers not found in the license board are omitted from output

with_email_df = pd.DataFrame(with_email).reset_index(drop=True)
no_email_df   = pd.DataFrame(no_email).reset_index(drop=True)

print(f"  Found in license board with email   : {len(with_email_df)}")
print(f"  Found in license board, no email    : {len(no_email_df)}")
print(f"  Not found in license board          : {len(combined) - len(with_email_df) - len(no_email_df)}")

_clean_cols = ["last_name", "first_name", "email", "specialty_boards", "certification",
               "specialty_boards_missing", "certification_missing"]

with_email_clean = (
    with_email_df[_clean_cols]
    .sort_values(["last_name", "first_name"], key=lambda s: s.str.lower())
    .reset_index(drop=True)
)

no_email_clean = (
    no_email_df[_clean_cols]
    .sort_values(["last_name", "first_name"], key=lambda s: s.str.lower())
    .reset_index(drop=True)
)

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    with_email_df.to_excel(   writer, sheet_name="with_email",             index=False)
    with_email_clean.to_excel(writer, sheet_name="with_email_noduplicates",index=False)
    no_email_df.to_excel(     writer, sheet_name="no_email",               index=False)
    no_email_clean.to_excel(  writer, sheet_name="no_email_noduplicates",  index=False)
    print(f"  with_email             : {len(with_email_df)} providers (with source & match detail)")
    print(f"  with_email_noduplicates: {len(with_email_clean)} providers (sorted, clean)")
    print(f"  no_email               : {len(no_email_df)} providers (with source & match detail)")
    print(f"  no_email_noduplicates  : {len(no_email_clean)} providers (sorted, clean)")

print("\nDone.")
