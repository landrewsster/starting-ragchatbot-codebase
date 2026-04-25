#!/usr/bin/env python3
"""
check_gold_vs_mailing.py

Verify the gold reference file is complete and duplicate-free relative to
both mailing lists, and surface the ~384 providers who appear in the gold
reference but were not included in either mailing.

Checks performed:
  1. Duplicate NPI in gold reference (should be 0)
  2. Gold reference providers NOT in any mailing (~384 expected)
  3. Mailing list providers NOT found in gold reference (should be 0)

Matching strategy:
  - NPI match first (MailingListAddition has NPI; 241498 main file does not)
  - Name match as fallback for both mailing files

Usage:
    python3 check_gold_vs_mailing.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE          = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
GOLD_FILE     = BASE / "gold_reference_providers.xlsx"
MAIL_MAIN     = BASE / "241498 0976 042026 v3.xlsx"
MAIL_ADDITION = BASE / "MailingListAddition20260423.xlsx"
OUTPUT_FILE   = BASE / "gold_vs_mailing_check.xlsx"

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def zip5(s) -> str:
    return re.sub(r"\D", "", norm(s))[:5]

def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$"
)

def clean_name(s) -> str:
    s = STRIP_SUFFIXES.sub("", norm(s)).strip()
    return re.sub(r"[,.]", "", s).strip()

def name_variants(last: str, first: str) -> set[str]:
    last  = clean_name(last)
    first = clean_name(first)
    first1 = first.split()[0] if first else ""
    cands = set()
    for f in {first, first1} - {""}:
        if last:
            cands.add(f"{f} {last}")
            cands.add(f"{last} {f}")
            cands.add(f"{last}, {f}")
        else:
            cands.add(f)
    if not cands and last:
        cands.add(last)
    return cands

def fullname_variants(fn: str) -> set[str]:
    fn = fn.strip()
    if not fn:
        return set()
    variants = {fn}
    stripped = STRIP_SUFFIXES.sub("", fn).strip()
    if stripped and stripped != fn:
        variants.add(stripped)
    words = (stripped or fn).split()
    if len(words) >= 2:
        variants.add(f"{words[0]} {words[-1]}")
        variants.add(f"{words[-1]} {words[0]}")
        variants.add(f"{words[-1]}, {words[0]}")
    return variants

# ── Load gold reference ───────────────────────────────────────────────────────
print(f"\nLoading gold reference: {GOLD_FILE.name}")
try:
    gold = pd.read_excel(GOLD_FILE, dtype=str).fillna("")
except FileNotFoundError:
    sys.exit(f"ERROR: {GOLD_FILE} not found — run build_gold_reference.py first")

print(f"  {len(gold)} rows | {len(gold.columns)} columns")

gold_npi_col   = find_col(gold, ["npi", "NPI"])
gold_last_col  = find_col(gold, ["last_name", "LastName"])
gold_first_col = find_col(gold, ["first_name", "FirstName"])
gold_src_col   = find_col(gold, ["_source_tab"])

print(f"  Cols: npi={gold_npi_col} last={gold_last_col} first={gold_first_col} "
      f"source_tab={gold_src_col}")

# ── Check 1: Duplicate NPIs in gold reference ─────────────────────────────────
print(f"\nCheck 1: Duplicate NPI in gold reference ...")
gold_npis = gold[gold_npi_col].apply(norm) if gold_npi_col else pd.Series([""] * len(gold))
dup_mask  = gold_npis.duplicated(keep=False) & (gold_npis != "")
dup_gold  = gold[dup_mask].copy()
print(f"  Duplicate NPI rows: {len(dup_gold)}")
if len(dup_gold):
    print(dup_gold[[gold_npi_col, gold_last_col, gold_first_col]].to_string())

# ── Build mailed-NPI and mailed-name sets ─────────────────────────────────────
mailed_npis:  set[str] = set()
mailed_names: set[str] = set()

# ── Load MailingListAddition ──────────────────────────────────────────────────
print(f"\nLoading addition mailing file: {MAIL_ADDITION.name}")
try:
    add_sheets = pd.read_excel(MAIL_ADDITION, sheet_name=None, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {MAIL_ADDITION} not found")

add_rows_all = []
for sheet_name, df in add_sheets.items():
    df = df.fillna("")
    npi_col   = find_col(df, ["npi", "NPI"])
    last_col  = find_col(df, ["last_name", "LastName"])
    first_col = find_col(df, ["first_name", "FirstName", "First Name"])
    print(f"  Sheet '{sheet_name}': {len(df)} rows  "
          f"(npi={npi_col} last={last_col} first={first_col})")
    for _, row in df.iterrows():
        npi   = norm(row[npi_col])   if npi_col   else ""
        last  = norm(row[last_col])  if last_col  else ""
        first = norm(row[first_col]) if first_col else ""
        if npi:
            mailed_npis.add(npi)
        for v in name_variants(last, first):
            mailed_names.add(v)
    add_rows_all.append(df)

add_df = pd.concat(add_rows_all, ignore_index=True)
print(f"  Total addition rows: {len(add_df)}")

# ── Load main 241498 mailing file ─────────────────────────────────────────────
print(f"\nLoading main mailing file: {MAIL_MAIN.name}")
try:
    main_df = pd.read_excel(MAIL_MAIN, dtype=str).fillna("")
except FileNotFoundError:
    sys.exit(f"ERROR: {MAIL_MAIN} not found")

fn_col    = find_col(main_df, ["FULL NAME", "Full Name", "full_name"])
fname_col = find_col(main_df, ["First Name", "first_name", "FirstName"])
lname_col = find_col(main_df, ["last_name", "LastName"])
print(f"  {len(main_df)} rows  (fullname={fn_col} first={fname_col} last={lname_col})")

for _, row in main_df.iterrows():
    if fn_col and row.get(fn_col, "").strip():
        for v in fullname_variants(norm(row[fn_col])):
            mailed_names.add(clean_name(v))
    elif lname_col and fname_col:
        for v in name_variants(row.get(lname_col, ""), row.get(fname_col, "")):
            mailed_names.add(v)

print(f"\n  Combined mailed NPIs : {len(mailed_npis)}")
print(f"  Combined mailed names: {len(mailed_names)}")

# ── Check 2: Gold reference providers NOT in any mailing ─────────────────────
print(f"\nCheck 2: Gold providers not in any mailing ...")
in_mailing_flags  = []
match_method_list = []

for _, row in gold.iterrows():
    npi   = norm(row[gold_npi_col])   if gold_npi_col   else ""
    last  = norm(row[gold_last_col])  if gold_last_col  else ""
    first = norm(row[gold_first_col]) if gold_first_col else ""

    if npi and npi in mailed_npis:
        in_mailing_flags.append(True)
        match_method_list.append("npi")
    else:
        matched_name = False
        for v in name_variants(last, first):
            if v in mailed_names:
                matched_name = True
                break
        if matched_name:
            in_mailing_flags.append(True)
            match_method_list.append("name")
        else:
            in_mailing_flags.append(False)
            match_method_list.append("not_mailed")

gold["in_any_mailing"] = in_mailing_flags
gold["mailing_match"]  = match_method_list

not_mailed  = gold[gold["in_any_mailing"] == False].copy()
in_mailed   = gold[gold["in_any_mailing"] == True].copy()
print(f"  In mailing   : {len(in_mailed)}")
print(f"  NOT in mailing: {len(not_mailed)}")

if gold_src_col:
    print(f"\n  Not-mailed breakdown by source tab:")
    print(not_mailed[gold_src_col].value_counts().to_string())

# ── Check 3: Mailing providers NOT in gold reference ─────────────────────────
print(f"\nCheck 3: Mailing providers not found in gold reference ...")

# Build gold lookup
gold_npi_set  = set()
gold_name_set = set()
for _, row in gold.iterrows():
    npi   = norm(row[gold_npi_col])   if gold_npi_col   else ""
    last  = norm(row[gold_last_col])  if gold_last_col  else ""
    first = norm(row[gold_first_col]) if gold_first_col else ""
    if npi:
        gold_npi_set.add(npi)
    for v in name_variants(last, first):
        gold_name_set.add(v)

# Check main mailing
main_not_in_gold = []
for _, row in main_df.iterrows():
    fns = []
    if fn_col and row.get(fn_col, "").strip():
        fns = [clean_name(v) for v in fullname_variants(norm(row[fn_col]))]
    elif lname_col and fname_col:
        fns = list(name_variants(row.get(lname_col, ""), row.get(fname_col, "")))
    found = any(v in gold_name_set for v in fns)
    if not found:
        main_not_in_gold.append(row)

main_missing_df = pd.DataFrame(main_not_in_gold)
print(f"  Main mailing not in gold  : {len(main_missing_df)}")

# Check addition
add_not_in_gold = []
for _, row in add_df.iterrows():
    add_npi_col   = find_col(add_df, ["npi", "NPI"])
    add_last_col  = find_col(add_df, ["last_name", "LastName"])
    add_first_col = find_col(add_df, ["first_name", "FirstName", "First Name"])
    npi   = norm(row[add_npi_col])   if add_npi_col   else ""
    last  = norm(row[add_last_col])  if add_last_col  else ""
    first = norm(row[add_first_col]) if add_first_col else ""
    if npi and npi in gold_npi_set:
        continue
    if any(v in gold_name_set for v in name_variants(last, first)):
        continue
    add_not_in_gold.append(row)

add_missing_df = pd.DataFrame(add_not_in_gold)
print(f"  Addition mailing not in gold: {len(add_missing_df)}")

# ── Summary ───────────────────────────────────────────────────────────────────
print(f"\n{'='*60}")
print(f"SUMMARY")
print(f"{'='*60}")
print(f"  Gold reference total         : {len(gold)}")
print(f"  In main mailing (241498)     : {len(main_df)}")
print(f"  In addition mailing          : {len(add_df)}")
print(f"  Combined unique mailed (est) : {len(main_df) + len(add_df)}")
print(f"  Duplicate NPI in gold        : {len(dup_gold)}")
print(f"  Gold providers NOT mailed    : {len(not_mailed)}")
print(f"  Main mailing NOT in gold     : {len(main_missing_df)}")
print(f"  Addition NOT in gold         : {len(add_missing_df)}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE}")

# Sort not_mailed by source tab then last name for easier review
sort_cols = []
if gold_src_col:
    sort_cols.append(gold_src_col)
if gold_last_col:
    sort_cols.append(gold_last_col)
if sort_cols:
    not_mailed = not_mailed.sort_values(sort_cols, ignore_index=True)

summary_rows = [
    {"check": "Gold reference total",          "count": len(gold)},
    {"check": "In main mailing (241498)",       "count": len(main_df)},
    {"check": "In addition mailing",            "count": len(add_df)},
    {"check": "Combined mailed (main+addition)","count": len(main_df) + len(add_df)},
    {"check": "Duplicate NPI in gold",          "count": len(dup_gold)},
    {"check": "Gold providers NOT mailed",      "count": len(not_mailed)},
    {"check": "Main mailing NOT in gold",       "count": len(main_missing_df)},
    {"check": "Addition NOT in gold",           "count": len(add_missing_df)},
]
summary_df = pd.DataFrame(summary_rows)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    summary_df.to_excel(writer,       sheet_name="summary",                index=False)
    not_mailed.to_excel(writer,       sheet_name="not_in_any_mailing",     index=False)
    gold.to_excel(writer,             sheet_name="gold_all_flagged",        index=False)
    dup_gold.to_excel(writer,         sheet_name="gold_duplicate_npi",      index=False)
    main_missing_df.to_excel(writer,  sheet_name="main_mailing_not_in_gold",index=False)
    add_missing_df.to_excel(writer,   sheet_name="addition_not_in_gold",    index=False)
    print(f"  summary                  : {len(summary_df)} rows")
    print(f"  not_in_any_mailing       : {len(not_mailed)}")
    print(f"  gold_all_flagged         : {len(gold)}")
    print(f"  gold_duplicate_npi       : {len(dup_gold)}")
    print(f"  main_mailing_not_in_gold : {len(main_missing_df)}")
    print(f"  addition_not_in_gold     : {len(add_missing_df)}")

print("\nDone.")
