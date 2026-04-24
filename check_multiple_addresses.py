#!/usr/bin/env python3
"""
check_multiple_addresses.py

For each person in the 241498 mailing file, find all addresses they have
across the two original mailing list files and classify:

  unique              — only one address found
  same_address_variant — multiple entries but addresses normalize to the same
  multiple_addresses   — genuinely different addresses (multiple locations)
  not_in_originals    — name not found in either original file

Usage:
    python3 check_multiple_addresses.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project"
MAIL_FILE   = BASE / "Current Mailing Files" / "241498 0976 042026 v3.xlsx"
ORIG_FILES  = [
    BASE / "Current Mailing Files" / "Mailing List 2 for Printing Services_titled.xlsx",
    BASE / "Current Mailing Files" / "Mailing List for Printing Services copy_titled.xlsx",
]
OUTPUT_FILE = BASE / "Current Mailing Files" / "multiple_address_check.xlsx"

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

# Address abbreviation normalization
ADDR_ABBREVS = [
    (r"\bstreet\b",    "st"),
    (r"\bavenue\b",    "ave"),
    (r"\bboulevard\b", "blvd"),
    (r"\bdrive\b",     "dr"),
    (r"\bcourt\b",     "ct"),
    (r"\bcircle\b",    "cir"),
    (r"\blane\b",      "ln"),
    (r"\broad\b",      "rd"),
    (r"\bplace\b",     "pl"),
    (r"\bsuite\b",     "ste"),
    (r"\bnorth\b",     "n"),
    (r"\bsouth\b",     "s"),
    (r"\beast\b",      "e"),
    (r"\bwest\b",      "w"),
    (r"[,#\.]",        " "),   # strip punctuation
]

def norm_addr(s) -> str:
    """Normalize address for similarity comparison."""
    s = norm(s)
    for pattern, repl in ADDR_ABBREVS:
        s = re.sub(pattern, repl, s)
    return re.sub(r"\s+", " ", s).strip()

def name_variants(last: str, first: str) -> set[str]:
    last  = STRIP_SUFFIXES.sub("", norm(last)).strip()
    first = STRIP_SUFFIXES.sub("", norm(first)).strip()
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
    if len(words) >= 3:
        variants.add(f"{words[0]} {words[-1]}")
        variants.add(f"{words[-1]} {words[0]}")
        variants.add(f"{words[-1]}, {words[0]}")
    return variants

def detect_cols(df, label):
    cols = {
        "fullname": find_col(df, ["FULL NAME", "Full Name", "full_name"]),
        "last":     find_col(df, ["last_name", "LastName"]),
        "first":    find_col(df, ["first_name", "FirstName", "First Name"]),
        "addr":     find_col(df, ["Delivery Address", "primary_address_1",
                                   "npi_primary_address_1", "address_line1",
                                   "Alternate 1 Address"]),
        "city":     find_col(df, ["City", "primary_city", "npi_primary_city", "city"]),
        "zip":      find_col(df, ["ZIP+4", "zip5", "npi_primary_zip", "zip", "Zip"]),
    }
    print(f"  [{label}] fullname={cols['fullname']} last={cols['last']} "
          f"first={cols['first']} addr={cols['addr']}")
    return cols

def get_name_keys(row, cols) -> set[str]:
    if cols["fullname"] and row.get(cols["fullname"], ""):
        return fullname_variants(norm(row[cols["fullname"]]))
    last  = row.get(cols["last"],  "") if cols["last"]  else ""
    first = row.get(cols["first"], "") if cols["first"] else ""
    return name_variants(last, first)

def get_addr_key(row, cols) -> str:
    addr = norm(row[cols["addr"]]) if cols["addr"] else ""
    city = norm(row[cols["city"]]) if cols["city"] else ""
    z    = zip5(row[cols["zip"]])  if cols["zip"]  else ""
    return f"{addr}|{city}|{z}"

def get_norm_addr_key(row, cols) -> str:
    addr = norm_addr(row[cols["addr"]] if cols["addr"] else "")
    city = norm(row[cols["city"]]) if cols["city"] else ""
    z    = zip5(row[cols["zip"]])  if cols["zip"]  else ""
    return f"{addr}|{city}|{z}"

# ── Load files ────────────────────────────────────────────────────────────────
print(f"\nLoading main mailing file: {MAIL_FILE.name}")
try:
    mail_df = pd.read_excel(MAIL_FILE, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {MAIL_FILE} not found")
print(f"  {len(mail_df)} rows | columns: {list(mail_df.columns)}")
mail_cols = detect_cols(mail_df, "main")

# Build lookup from original files: name_key → list of (raw_addr, norm_addr, source_file)
print(f"\nLoading original mailing files ...")
name_to_addresses: dict[str, list[tuple]] = {}

for orig_path in ORIG_FILES:
    print(f"  {orig_path.name}")
    try:
        df = pd.read_excel(orig_path, dtype=str)
    except FileNotFoundError:
        print(f"    WARNING: not found, skipping")
        continue
    print(f"    {len(df)} rows | columns: {list(df.columns)}")
    cols = detect_cols(df, orig_path.stem[:20])

    for _, row in df.iterrows():
        keys = get_name_keys(row, cols)
        raw_addr  = get_addr_key(row, cols)
        norm_addr_key = get_norm_addr_key(row, cols)
        for key in keys:
            if key not in name_to_addresses:
                name_to_addresses[key] = []
            name_to_addresses[key].append((raw_addr, norm_addr_key, orig_path.name))

print(f"\n  Combined name lookup: {len(name_to_addresses)} unique name keys")

# ── Classify each row in main mailing file ────────────────────────────────────
print(f"\nClassifying {len(mail_df)} rows in main mailing file ...")

statuses       = []
addr_counts    = []
addresses_found = []

for _, row in mail_df.iterrows():
    keys = get_name_keys(row, mail_cols)

    # Collect all address entries for this person across original files
    all_entries = []
    for key in keys:
        if key in name_to_addresses:
            all_entries.extend(name_to_addresses[key])

    if not all_entries:
        statuses.append("not_in_originals")
        addr_counts.append(0)
        addresses_found.append("")
        continue

    # Deduplicate entries by raw address
    seen_raw  = {}
    for raw, norm_a, src in all_entries:
        seen_raw.setdefault(raw, (norm_a, src))

    unique_raw   = list(seen_raw.keys())
    unique_norms = [seen_raw[r][0] for r in unique_raw]
    addr_counts.append(len(unique_raw))
    addresses_found.append(" | ".join(unique_raw))

    if len(unique_raw) == 1:
        statuses.append("unique")
    else:
        # Check if all normalized addresses are the same
        if len(set(unique_norms)) == 1:
            statuses.append("same_address_variant")
        else:
            statuses.append("multiple_addresses")

mail_df["address_status"] = statuses
mail_df["address_count"]  = addr_counts
mail_df["addresses_found"] = addresses_found

# ── Summary ───────────────────────────────────────────────────────────────────
status_counts = pd.Series(statuses).value_counts()
print(f"\nResults:")
for status, cnt in status_counts.items():
    print(f"  {status}: {cnt}")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE}")

multi   = mail_df[mail_df["address_status"] == "multiple_addresses"]
variant = mail_df[mail_df["address_status"] == "same_address_variant"]
not_in  = mail_df[mail_df["address_status"] == "not_in_originals"]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    mail_df.to_excel(writer, sheet_name="all_flagged", index=False)
    multi.to_excel(writer, sheet_name="multiple_addresses", index=False)
    variant.to_excel(writer, sheet_name="same_address_variant", index=False)
    not_in.to_excel(writer, sheet_name="not_in_originals", index=False)
    print(f"  all_flagged          : {len(mail_df)}")
    print(f"  multiple_addresses   : {len(multi)}")
    print(f"  same_address_variant : {len(variant)}")
    print(f"  not_in_originals     : {len(not_in)}")

print("\nDone.")
