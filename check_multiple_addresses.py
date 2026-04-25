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
    BASE / "Current Mailing Files" / "gold_reference_providers.xlsx",
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
        # All address column sets — each is (addr, city, zip)
        "addr_sets": []
    }
    # Primary address
    addr = find_col(df, ["npi_primary_address_1", "primary_address_1",
                          "Delivery Address", "address_line1", "Alternate 1 Address"])
    city = find_col(df, ["npi_primary_city", "primary_city", "City", "city"])
    z    = find_col(df, ["npi_primary_zip", "zip5", "ZIP+4", "zip", "Zip"])
    if addr:
        cols["addr_sets"].append((addr, city, z))

    # Mailing address (separate from primary)
    m_addr = find_col(df, ["npi_mailing_address_1", "mailing_address_1"])
    m_city = find_col(df, ["npi_mailing_city", "mailing_city"])
    m_zip  = find_col(df, ["npi_mailing_zip", "mailing_zip", "mailing_postal_code"])
    if m_addr and m_addr != addr:
        cols["addr_sets"].append((m_addr, m_city, m_zip))

    # MN-list address (matched tab wide output)
    mn_addr = find_col(df, ["mn_address_1"])
    mn_city = find_col(df, ["mn_city"])
    mn_zip  = find_col(df, ["mn_zip"])
    if mn_addr and mn_addr != addr:
        cols["addr_sets"].append((mn_addr, mn_city, mn_zip))

    # Use first address set as default for backward compat
    cols["addr"] = cols["addr_sets"][0][0] if cols["addr_sets"] else None
    cols["city"] = cols["addr_sets"][0][1] if cols["addr_sets"] else None
    cols["zip"]  = cols["addr_sets"][0][2] if cols["addr_sets"] else None

    print(f"  [{label}] last={cols['last']} first={cols['first']} "
          f"addr_sets={len(cols['addr_sets'])}: {[a[0] for a in cols['addr_sets']]}")
    return cols

def get_name_keys(row, cols) -> set[str]:
    if cols["fullname"] and row.get(cols["fullname"], ""):
        return fullname_variants(norm(row[cols["fullname"]]))
    last  = row.get(cols["last"],  "") if cols["last"]  else ""
    first = row.get(cols["first"], "") if cols["first"] else ""
    return name_variants(last, first)

def get_all_addr_keys(row, cols) -> list[tuple[str, str]]:
    """Return list of (raw_key, norm_key) for every address column set in the row."""
    results = []
    for addr_col, city_col, zip_col in cols["addr_sets"]:
        addr = norm(row[addr_col]) if addr_col and addr_col in row.index else ""
        city = norm(row[city_col]) if city_col and city_col in row.index else ""
        z    = zip5(row[zip_col])  if zip_col  and zip_col  in row.index else ""
        if addr:
            raw  = f"{addr}|{city}|{z}"
            na   = norm_addr(row[addr_col] if addr_col else "")
            norm_key = f"{na}|{city}|{z}"
            results.append((raw, norm_key))
    return results

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
        sheets = pd.read_excel(orig_path, sheet_name=None, dtype=str)
    except FileNotFoundError:
        print(f"    WARNING: not found, skipping")
        continue
    for sheet_name, df in sheets.items():
        print(f"    Sheet '{sheet_name}': {len(df)} rows")
        cols = detect_cols(df, f"{orig_path.stem[:15]}/{sheet_name}")

        npi_col    = find_col(df, ["npi", "NPI"])
        middle_col = find_col(df, ["middle_name", "MiddleName"])

        for _, row in df.iterrows():
            name_keys  = get_name_keys(row, cols)
            addr_pairs = get_all_addr_keys(row, cols)
            npi    = norm(row[npi_col])    if npi_col    else ""
            middle = norm(row[middle_col]) if middle_col else ""
            middle1 = middle[0] if middle else ""   # first initial only
            for name_key in name_keys:
                if name_key not in name_to_addresses:
                    name_to_addresses[name_key] = []
                for raw_addr, norm_addr_key in addr_pairs:
                    # Store npi + middle initial alongside address for disambiguation
                    name_to_addresses[name_key].append(
                        (raw_addr, norm_addr_key, orig_path.name, npi, middle1)
                    )

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

    # Check if multiple distinct NPIs exist — means different people, not multiple addresses
    npis    = {e[3] for e in all_entries if e[3]}
    middles = {e[4] for e in all_entries if e[4]}
    different_people = (
        len(npis) > 1 or                          # different NPI numbers
        (len(middles) > 1 and "" not in middles)  # different non-blank middle initials
    )

    if different_people:
        statuses.append("name_collision_different_people")
        addr_counts.append(len(npis) or len(middles))
        npis_str = " | ".join(sorted(npis)) if npis else "no NPI"
        addresses_found.append(f"NPIs: {npis_str}")
        continue

    # Deduplicate entries by raw address
    seen_raw = {}
    for raw, norm_a, src, npi, mid in all_entries:
        seen_raw.setdefault(raw, norm_a)

    unique_raw   = list(seen_raw.keys())
    unique_norms = list(seen_raw.values())
    addr_counts.append(len(unique_raw))
    addresses_found.append(" | ".join(unique_raw))

    if len(unique_raw) == 1:
        statuses.append("unique")
    elif len(set(unique_norms)) == 1:
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
not_in     = mail_df[mail_df["address_status"] == "not_in_originals"]
collisions = mail_df[mail_df["address_status"] == "name_collision_different_people"]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    mail_df.to_excel(writer, sheet_name="all_flagged", index=False)
    multi.to_excel(writer, sheet_name="multiple_addresses", index=False)
    variant.to_excel(writer, sheet_name="same_address_variant", index=False)
    collisions.to_excel(writer, sheet_name="name_collision", index=False)
    not_in.to_excel(writer, sheet_name="not_in_originals", index=False)
    print(f"  all_flagged                    : {len(mail_df)}")
    print(f"  multiple_addresses             : {len(multi)}")
    print(f"  same_address_variant           : {len(variant)}")
    print(f"  name_collision_diff_people     : {len(collisions)}")
    print(f"  not_in_originals               : {len(not_in)}")

print("\nDone.")
