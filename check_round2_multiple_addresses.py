#!/usr/bin/env python3
"""
check_round2_multiple_addresses.py

For each person in MailingListRound2.xlsx, find all addresses they have
in gold_reference_providers.xlsx and classify:

  unique                        — only one address found in gold reference
  same_address_variant          — multiple entries but addresses normalize to the same
  multiple_addresses            — genuinely different addresses (multiple locations)
  name_collision_different_people — same name but different NPIs / middle initials
  not_in_gold                   — name/NPI not found in gold reference at all

Round 2 column format (mixed across rows):
  Col A (index 0): last name  OR first name (for ~3500 mailingaddition entries)
  Col B (index 1): first name OR full name  (for ~3500 mailingaddition entries)
  Col C (index 2): middle name
  Col D (index 3): address line 1
  Col E (index 4): city  (or address line 2)
  Col F (index 5): zip   (or city)

Matching strategy: tries col_a-as-last + col_b-as-first AND col_b-as-fullname,
plus NPI when available.

Usage:
    python3 check_round2_multiple_addresses.py
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project"
MAIL_FILE   = BASE / "Current Mailing Files" / "MailingListRound2 copy 2.xlsx"
GOLD_FILE   = BASE / "Current Mailing Files" / "gold_reference_providers.xlsx"
OUTPUT_FILE = BASE / "Current Mailing Files" / "round2_multiple_address_check.xlsx"

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
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
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$",
    re.IGNORECASE,
)

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
    (r"[,#\.]",        " "),
]

def norm_addr(s) -> str:
    s = norm(s)
    for pattern, repl in ADDR_ABBREVS:
        s = re.sub(pattern, repl, s)
    return re.sub(r"\s+", " ", s).strip()

def name_variants(last: str, first: str) -> set[str]:
    """Keys for last + first matching."""
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
    """Keys from a 'First Last' or 'Last, First' full-name string."""
    fn = STRIP_SUFFIXES.sub("", norm(fn)).strip()
    if not fn:
        return set()
    variants = {fn}
    words = fn.split()
    if len(words) >= 2:
        variants.add(f"{words[0]} {words[-1]}")
        variants.add(f"{words[-1]} {words[0]}")
        variants.add(f"{words[-1]}, {words[0]}")
    return variants

def round2_name_keys(col_a_val: str, col_b_val: str) -> set[str]:
    """
    Multi-strategy matching for Round 2's mixed column format.

    Strategy 1: col_a = last name, col_b = first name (standard format for ~5500 rows)
    Strategy 2: col_b = full name (when col_b has a space — mailingaddition format)

    Both strategies are tried and results unioned so we match regardless of format.
    """
    a = norm(col_a_val)
    b = norm(col_b_val)
    keys: set[str] = set()

    # Strategy 1: last + first
    if a or b:
        keys |= name_variants(a, b)

    # Strategy 2: col_b as full name (covers mailingaddition where col_b = "First Last")
    if " " in b:
        keys |= fullname_variants(b)
        # Also try: col_a = first name, extract last from fullname in col_b
        words = b.split()
        if len(words) >= 2:
            keys |= name_variants(words[-1], a)

    return keys

def get_all_addr_keys(row, cols) -> list[tuple[str, str]]:
    """Return (raw_key, norm_key) for every address set in the row."""
    results = []
    for addr_col, city_col, zip_col in cols["addr_sets"]:
        addr = norm(row[addr_col]) if addr_col and addr_col in row.index else ""
        city = norm(row[city_col]) if city_col and city_col in row.index else ""
        z    = zip5(row[zip_col])  if zip_col  and zip_col  in row.index else ""
        if addr:
            raw      = f"{addr}|{city}|{z}"
            norm_key = f"{norm_addr(row[addr_col])}|{city}|{z}"
            results.append((raw, norm_key))
    return results

def detect_gold_cols(df, label):
    cols = {
        "fullname": find_col(df, ["FULL NAME", "Full Name", "full_name"]),
        "last":     find_col(df, ["last_name", "LastName", "Last Name"]),
        "first":    find_col(df, ["first_name", "FirstName", "First Name"]),
        "addr_sets": [],
    }
    if find_col(df, ["_addr1"]):
        cols["addr_sets"].append((find_col(df, ["_addr1"]),
                                  find_col(df, ["_city"]),
                                  find_col(df, ["_zip"])))
        mail = find_col(df, ["_mail_addr"])
        if mail:
            cols["addr_sets"].append((mail,
                                      find_col(df, ["_mail_city"]),
                                      find_col(df, ["_mail_zip"])))
    else:
        addr = find_col(df, ["npi_primary_address_1", "primary_address_1",
                              "Delivery Address", "address_line1", "Alternate 1 Address"])
        city = find_col(df, ["npi_primary_city", "primary_city", "City", "city"])
        z    = find_col(df, ["npi_primary_zip", "zip5", "ZIP+4", "zip", "Zip"])
        if addr:
            cols["addr_sets"].append((addr, city, z))

        m_addr = find_col(df, ["npi_mailing_address_1", "mailing_address_1"])
        m_city = find_col(df, ["npi_mailing_city", "mailing_city"])
        m_zip  = find_col(df, ["npi_mailing_zip", "mailing_zip", "mailing_postal_code"])
        if m_addr and m_addr != addr:
            cols["addr_sets"].append((m_addr, m_city, m_zip))

        mn_addr = find_col(df, ["mn_address_1"])
        mn_city = find_col(df, ["mn_city"])
        mn_zip  = find_col(df, ["mn_zip"])
        if mn_addr and mn_addr != addr:
            cols["addr_sets"].append((mn_addr, mn_city, mn_zip))

    print(f"  [{label}] last={cols['last']} first={cols['first']} "
          f"addr_sets={len(cols['addr_sets'])}: {[a[0] for a in cols['addr_sets']]}")
    return cols

def parse_addresses_found(s: str) -> list[tuple[str, str, str]]:
    if not s:
        return []
    results = []
    for entry in s.split(" | "):
        parts = entry.split("|")
        addr = parts[0].strip() if len(parts) > 0 else ""
        city = parts[1].strip() if len(parts) > 1 else ""
        z    = parts[2].strip() if len(parts) > 2 else ""
        if addr:
            results.append((addr, city, z))
    return results

# ── Load Round 2 mailing file ─────────────────────────────────────────────────
print(f"\nLoading Round 2 mailing file: {MAIL_FILE.name}")
try:
    mail_df = pd.read_excel(MAIL_FILE, dtype=str).fillna("")
except FileNotFoundError:
    sys.exit(f"ERROR: {MAIL_FILE} not found")

print(f"  {len(mail_df)} rows | {len(mail_df.columns)} columns")
print(f"  Columns: {list(mail_df.columns)}")

# Positional column access for Round 2
r2_col_a   = mail_df.columns[0]                                         # last OR first
r2_col_b   = mail_df.columns[1] if len(mail_df.columns) > 1 else None  # first OR fullname
r2_npi_col = find_col(mail_df, ["npi", "NPI"])
# Address columns D, E, F (positions 3-5)
r2_addr_col = mail_df.columns[3] if len(mail_df.columns) > 3 else None
r2_city_col = mail_df.columns[4] if len(mail_df.columns) > 4 else None
r2_zip_col  = mail_df.columns[5] if len(mail_df.columns) > 5 else None

print(f"  col_a={r2_col_a!r}  col_b={r2_col_b!r}  npi={r2_npi_col}")
print(f"  addr={r2_addr_col!r}  city={r2_city_col!r}  zip={r2_zip_col!r}")

# ── Assign source group ───────────────────────────────────────────────────────
# If the file already has a _source column (user pre-labeled goldaddition rows),
# use it.  Otherwise auto-detect: multi-word col B → mailingaddition, else → firstmailing.
# To get three-way labels, add a column named "_source" to MailingListRound2.xlsx
# and mark your ~623 added rows as "goldaddition" before running this script.
_source_col = find_col(mail_df, ["_source", "source"])
if _source_col:
    mail_df["_source"] = mail_df[_source_col].apply(norm)
    # Auto-fill blank mailingaddition rows from col B content
    mask_ma = mail_df["_source"].eq("") & mail_df[r2_col_b].apply(lambda v: " " in str(v).strip())
    mail_df.loc[mask_ma, "_source"] = "mailingaddition"
    # Default any remaining blank rows to firstmailing
    mask_fm = mail_df["_source"].eq("")
    mail_df.loc[mask_fm, "_source"] = "firstmailing"
    print(f"  Using existing '_source' column; auto-filled {mask_ma.sum()} mailingaddition, "
          f"{mask_fm.sum()} firstmailing rows")
else:
    mail_df["_source"] = mail_df[r2_col_b].apply(
        lambda v: "mailingaddition" if " " in str(v).strip() else "firstmailing"
    )
    src_counts = mail_df["_source"].value_counts().to_dict()
    print(f"  Auto-detected sources: {src_counts}")
    print(f"  (add a '_source' column to the file to label 'goldaddition' rows separately)")

# ── Build lookup from gold reference ──────────────────────────────────────────
print(f"\nLoading gold reference: {GOLD_FILE.name}")

npi_to_addresses:  dict[str, list[tuple]] = {}  # npi  → [(raw, norm, src, npi, mid)]
name_to_addresses: dict[str, list[tuple]] = {}  # name → [(raw, norm, src, npi, mid)]

try:
    sheets = pd.read_excel(GOLD_FILE, sheet_name=None, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {GOLD_FILE} not found")

for sheet_name, df in sheets.items():
    print(f"  Sheet '{sheet_name}': {len(df)} rows")
    df = df.fillna("")
    cols = detect_gold_cols(df, sheet_name)

    npi_col    = find_col(df, ["npi", "NPI"])
    middle_col = find_col(df, ["middle_name", "MiddleName"])

    for _, row in df.iterrows():
        npi    = norm(row[npi_col])    if npi_col    else ""
        middle = norm(row[middle_col]) if middle_col else ""
        middle1 = middle[0] if middle else ""

        # Name keys from gold
        if cols["fullname"] and norm(row.get(cols["fullname"], "")):
            name_keys = fullname_variants(norm(row[cols["fullname"]]))
        else:
            last  = norm(row.get(cols["last"],  "")) if cols["last"]  else ""
            first = norm(row.get(cols["first"], "")) if cols["first"] else ""
            name_keys = name_variants(last, first)

        addr_pairs = get_all_addr_keys(row, cols)

        for name_key in name_keys:
            name_to_addresses.setdefault(name_key, [])
            for raw_addr, norm_addr_key in addr_pairs:
                name_to_addresses[name_key].append(
                    (raw_addr, norm_addr_key, GOLD_FILE.name, npi, middle1)
                )

        if npi:
            npi_to_addresses.setdefault(npi, [])
            for raw_addr, norm_addr_key in addr_pairs:
                npi_to_addresses[npi].append(
                    (raw_addr, norm_addr_key, GOLD_FILE.name, npi, middle1)
                )

print(f"\n  Gold NPI keys   : {len(npi_to_addresses)}")
print(f"  Gold name keys  : {len(name_to_addresses)}")

# ── Classify each Round 2 row ─────────────────────────────────────────────────
print(f"\nClassifying {len(mail_df)} Round 2 rows ...")

statuses        = []
addr_counts     = []
addresses_found = []
match_methods   = []

for _, row in mail_df.iterrows():
    npi = norm(row[r2_npi_col]) if r2_npi_col else ""

    # Gather address entries: NPI lookup first, then name lookup
    all_entries: list[tuple] = []
    used_method = "name"

    if npi and npi in npi_to_addresses:
        all_entries = npi_to_addresses[npi]
        used_method = "npi"
    else:
        a_val = row[r2_col_a] if r2_col_a else ""
        b_val = row[r2_col_b] if r2_col_b else ""
        keys  = round2_name_keys(a_val, b_val)
        for key in keys:
            if key in name_to_addresses:
                all_entries.extend(name_to_addresses[key])

    if not all_entries:
        statuses.append("not_in_gold")
        addr_counts.append(0)
        addresses_found.append("")
        match_methods.append("none")
        continue

    # Multiple distinct NPIs → different people with same name
    npis    = {e[3] for e in all_entries if e[3]}
    middles = {e[4] for e in all_entries if e[4]}
    different_people = (
        len(npis) > 1 or
        (len(middles) > 1 and "" not in middles)
    )

    if different_people:
        statuses.append("name_collision_different_people")
        addr_counts.append(len(npis) or len(middles))
        npis_str = " | ".join(sorted(npis)) if npis else "no NPI"
        addresses_found.append(f"NPIs: {npis_str}")
        match_methods.append(used_method)
        continue

    # Deduplicate by raw address string
    seen_raw: dict[str, str] = {}
    for raw, norm_a, src, n, mid in all_entries:
        seen_raw.setdefault(raw, norm_a)

    unique_raw   = list(seen_raw.keys())
    unique_norms = list(seen_raw.values())
    addr_counts.append(len(unique_raw))
    addresses_found.append(" | ".join(unique_raw))
    match_methods.append(used_method)

    if len(unique_raw) == 1:
        statuses.append("unique")
    elif len(set(unique_norms)) == 1:
        statuses.append("same_address_variant")
    else:
        statuses.append("multiple_addresses")

mail_df["address_status"]  = statuses
mail_df["address_count"]   = addr_counts
mail_df["addresses_found"] = addresses_found
mail_df["_match_method"]   = match_methods

# ── Summary ───────────────────────────────────────────────────────────────────
status_counts = pd.Series(statuses).value_counts()
print(f"\nResults (all rows):")
for status, cnt in status_counts.items():
    print(f"  {status}: {cnt}")

print(f"\nMultiple-address rows by source:")
multi_tmp = mail_df[mail_df["address_status"] == "multiple_addresses"]
for src, cnt in multi_tmp["_source"].value_counts().items():
    print(f"  {src}: {cnt}")

# ── Subsets ───────────────────────────────────────────────────────────────────
multi      = mail_df[mail_df["address_status"] == "multiple_addresses"]
variant    = mail_df[mail_df["address_status"] == "same_address_variant"]
not_in     = mail_df[mail_df["address_status"] == "not_in_gold"]
collisions = mail_df[mail_df["address_status"] == "name_collision_different_people"]

# ── Side-by-side compare for multiple_addresses rows ─────────────────────────
compare_rows = []
for _, row in multi.iterrows():
    r = {}
    r["_source"]      = row.get("_source", "")
    r["col_a"]        = row.get(r2_col_a,   "") if r2_col_a   else ""
    r["col_b"]        = row.get(r2_col_b,   "") if r2_col_b   else ""
    r["mail_address"] = row.get(r2_addr_col, "") if r2_addr_col else ""
    r["mail_city"]    = row.get(r2_city_col, "") if r2_city_col else ""
    r["mail_zip"]     = row.get(r2_zip_col,  "") if r2_zip_col  else ""
    parsed = parse_addresses_found(str(row.get("addresses_found", "")))
    for n, (addr, city, z) in enumerate(parsed, start=1):
        r[f"gold_addr_{n}"] = addr
        r[f"gold_city_{n}"] = city
        r[f"gold_zip_{n}"]  = z
    compare_rows.append(r)

compare_df = pd.DataFrame(compare_rows)

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    mail_df.to_excel(writer, sheet_name="all",                        index=False)
    multi.to_excel(  writer, sheet_name="multiple_addresses",         index=False)
    compare_df.to_excel(writer, sheet_name="multiple_addresses_compare", index=False)
    variant.to_excel(writer, sheet_name="same_address_variant",       index=False)
    collisions.to_excel(writer, sheet_name="name_collision",          index=False)
    not_in.to_excel( writer, sheet_name="not_in_gold",                index=False)
    print(f"  all                          : {len(mail_df)}")
    print(f"  multiple_addresses           : {len(multi)}")
    print(f"  multiple_addresses_compare   : {len(compare_df)} rows, {len(compare_df.columns)} cols")
    print(f"  same_address_variant         : {len(variant)}")
    print(f"  name_collision               : {len(collisions)}")
    print(f"  not_in_gold                  : {len(not_in)}")

print("\nDone.")
