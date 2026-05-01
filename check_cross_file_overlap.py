#!/usr/bin/env python3
"""
check_cross_file_overlap.py

Check for provider overlap across multiple files/sheets.
Matches on NPI first, then name (last + first + middle initial) as fallback.

Reports:
  - Provider count per file/sheet
  - Which providers appear in more than one file
  - Overlap matrix between files

Usage:
    python3 check_cross_file_overlap.py
"""

import re
import sys
from pathlib import Path
from itertools import combinations

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE  = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
# Files treated as the "gold" provider source
PROVIDER_FILES = [
    BASE / "physician_pa_nurse_20260422_combined.xlsx",
    BASE / "nurse_20260422_combined.xlsx",
]
# File to check against the gold source
ADDITION_FILE  = BASE / "MailingListAddition20260423.xlsx"

FILES  = PROVIDER_FILES + [ADDITION_FILE]
OUTPUT = BASE / "cross_file_overlap.xlsx"

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s):
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

# ── Load all files into a flat list of records ────────────────────────────────
# Each record: {source_label, npi, name_key, last, first, middle, row_data}
print("\nLoading files ...")
all_records: list[dict] = []
file_labels: list[str] = []

for filepath in FILES:
    print(f"\n  {filepath.name}")
    if not filepath.exists():
        print(f"    WARNING: not found, skipping")
        continue

    try:
        sheets = pd.read_excel(filepath, sheet_name=None, dtype=str)
    except Exception as e:
        print(f"    ERROR: {e}")
        continue

    for sheet_name, df in sheets.items():
        label = f"{filepath.stem} / {sheet_name}"
        file_labels.append(label)

        npi_col    = find_col(df, ["npi", "NPI"])
        last_col   = find_col(df, ["last_name", "LastName"])
        first_col  = find_col(df, ["first_name", "FirstName", "First Name"])
        middle_col = find_col(df, ["middle_name", "MiddleName"])

        count = 0
        for _, row in df.iterrows():
            npi    = norm(row[npi_col])    if npi_col    else ""
            last   = norm(row[last_col])   if last_col   else ""
            first  = norm(row[first_col])  if first_col  else ""
            middle = norm(row[middle_col]) if middle_col else ""
            mid1   = middle[0] if middle else ""

            name_key = f"{last}|{first}|{mid1}"

            all_records.append({
                "source":   label,
                "npi":      npi,
                "name_key": name_key,
                "last":     last,
                "first":    first,
                "middle":   middle,
            })
            count += 1

        print(f"    Sheet '{sheet_name}': {count} rows  "
              f"(npi={npi_col} last={last_col} first={first_col} middle={middle_col})")

print(f"\n  Total records loaded: {len(all_records)}")

# ── Build provider index: unique_id → list of source labels ──────────────────
# Priority: NPI if present, else name_key
print("\nBuilding provider index ...")
npi_to_sources:  dict[str, list[str]] = {}
name_to_sources: dict[str, list[str]] = {}

for rec in all_records:
    if rec["npi"] and rec["npi"] not in ("nan", ""):
        npi_to_sources.setdefault(rec["npi"], []).append(rec["source"])
    elif rec["name_key"] and rec["name_key"] != "||":
        name_to_sources.setdefault(rec["name_key"], []).append(rec["source"])

# ── Find providers appearing in more than one file ────────────────────────────
def files_from_sources(sources: list[str]) -> set[str]:
    """Extract the file name (before ' / ') from each source label."""
    return {s.split(" / ")[0] for s in sources}

npi_overlaps  = {k: v for k, v in npi_to_sources.items()  if len(files_from_sources(v)) > 1}
name_overlaps = {k: v for k, v in name_to_sources.items() if len(files_from_sources(v)) > 1}

print(f"  Providers in multiple files (by NPI) : {len(npi_overlaps)}")
print(f"  Providers in multiple files (by name): {len(name_overlaps)}")

# ── Build overlap matrix ──────────────────────────────────────────────────────
file_names = [f.stem for f in FILES]
matrix_data = {}

for f1, f2 in combinations(file_names, 2):
    pair = f"{f1} ↔ {f2}"
    shared_npi  = sum(1 for sources in npi_to_sources.values()
                      if any(f1 in s for s in sources) and any(f2 in s for s in sources))
    shared_name = sum(1 for sources in name_to_sources.values()
                      if any(f1 in s for s in sources) and any(f2 in s for s in sources))
    matrix_data[pair] = {"shared_by_npi": shared_npi, "shared_by_name": shared_name,
                          "total_shared": shared_npi + shared_name}
    print(f"  {pair}: {shared_npi} NPI matches, {shared_name} name matches")

# ── Build overlap detail rows ────────────────────────────────────────────────
overlap_rows = []
for npi, sources in npi_overlaps.items():
    file_list = sorted(files_from_sources(sources))
    overlap_rows.append({
        "match_type": "npi",
        "identifier": npi,
        "files":      " | ".join(file_list),
        "sources":    " | ".join(sorted(set(sources))),
    })

for name_key, sources in name_overlaps.items():
    parts = name_key.split("|")
    file_list = sorted(files_from_sources(sources))
    overlap_rows.append({
        "match_type": "name",
        "identifier": name_key,
        "last":  parts[0] if len(parts) > 0 else "",
        "first": parts[1] if len(parts) > 1 else "",
        "middle_initial": parts[2] if len(parts) > 2 else "",
        "files":   " | ".join(file_list),
        "sources": " | ".join(sorted(set(sources))),
    })

overlap_df = pd.DataFrame(overlap_rows)
matrix_df  = pd.DataFrame([
    {"file_pair": pair, **counts} for pair, counts in matrix_data.items()
])

# Per-file counts
counts_rows = []
for label in file_labels:
    n = sum(1 for r in all_records if r["source"] == label)
    counts_rows.append({"source": label, "row_count": n})
counts_df = pd.DataFrame(counts_rows)

# ── Directional check: MailingListAddition → provider files ──────────────────
# Build a set of all identifiers from the provider files only
provider_npi_set  = set()
provider_name_set = set()
provider_stems    = {f.stem for f in PROVIDER_FILES}

for rec in all_records:
    file_stem = rec["source"].split(" / ")[0]
    if file_stem not in provider_stems:
        continue
    if rec["npi"] and rec["npi"] not in ("nan", ""):
        provider_npi_set.add(rec["npi"])
    if rec["name_key"] and rec["name_key"] != "||":
        provider_name_set.add(rec["name_key"])

print(f"\nProvider reference set: {len(provider_npi_set)} NPIs, "
      f"{len(provider_name_set)} name keys")

# Load MailingListAddition fresh to get all original columns
print(f"\nChecking {ADDITION_FILE.name} against provider files ...")
try:
    add_sheets = pd.read_excel(ADDITION_FILE, sheet_name=None, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {ADDITION_FILE} not found")

addition_frames = []
for sheet_name, df in add_sheets.items():
    npi_col    = find_col(df, ["npi", "NPI"])
    last_col   = find_col(df, ["last_name", "LastName"])
    first_col  = find_col(df, ["first_name", "FirstName", "First Name"])
    middle_col = find_col(df, ["middle_name", "MiddleName"])

    found_flags   = []
    match_methods = []
    for _, row in df.iterrows():
        npi    = norm(row[npi_col])    if npi_col    else ""
        last   = norm(row[last_col])   if last_col   else ""
        first  = norm(row[first_col])  if first_col  else ""
        middle = norm(row[middle_col]) if middle_col else ""
        mid1   = middle[0] if middle else ""
        name_key = f"{last}|{first}|{mid1}"

        if npi and npi in provider_npi_set:
            found_flags.append(True)
            match_methods.append("npi")
        elif name_key != "||" and name_key in provider_name_set:
            found_flags.append(True)
            match_methods.append("name")
        else:
            found_flags.append(False)
            match_methods.append("not_found")

    df["in_provider_files"] = found_flags
    df["match_method"]      = match_methods
    addition_frames.append((sheet_name, df))

    n_found = sum(found_flags)
    n_miss  = len(found_flags) - n_found
    print(f"  Sheet '{sheet_name}': {n_found} found in provider files, "
          f"{n_miss} NOT found")
    for m, cnt in pd.Series(match_methods).value_counts().items():
        print(f"    {m}: {cnt}")

# Combine all addition sheets and extract missing
all_addition_df = pd.concat([df for _, df in addition_frames], ignore_index=True)
missing_df = all_addition_df[all_addition_df["in_provider_files"] == False].drop(
    columns=["in_provider_files", "match_method"]
)
print(f"\n  Total in addition file : {len(all_addition_df)}")
print(f"  Found in provider files: {sum(all_addition_df['in_provider_files'])}")
print(f"  NOT in provider files  : {len(missing_df)}")

# ── Write output ─────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT}")
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    matrix_df.to_excel(writer,        sheet_name="overlap_matrix",         index=False)
    counts_df.to_excel(writer,        sheet_name="file_counts",            index=False)
    overlap_df.to_excel(writer,       sheet_name="overlap_detail",         index=False)
    all_addition_df.to_excel(writer,  sheet_name="addition_flagged",       index=False)
    missing_df.to_excel(writer,       sheet_name="addition_not_in_providers", index=False)
    print(f"  overlap_matrix            : {len(matrix_df)} file pairs")
    print(f"  file_counts               : {len(counts_df)} sheets")
    print(f"  overlap_detail            : {len(overlap_df)} overlapping providers")
    print(f"  addition_flagged          : {len(all_addition_df)} rows")
    print(f"  addition_not_in_providers : {len(missing_df)} rows")

print("\nDone.")
