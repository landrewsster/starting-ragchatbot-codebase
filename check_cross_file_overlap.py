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
FILES = [
    BASE / "physician_pa_nurse_20260422_combined.xlsx",
    BASE / "nurse_20260422_combined.xlsx",
    BASE / "MailingListAddition20260423.xlsx",
]
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

# ── Write output ─────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT}")
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    matrix_df.to_excel(writer,  sheet_name="overlap_matrix",  index=False)
    counts_df.to_excel(writer,  sheet_name="file_counts",     index=False)
    overlap_df.to_excel(writer, sheet_name="overlap_detail",  index=False)
    print(f"  overlap_matrix : {len(matrix_df)} file pairs")
    print(f"  file_counts    : {len(counts_df)} sheets")
    print(f"  overlap_detail : {len(overlap_df)} overlapping providers")

print("\nDone.")
