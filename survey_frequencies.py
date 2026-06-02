#!/usr/bin/env python3
"""
survey_frequencies.py

Generates frequency tables from the cannabis screening survey REDCap CSV export.
Handles single-choice, checkbox (select all that apply), Likert, and confidence
questions. Splits eligible (full survey) from ineligible (screener + demographics only).

Output: survey_frequencies.xlsx
  eligible    — all question frequencies for eligible respondents
  ineligible  — all question frequencies for ineligible respondents

Usage:
    python3 survey_frequencies.py                        # uses default filename
    python3 survey_frequencies.py my_data_file.csv       # specify filename
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE = Path.home() / "Downloads" / "CRC MDH Project"

if len(sys.argv) > 1:
    INPUT_FILE = BASE / sys.argv[1]
else:
    # Update this to match your actual filename
    INPUT_FILE = BASE / "survey_data.csv"

OUTPUT_FILE = INPUT_FILE.with_name(INPUT_FILE.stem + "_frequencies.xlsx")

print(f"Input : {INPUT_FILE.name}")
print(f"Output: {OUTPUT_FILE.name}")

# ── Load ──────────────────────────────────────────────────────────────────────
try:
    df = pd.read_csv(INPUT_FILE, dtype=str)
except FileNotFoundError:
    raise SystemExit(f"ERROR: {INPUT_FILE} not found")

df = df.fillna("")
print(f"Loaded {len(df)} rows, {len(df.columns)} columns")

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_checked(val):
    """True for any checkbox-positive coding: Checked / 1 / Yes / True."""
    return str(val).strip().lower() in ("checked", "1", "yes", "true")

def choice_label(col):
    """Extract option text from a checkbox column name."""
    m = re.search(r'\(choice=(.+?)\)\s*$', col)
    return m.group(1).strip() if m else col

def parent_question(col):
    """Extract parent question text from a checkbox column name."""
    q = re.sub(r'\s*\(select all that apply\).*$', '', col, flags=re.IGNORECASE)
    q = re.sub(r'\s*\(choice=.*$', '', q)
    return re.sub(r'\s+', ' ', q).strip()

def short_q(s, n=120):
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s[:n] + "…" if len(s) > n else s

def norm_val(v):
    return str(v).strip() if str(v).strip() not in ("", "nan") else None

# ── Column classification ─────────────────────────────────────────────────────
checkbox_cols = [c for c in df.columns if "(choice=" in c]

# Group checkbox columns by parent question text, preserving column order
checkbox_groups: dict[str, list[str]] = {}
for c in checkbox_cols:
    p = parent_question(c)
    checkbox_groups.setdefault(p, []).append(c)

# Free-text / specify columns — skip
free_text_cols = {
    c for c in df.columns
    if re.search(r'\bspecify\b|\bplease describe\b|\bexplain\b', c, re.IGNORECASE)
    or c.strip() == "Specify"
}

# Detect timestamp columns before building skip set
def _looks_like_timestamps(series):
    """True if >50% of non-empty values parse as datetimes."""
    non_empty = series[series.str.strip().ne("")].head(20)
    if len(non_empty) < 3:
        return False
    hits = non_empty.str.match(r'^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}').sum()
    return hits / len(non_empty) > 0.5

timestamp_cols = [
    c for c in df.columns
    if re.match(r'^Unnamed:\s*\d+$', c.strip())
    or _looks_like_timestamps(df[c].astype(str))
]

# System / admin columns — skip (Record ID, empty, and all timestamp columns)
system_cols = {
    c for c in df.columns
    if c.strip() in ("Record ID", "")
} | set(timestamp_cols)

# Find completion status column(s) — REDCap names them "<Form> Complete?"
complete_cols = [
    c for c in df.columns
    if re.search(r'complete\??\s*$', c.strip(), re.IGNORECASE)
    and c.strip() not in ("Record ID", "")
]

skip = set(checkbox_cols) | free_text_cols | system_cols

# Single-choice columns (everything not checkbox, free-text, or system)
single_cols = [c for c in df.columns if c not in skip]

# Ordered response categories for known question types
ORDERED_LIKERT     = [
    "Strongly agree", "Agree", "Neither agree nor disagree",
    "Disagree", "Strongly disagree", "Don't know",
]
ORDERED_CONFIDENCE = ["Very confident", "Somewhat confident", "Not confident", "Don't know"]

LIKERT_PATTERNS = [
    r"there is no safe level",
    r"potential risks.+outweigh",
    r"therapeutic reasons",
    r"contraindication to breastfeeding",
    r"accurately report",
    r"routine toxicology screening",
    r"clinicians should screen",
]
CONFIDENCE_PATTERNS = [r"talk about (cannabis|tobacco|alcohol)"]

def ordered_for_col(col):
    if any(re.search(p, col, re.IGNORECASE) for p in LIKERT_PATTERNS):
        return ORDERED_LIKERT
    if any(re.search(p, col, re.IGNORECASE) for p in CONFIDENCE_PATTERNS):
        return ORDERED_CONFIDENCE
    return None

print(f"\nColumn types detected:")
print(f"  Checkbox groups      : {len(checkbox_groups)}")
print(f"  Single-choice        : {len(single_cols)}")
print(f"  Free-text (skipped)  : {len(free_text_cols)}")

# ── Split eligible vs ineligible ──────────────────────────────────────────────
elig_col = next(
    (c for c in df.columns
     if re.search(r'prenatal.*delivery.*postpartum', c, re.IGNORECASE)
     or re.search(r'provide prenatal', c, re.IGNORECASE)),
    None
)

if elig_col:
    eligible   = df[df[elig_col].str.strip().str.lower() == "yes"].reset_index(drop=True)
    ineligible = df[df[elig_col].str.strip().str.lower() != "yes"].reset_index(drop=True)
    print(f"\nEligibility split on: {short_q(elig_col, 80)}")
    print(f"  Eligible   (full survey) : {len(eligible)}")
    print(f"  Ineligible (screener + demo only) : {len(ineligible)}")
else:
    eligible   = df.copy()
    ineligible = pd.DataFrame(columns=df.columns)
    print("\nWARNING: eligibility column not found — treating all records as eligible")

# ── Completion summary ────────────────────────────────────────────────────────
if complete_cols:
    print(f"\nCompletion status columns found: {complete_cols}")
    for cc in complete_cols:
        if cc in eligible.columns:
            counts = eligible[cc].str.strip().value_counts(dropna=False)
            print(f"\n  [{cc}] — eligible respondents (n={len(eligible)}):")
            for val, n in counts.items():
                print(f"    {str(val):<30} {n}")
else:
    print("\nNo 'Complete?' column found in data")

# ── Completion time ───────────────────────────────────────────────────────────
completion_time_df = pd.DataFrame()

print(f"\nTimestamp columns detected: {timestamp_cols if timestamp_cols else 'none'}")

if len(timestamp_cols) >= 2:
    # Assume first = start, second = end (REDCap order)
    start_col, end_col = timestamp_cols[0], timestamp_cols[1]
    for grp_name, grp_df in [("eligible", eligible), ("ineligible", ineligible)]:
        if grp_df.empty:
            continue
        starts = pd.to_datetime(grp_df[start_col], errors="coerce")
        ends   = pd.to_datetime(grp_df[end_col],   errors="coerce")
        mins   = (ends - starts).dt.total_seconds() / 60
        valid  = mins.dropna()
        valid  = valid[valid >= 0]  # drop negative (data errors)
        if not valid.empty:
            print(f"\n  Completion time — {grp_name} (n={len(valid)} with both timestamps):")
            print(f"    Mean   : {valid.mean():.1f} min")
            print(f"    Median : {valid.median():.1f} min")
            print(f"    Min    : {valid.min():.1f} min")
            print(f"    Max    : {valid.max():.1f} min")
elif len(timestamp_cols) == 1:
    # Only one timestamp — report it as survey date distribution
    ts_col = timestamp_cols[0]
    parsed = pd.to_datetime(eligible[ts_col], errors="coerce")
    print(f"\n  Single timestamp column found ('{ts_col}') — cannot compute duration.")
    print(f"  Survey dates range: {parsed.min()} to {parsed.max()}")

# ── Frequency functions ───────────────────────────────────────────────────────
def freq_single(series, question_text, ordered=None):
    """Frequency table for a single-choice column."""
    vals = series.apply(norm_val).dropna()
    n_total = len(vals)
    if n_total == 0:
        return None
    counts = vals.value_counts(dropna=True)
    if ordered:
        idx = [o for o in ordered if o in counts.index]
        idx += [v for v in counts.index if v not in idx]
        counts = counts.reindex(idx).dropna()
    pct = (counts / n_total * 100).round(1)
    out = pd.DataFrame({"Response": counts.index, "n": counts.values, "%": pct.values})
    out.insert(0, "Question", short_q(question_text))
    out["N (answered)"] = n_total
    return out

def freq_checkbox_group(df_sub, parent_q, child_cols):
    """Frequency table for a checkbox group."""
    # Denominator = rows where at least one child column is not blank
    has_response = df_sub[child_cols].apply(
        lambda col: col.str.strip().ne("")
    ).any(axis=1)
    n_denom = has_response.sum()
    if n_denom == 0:
        return None
    rows = []
    for col in child_cols:
        n_checked = df_sub.loc[has_response, col].apply(is_checked).sum()
        rows.append({
            "Question":        short_q(parent_q),
            "Response":        choice_label(col),
            "n":               int(n_checked),
            "%":               round(n_checked / n_denom * 100, 1),
            "N (denominator)": int(n_denom),
        })
    return pd.DataFrame(rows)

def build_all_frequencies(df_sub):
    """
    Build one combined frequency table for df_sub, processing columns in the
    order they appear in the original data (single-choice and checkbox groups
    are interleaved to match the survey flow).
    """
    parts = []

    # Walk columns in original order; emit single or checkbox as encountered
    emitted_checkbox_parents = set()

    for col in df_sub.columns:
        if col in skip:
            continue

        if col in checkbox_cols:
            # Emit the whole checkbox group when we hit the first child column
            parent = parent_question(col)
            if parent in emitted_checkbox_parents:
                continue
            emitted_checkbox_parents.add(parent)
            child_cols = [c for c in checkbox_groups[parent] if c in df_sub.columns]
            if child_cols:
                t = freq_checkbox_group(df_sub, parent, child_cols)
                if t is not None:
                    parts.append(t)
        else:
            # Single-choice
            ordered = ordered_for_col(col)
            t = freq_single(df_sub[col], col, ordered)
            if t is not None:
                parts.append(t)

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

# ── Build sheets ──────────────────────────────────────────────────────────────
print("\nBuilding frequency tables ...")

elig_freqs   = build_all_frequencies(eligible)
inelig_freqs = build_all_frequencies(ineligible)

# ── Build completion time sheet ───────────────────────────────────────────────
completion_time_rows = []

if len(timestamp_cols) >= 2:
    start_col, end_col = timestamp_cols[0], timestamp_cols[1]
    for grp_name, grp_df in [("eligible", eligible), ("ineligible", ineligible)]:
        if grp_df.empty:
            continue
        starts = pd.to_datetime(grp_df[start_col], errors="coerce")
        ends   = pd.to_datetime(grp_df[end_col],   errors="coerce")
        mins   = (ends - starts).dt.total_seconds() / 60
        # Attach per-record completion times
        for i, (m, s, e) in enumerate(zip(mins, starts, ends)):
            completion_time_rows.append({
                "group":       grp_name,
                "start_time":  str(s) if pd.notna(s) else "",
                "end_time":    str(e) if pd.notna(e) else "",
                "minutes":     round(m, 1) if pd.notna(m) and m >= 0 else None,
            })

completion_time_df = pd.DataFrame(completion_time_rows) if completion_time_rows else pd.DataFrame()

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")

sheets = {
    "eligible":        elig_freqs,
    "ineligible":      inelig_freqs,
    "completion_time": completion_time_df,
}

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for sheet_name, data in sheets.items():
        if data is not None and not data.empty:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {sheet_name:<12}: {len(data)} rows")
        else:
            print(f"  {sheet_name:<12}: (empty — skipped)")

print("\nDone.")
