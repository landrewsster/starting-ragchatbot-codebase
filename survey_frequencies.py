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

# System / admin columns — skip
system_cols = {
    c for c in df.columns
    if c.strip() in ("Record ID", "Complete?", "")
    or re.match(r'complete\??\s*$', c.strip(), re.IGNORECASE)
}

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
# Survey routing:
#   Q1 (prenatal/delivery/postpartum) "No"  → sent to demographics (ineligible)
#   Q1 "Yes" + Q2 (days see patients) "0"   → sent to demographics (ineligible)
#   Q1 "Yes" + Q2 blank (not required)      → uncertain; warn and include as eligible
#   Q1 "Yes" + Q2 > 0                       → eligible

elig_col = next(
    (c for c in df.columns
     if re.search(r'prenatal.*delivery.*postpartum', c, re.IGNORECASE)
     or re.search(r'provide prenatal', c, re.IGNORECASE)),
    None
)

days_col = next(
    (c for c in df.columns
     if re.search(r'how many days.*see patients|average week.*how many', c, re.IGNORECASE)
     and '(choice=' not in c),
    None
)

if elig_col:
    q1_yes = df[elig_col].str.strip().str.lower() == "yes"

    if days_col:
        q2_val = df[days_col].astype(str).str.strip()
        q2_zero  = q2_val.str.lower().isin(["0", "0 days", "0.0"])
        q2_blank = q2_val.isin(["", "nan", "NaN"])

        # Eligible: Q1=Yes AND Q2 is not "0" (blank = uncertain, include with warning)
        eligible   = df[ q1_yes & ~q2_zero].reset_index(drop=True)
        ineligible = df[~q1_yes | q2_zero ].reset_index(drop=True)

        n_yes       = int(q1_yes.sum())
        n_zero      = int((q1_yes & q2_zero).sum())
        n_blank_q2  = int((q1_yes & q2_blank).sum())
        n_gt0       = int((q1_yes & ~q2_zero & ~q2_blank).sum())

        print(f"\nEligibility split on:")
        print(f"  Q1: {short_q(elig_col, 80)}")
        print(f"  Q2: {short_q(days_col, 80)}")
        print(f"\n  Q1 = Yes                : {n_yes}")
        print(f"    Q2 > 0 days (eligible): {n_gt0}")
        print(f"    Q2 = 0 days (inelig.) : {n_zero}")
        print(f"    Q2 blank (uncertain)  : {n_blank_q2}  ← included as eligible; verify manually")
        print(f"  Q1 = No  (ineligible)   : {int((~q1_yes).sum())}")
        print(f"\n  Eligible   (full survey)          : {len(eligible)}")
        print(f"  Ineligible (screener + demo only) : {len(ineligible)}")
    else:
        # Q2 column not found — fall back to Q1 only
        eligible   = df[ q1_yes].reset_index(drop=True)
        ineligible = df[~q1_yes].reset_index(drop=True)
        print(f"\nEligibility split on Q1 only (Q2 column not found): {short_q(elig_col, 80)}")
        print(f"  Eligible   (full survey)          : {len(eligible)}")
        print(f"  Ineligible (screener + demo only) : {len(ineligible)}")
else:
    eligible   = df.copy()
    ineligible = pd.DataFrame(columns=df.columns)
    print("\nWARNING: eligibility column not found — treating all records as eligible")

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

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")

sheets = {
    "eligible":   elig_freqs,
    "ineligible": inelig_freqs,
}

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for sheet_name, data in sheets.items():
        if data is not None and not data.empty:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {sheet_name:<12}: {len(data)} rows")
        else:
            print(f"  {sheet_name:<12}: (empty — skipped)")

print("\nDone.")
