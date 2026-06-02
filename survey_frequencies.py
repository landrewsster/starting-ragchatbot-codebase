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
BASE = Path.home() / "Downloads" / "CRC MDH Project" / "MDH analysis"

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

# Demographic column patterns — questions shown to ALL respondents
# Defined early so free-text detection can exclude these columns from cardinality check.
DEMO_PATTERNS = [
    r"what is your profession",
    r"do you primarily see pregnant",
    r"how long have you been practicing",
    r"what is your primary specialty",
    r"what is your secondary or sub.specialty",
    r"which of the following best describes your primary practice setting",
    r"how would you describe the insurance status",
    r"what is the county of your practice",
    r"what is your gender",
    r"what is your age",
    r"what is your race/ethnicity",
]

def is_demo_col(col):
    return any(re.search(p, col, re.IGNORECASE) for p in DEMO_PATTERNS)

# Free-text / specify columns — skip by name pattern or high cardinality.
# Demographic columns are always categorical regardless of cardinality (e.g. county).
def _looks_like_free_text(series):
    """True if the column reads as open-ended: nearly all non-empty values are unique."""
    non_empty = series[series.str.strip().ne("")]
    if len(non_empty) < 5:
        return False
    unique_ratio = non_empty.nunique() / len(non_empty)
    return unique_ratio > 0.7  # >70% unique values → almost certainly free text

free_text_cols = {
    c for c in df.columns
    if c not in checkbox_cols    # never treat binary checkbox columns as free text
    and not is_demo_col(c)       # demographic columns are always categorical
    and (
        re.search(r'\bspecify\b|\bplease describe\b|\bexplain\b', c, re.IGNORECASE)
        or c.strip() == "Specify"
        or _looks_like_free_text(df[c].astype(str))
    )
}

# Detect timestamp columns before building skip set
def _looks_like_timestamps(series):
    """True if >50% of non-empty values parse as datetimes (ISO or US M/D/YY format)."""
    non_empty = series[series.str.strip().ne("")].head(20)
    if len(non_empty) < 3:
        return False
    iso = non_empty.str.match(r'^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}').sum()
    us  = non_empty.str.match(r'^\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}').sum()
    return (iso + us) / len(non_empty) > 0.5

TIMESTAMP_COL_PATTERNS = [
    r'^screener_start$', r'^screener_end$',
    r'^survey_start$',   r'^survey_end$',
    r'^inel_demo_start$', r'^inel_demo_end$',
]

timestamp_cols = [
    c for c in df.columns
    if re.match(r'^Unnamed:\s*\d+$', c.strip())
    or _looks_like_timestamps(df[c].astype(str))
    or any(re.search(p, c.strip(), re.IGNORECASE) for p in TIMESTAMP_COL_PATTERNS)
]

# System / admin columns — skip (Record ID, empty, and all timestamp columns)
system_cols = {
    c for c in df.columns
    if c.strip() in ("Record ID", "")
} | set(timestamp_cols)

# Find completion status column(s) — REDCap names them "<Form> Complete?"
# Label each by the last substantive question before it so the section is clear.
COMPLETE_RE = re.compile(r'complete\??\s*$|_complete\s*$', re.IGNORECASE)

complete_cols = [
    c for c in df.columns
    if COMPLETE_RE.search(c.strip())
    and c.strip() not in ("Record ID", "")
]

def _complete_col_label(col):
    """Return 'Complete? [after: <preceding question>]' for disambiguation."""
    idx = list(df.columns).index(col)
    for i in range(idx - 1, -1, -1):
        prev = df.columns[i]
        if (prev not in system_cols and prev not in timestamp_cols
                and not COMPLETE_RE.search(prev.strip())
                and prev.strip()):
            return f"Complete? [after: {short_q(prev, 60)}]"
    return f"Complete? [col {idx}]"

complete_col_labels = {c: _complete_col_label(c) for c in complete_cols}

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

# Record ID column (needed in both timestamp and free-text sections)
record_id_col = next((c for c in df.columns if c.strip() == "Record ID"), None)

# ── Completion summary ────────────────────────────────────────────────────────
if complete_cols:
    print(f"\nCompletion status ({len(complete_cols)} form(s)):")
    for cc in complete_cols:
        label = complete_col_labels.get(cc, cc)
        if cc in eligible.columns:
            counts = eligible[cc].str.strip().value_counts(dropna=False)
            print(f"\n  {label} — eligible (n={len(eligible)}):")
            for val, n in counts.items():
                print(f"    {str(val):<30} {n}")
else:
    print("\nNo 'Complete?' column found in data")

# ── Completion time ───────────────────────────────────────────────────────────
# Start:            screener_start_time  (all respondents)
# End (eligible):   survey_end_time
# End (ineligible): dem_end_time_2

def _find_col(columns, pattern):
    return next((c for c in columns if re.search(pattern, c, re.IGNORECASE)), None)

start_col      = _find_col(df.columns, r'^screener_start$')
end_elig_col   = _find_col(df.columns, r'^survey_end$')
end_inelig_col = _find_col(df.columns, r'^inel_demo_end$')

print(f"\nCompletion time columns:")
print(f"  Start       : {start_col or 'NOT FOUND'}")
print(f"  End eligible: {end_elig_col or 'NOT FOUND'}")
print(f"  End inelig. : {end_inelig_col or 'NOT FOUND'}")

# Verify mapping: show how many eligible vs ineligible have values in each col
for label, col in [("start", start_col), ("end_elig", end_elig_col), ("end_inelig", end_inelig_col)]:
    if col and col in df.columns:
        n_elig   = eligible[col].astype(str).str.strip().ne("").sum()   if col in eligible.columns   else 0
        n_inelig = ineligible[col].astype(str).str.strip().ne("").sum() if col in ineligible.columns else 0
        sample   = df[col].dropna().astype(str).str.strip()
        sample   = sample[sample != ""].iloc[0] if not sample[sample != ""].empty else ""
        print(f"    {label} ({col}): eligible={n_elig}, ineligible={n_inelig}, e.g. '{sample}'")

completion_rows = []

for grp_name, grp_df, end_col in [
    ("eligible",   eligible,   end_elig_col),
    ("ineligible", ineligible, end_inelig_col),
]:
    if grp_df.empty:
        continue
    if not start_col or start_col not in grp_df.columns:
        print(f"  WARNING: start column not found for {grp_name}")
        continue
    if not end_col or end_col not in grp_df.columns:
        print(f"  WARNING: end column not found for {grp_name}")
        continue

    start    = pd.to_datetime(grp_df[start_col], errors="coerce")
    end      = pd.to_datetime(grp_df[end_col],   errors="coerce")
    duration = (end - start).dt.total_seconds() / 60

    valid = start.notna() & end.notna() & (duration >= 0)
    d = duration[valid]
    print(f"\n  {grp_name} (n={valid.sum()}):")
    if not d.empty:
        print(f"    Mean   : {d.mean():.1f} min")
        print(f"    Median : {d.median():.1f} min")
        print(f"    Min    : {d.min():.1f} min")
        print(f"    Max    : {d.max():.1f} min")

    rec_ids = grp_df[record_id_col] if record_id_col else grp_df.iloc[:, 0]
    for rid, s, e, dur in zip(rec_ids, start, end, duration):
        if pd.isna(s) or pd.isna(e) or dur < 0:
            continue
        completion_rows.append({
            "record_id":        rid,
            "group":            grp_name,
            "start_time":       s.strftime("%Y-%m-%d %H:%M"),
            "end_time":         e.strftime("%Y-%m-%d %H:%M"),
            "duration_minutes": round(dur, 1),
        })

completion_time_df = pd.DataFrame(completion_rows) if completion_rows else pd.DataFrame()

# Duration bucket summary
completion_summary_df = pd.DataFrame()
if not completion_time_df.empty:
    bins   = [0, 5, 10, 15, 20, 30, float("inf")]
    labels = ["<5 min", "5-9 min", "10-14 min", "15-19 min", "20-29 min", "30+ min"]
    completion_time_df["duration_bucket"] = pd.cut(
        completion_time_df["duration_minutes"], bins=bins, labels=labels, right=False
    )
    completion_summary_df = (
        completion_time_df.groupby(["duration_bucket", "group"], observed=True)
        .size().reset_index(name="n")
        .pivot(index="duration_bucket", columns="group", values="n")
        .fillna(0).astype(int).reset_index()
        .rename(columns={"duration_bucket": "Duration"})
    )

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
    out.insert(0, "Question", re.sub(r'\s+', ' ', str(question_text)).strip())
    out.insert(1, "Type", "Single choice")
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
            "Question":        re.sub(r'\s+', ' ', str(parent_q)).strip(),
            "Type":            "Select all that apply",
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
        if col in free_text_cols or col in system_cols:
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
            # Single-choice — use disambiguated label for Complete? columns
            label   = complete_col_labels.get(col, col)
            ordered = ordered_for_col(col)
            t = freq_single(df_sub[col], label, ordered)
            if t is not None:
                parts.append(t)

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

# ── Build sheets ──────────────────────────────────────────────────────────────
print("\nBuilding frequency tables ...")

elig_freqs   = build_all_frequencies(eligible)
inelig_freqs = build_all_frequencies(ineligible)

# ── Build free-text sheet ─────────────────────────────────────────────────────

# Collect non-empty free-text responses in survey column order.
# Demographic free-text: include both groups. Main survey free-text: eligible only.
free_text_rows = []
seen_ft_cols = set()
for col in df.columns:
    if col not in free_text_cols or col in seen_ft_cols:
        continue
    if col in timestamp_cols or col in system_cols:
        continue
    seen_ft_cols.add(col)
    groups = (
        [("eligible", eligible), ("ineligible", ineligible)]
        if is_demo_col(col)
        else [("eligible", eligible)]
    )
    for grp_name, grp_df in groups:
        if col not in grp_df.columns:
            continue
        for _, row in grp_df.iterrows():
            val = str(row[col]).strip()
            if val and val.lower() not in ("nan", "", "checked", "unchecked", "0", "1"):
                rec = {
                    "group":    grp_name,
                    "question": short_q(col),
                    "response": val,
                    "recode":   "",
                }
                if record_id_col and record_id_col in grp_df.columns:
                    rec = {"record_id": row[record_id_col], **rec}
                free_text_rows.append(rec)

free_text_df = pd.DataFrame(free_text_rows) if free_text_rows else pd.DataFrame()
print(f"  Free-text responses collected: {len(free_text_rows)}")

# ── Build county sheet ────────────────────────────────────────────────────────
# County is a free-text field and REDCap may duplicate it across arms.
# Collect individual responses from ALL matching columns across both groups.
county_pattern = r'what is the county of your practice'
county_candidates = [
    c for c in df.columns
    if re.search(county_pattern, c, re.IGNORECASE)
    and "(choice=" not in c
]
print(f"\nCounty columns found: {len(county_candidates)}")

def normalize_county(val):
    """Basic standardization: strip, collapse spaces, St. → St, title case."""
    val = val.strip()
    val = re.sub(r'\s+', ' ', val)
    val = re.sub(r'\bSt\.', 'St', val, flags=re.IGNORECASE)
    val = val.title()
    return val

# ── County recode rules ───────────────────────────────────────────────────────
# Applied after normalize_county (which title-cases everything).
# Comparisons are case-insensitive (keys stored lower-case).
# "" (blank)  = invalid entry, excluded from county frequency table.
# "N/A"       = respondent does not practice in a Minnesota county
#               (out-of-state, ineligible practice type, or non-specific response).
COUNTY_RECODES = {
    # Invalid / uninterpretable entries
    "usa":                  "",      # not a county name

    # Health system name entered instead of county
    "mille lacs health":    "Mille Lacs",

    # "County" suffix — standardize to county name only
    "pima county":          "Pima",

    # Out-of-state responses
    "la crosse, wisconsin": "N/A",   # La Crosse County, WI
    "washburn wi":          "N/A",   # Washburn County, WI
    "st croix wi":          "N/A",   # St. Croix County, WI

    # Respondent does not serve pregnant/breastfeeding patients
    "don't care for pregnant or breastfeeding patients":                        "N/A",
    "unlikely to care for pregnant or breastfeeding patients in hospice":       "N/A",
    "none":                 "N/A",
}

def recode_county(normalized):
    """Apply specific recodes after normalization. Returns recoded value."""
    return COUNTY_RECODES.get(normalized.lower(), normalized)

county_df = pd.DataFrame()
county_freq_df = pd.DataFrame()

if county_candidates:
    county_rows = []
    for grp_name, grp_df in [("eligible", eligible), ("ineligible", ineligible)]:
        for _, row in grp_df.iterrows():
            raw = ""
            for col in county_candidates:
                if col in grp_df.columns:
                    v = str(row[col]).strip()
                    if v and v.lower() not in ("nan", ""):
                        raw = v
                        break
            if raw:
                normed  = normalize_county(raw)
                recoded = recode_county(normed)
                rec = {
                    "group":             grp_name,
                    "county_raw":        raw,
                    "county_normalized": normed,
                    "county_recoded":    recoded,
                }
                if record_id_col and record_id_col in grp_df.columns:
                    rec = {"record_id": row[record_id_col], **rec}
                county_rows.append(rec)

    if county_rows:
        county_df = (
            pd.DataFrame(county_rows)
            .sort_values(["group", "county_recoded"], key=lambda s: s.str.lower())
            .reset_index(drop=True)
        )
        print(f"  County responses: {len(county_df)} "
              f"(eligible: {(county_df['group']=='eligible').sum()}, "
              f"ineligible: {(county_df['group']=='ineligible').sum()})")

        # Frequency table uses county_recoded; blanks excluded, N/A included
        freq_data     = county_df[county_df["county_recoded"].str.strip().ne("")]
        elig_counts   = freq_data[freq_data["group"]=="eligible"]["county_recoded"].value_counts()
        inelig_counts = freq_data[freq_data["group"]=="ineligible"]["county_recoded"].value_counts()
        elig_n   = len(freq_data[freq_data["group"]=="eligible"])
        inelig_n = len(freq_data[freq_data["group"]=="ineligible"])

        all_counties = sorted(set(elig_counts.index) | set(inelig_counts.index))
        freq_rows = []
        for county in all_counties:
            e_n = int(elig_counts.get(county, 0))
            i_n = int(inelig_counts.get(county, 0))
            freq_rows.append({
                "County":          county,
                "Eligible n":      e_n,
                "Eligible %":      round(e_n / elig_n * 100, 1) if elig_n else None,
                "Ineligible n":    i_n,
                "Ineligible %":    round(i_n / inelig_n * 100, 1) if inelig_n else None,
                "Total n":         e_n + i_n,
            })

        county_freq_df = (
            pd.DataFrame(freq_rows)
            .sort_values("Total n", ascending=False)
            .reset_index(drop=True)
        )
        totals = pd.DataFrame([{
            "County":       "TOTAL",
            "Eligible n":   elig_n,
            "Eligible %":   100.0 if elig_n else None,
            "Ineligible n": inelig_n,
            "Ineligible %": 100.0 if inelig_n else None,
            "Total n":      elig_n + inelig_n,
        }])
        county_freq_df = pd.concat([county_freq_df, totals], ignore_index=True)
        print(f"  County frequency: {len(county_freq_df)-1} unique counties")
else:
    print("  County column not found — skipping county sheets")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")

sheets = {
    "eligible":           elig_freqs,
    "ineligible":         inelig_freqs,
    "county_freq":        county_freq_df,
    "county":             county_df,
    "free_text":          free_text_df,
    "completion_time":    completion_time_df,
    "completion_summary": completion_summary_df,
}

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for sheet_name, data in sheets.items():
        if data is not None and not data.empty:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {sheet_name:<12}: {len(data)} rows")
        else:
            print(f"  {sheet_name:<12}: (empty — skipped)")

print("\nDone.")
