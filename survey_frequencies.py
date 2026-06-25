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

import numpy as np
import pandas as pd

try:
    from scipy.stats import chi2_contingency, fisher_exact, norm as _norm, mannwhitneyu, wilcoxon
    _SCIPY = True
except ImportError:
    _SCIPY = False
    print("NOTE: scipy not installed — statistical tests will be skipped in cross-tabs.")
    print("      Install with: pip install scipy")

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE = Path.home() / "Downloads" / "CRC MDH Project" / "MDH analysis"

if len(sys.argv) > 1:
    INPUT_FILE = BASE / sys.argv[1]
else:
    INPUT_FILE = BASE / "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-15_0947_EDITED.csv"

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

# ── Merge pandas-deduplicated single-choice columns ───────────────────────────
# REDCap exports the same non-checkbox field in multiple instruments; pandas
# renames duplicates col.1, col.2, …  Each respondent answers in exactly one
# arm, so we coalesce into the base column and drop the extras.
_df_cols = set(df.columns)
_single_dedup: list[tuple[str, str]] = []
for _c in list(df.columns):
    _m = re.match(r'^(.+?)\.\d+$', _c)
    if _m and _m.group(1) in _df_cols and '(choice=' not in _c:
        _single_dedup.append((_m.group(1), _c))

if _single_dedup:
    for _base, _var in _single_dedup:
        _empty = df[_base].str.strip().eq("")
        n_filled = int(_empty.sum())
        df.loc[_empty, _base] = df.loc[_empty, _var]
        if n_filled > 0:
            _label = _var[:60] + "…" if len(_var) > 60 else _var
            print(f"  Merged '{_label}' → base ({n_filled} rows filled)")
    df = df.drop(columns=[_v for _, _v in _single_dedup])
    print(f"Merged {len(_single_dedup)} deduplicated column(s) into base columns")

# ── Filter to screener-complete respondents ───────────────────────────────────
# Exclude respondents who opened the survey but did not complete the screener.
_sc_col = next((c for c in df.columns if c.strip().lower() == "screener_complete"), None)
if _sc_col:
    n_before = len(df)
    df = df[df[_sc_col].str.strip().str.lower() == "complete"].reset_index(drop=True)
    print(f"Filtered to screener-complete: {len(df)} of {n_before} "
          f"({n_before - len(df)} excluded — did not finish screener)")
else:
    print("WARNING: Screener_complete column not found — all respondents included")

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_checked(val):
    """True for any checkbox-positive coding: Checked / 1 / Yes / True."""
    return str(val).strip().lower() in ("checked", "1", "yes", "true")

def choice_label(col):
    """Extract option text from a checkbox column name.
    Strips pandas dedup suffixes (.1, .2, …) before matching so that
    duplicate-exported columns are labelled identically to their originals."""
    col_clean = re.sub(r'\.\d+$', '', col.strip())
    m = re.search(r'\(choice=(.+?)\)\s*$', col_clean)
    return m.group(1).strip() if m else col_clean

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
# Response text that maps to tenure = 1 (20+ years).  Matched case-insensitively.
TENURE_HIGH_PATTERN = r"20 or more"

DEMO_PATTERNS = [
    r"what is your profession",
    r"do you primarily see pregnant",
    r"how long have you been practicing",
    r"what is your primary specialty",
    r"primary specialty",               # catches shorter / differently worded versions
    r"what is your secondary or sub.specialty",
    r"secondary.{0,30}specialty",       # "secondary specialty", "secondary or sub-specialty, if any", etc.
    r"sub.specialty",                   # "sub-specialty" alone
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

# "Other, please specify" columns — detected by name, not cardinality.
# These are shown inline in the frequency tables directly below their parent question.
# High-cardinality free-text columns (open-ended questions) stay in the free_text sheet only.
named_free_text_cols = {
    c for c in free_text_cols
    if re.search(r'\bspecify\b|\bplease describe\b|\bexplain\b', c, re.IGNORECASE)
    or c.strip() == "Specify"
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

# Optional manual label overrides — keys are exact column names from the CSV (case-sensitive).
# Use these when the auto-generated label isn't specific enough.
COMPLETE_COL_LABEL_OVERRIDES = {
    "Screener_complete":  "Complete? [after: In an average week, how many days do you see patients either…]",
    "Survey_complete":    "Complete? [after: full eligible survey]",
    "Inel_demo_complete": "Complete? [after: ineligible demographics]",
}

# Fallback column names used only when the named columns (Screener_start, etc.)
# are not present in the export. All six roles now have named columns, so these
# are set to None. Update any entry to "Unnamed: N" if reverting to an unnamed export.
TIMESTAMP_COL_OVERRIDES = {
    "screener_start":  None,
    "screener_end":    None,
    "survey_start":    None,
    "survey_end":      None,
    "inel_demo_start": None,
    "inel_demo_end":   None,
}

complete_cols = [
    c for c in df.columns
    if COMPLETE_RE.search(c.strip())
    and c.strip() not in ("Record ID", "")
]

def _complete_col_label(col):
    """Return a disambiguated label for a Complete? column.
    Checks COMPLETE_COL_LABEL_OVERRIDES first (case-insensitive), then auto-detects
    from the preceding column."""
    override = COMPLETE_COL_LABEL_OVERRIDES.get(col) or \
               COMPLETE_COL_LABEL_OVERRIDES.get(col.strip()) or \
               next((v for k, v in COMPLETE_COL_LABEL_OVERRIDES.items()
                     if k.lower() == col.strip().lower()), None)
    if override is not None:
        return override
    idx = list(df.columns).index(col)
    for i in range(idx - 1, -1, -1):
        prev = df.columns[i]
        if (prev not in system_cols and prev not in timestamp_cols
                and not COMPLETE_RE.search(prev.strip())
                and prev.strip()):
            return f"Complete? [after: {short_q(prev, 60)}]"
    return f"Complete? [col {idx}]"

complete_col_labels = {c: _complete_col_label(c) for c in complete_cols}

# County columns are handled separately in the county_freq sheet (with recoding).
# Exclude them from the eligible/ineligible frequency tables to avoid showing raw values.
county_skip_cols = {
    c for c in df.columns
    if re.search(r'what is the county of your practice', c, re.IGNORECASE)
    and '(choice=' not in c
}

skip = set(checkbox_cols) | free_text_cols | system_cols | county_skip_cols

# Single-choice columns (everything not checkbox, free-text, or system)
single_cols = [c for c in df.columns if c not in skip]

# ── Specialty column diagnostics ─────────────────────────────────────────────
# Two-part diagnostic:
#   Part 1 — every column whose name contains "specialty", with its classification.
#   Part 2 — columns GQ–GU by Excel-letter position, to find the main-survey
#             secondary specialty column (GS) whose label may not include the word
#             "specialty" and therefore doesn't appear in Part 1.
def _excel_col_letter(n):   # n is 1-based column index
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _col_class(c):
    if c in free_text_cols:  return "FREE_TEXT (excluded from freq/missingness)"
    if c in checkbox_cols:   return "checkbox"
    if c in system_cols:     return "system (excluded)"
    return "single-choice"

_spec_cols = [c for c in df.columns
              if re.search(r"specialty|sub.specialty", c, re.IGNORECASE)]
# Build position index so we can show Excel column letter alongside each column
_col_position = {c: i+1 for i, c in enumerate(df.columns)}
print("\nSpecialty column classification (with Excel column position):")
if _spec_cols:
    for _sc in _spec_cols:
        _n  = df[_sc].str.strip().ne("").sum()
        _ltr = _excel_col_letter(_col_position[_sc])
        print(f"  [col {_ltr}] [{_col_class(_sc)}] n={_n:>3}  {_sc[:90]}")
else:
    print("  (no columns with 'specialty' in their name)")

# Also print ALL single-choice columns whose Excel position falls in the
# analyst-supplied ranges.  Update _TARGET_LETTERS if the correct columns
# are at different positions.
_TARGET_LETTERS = {'GQ','GR','GS','GT', 'HT','HU','HV','HW'}
_target_rows = [(i+1, c) for i, c in enumerate(df.columns)
                if _excel_col_letter(i+1) in _TARGET_LETTERS]
if _target_rows:
    print("\nColumns GQ–GT and HT–HW by position:")
    for _i, _c in _target_rows:
        _ltr = _excel_col_letter(_i)
        _n   = df[_c].str.strip().ne("").sum()
        print(f"  [{_ltr}] [{_col_class(_c)}] n={_n:>3}  {_c[:90]}")

# ── Free-text column associations ─────────────────────────────────────────────
# Walk columns in survey order; each named free-text column (specify/describe)
# is attached to the last substantive question preceding it, so "Other, please
# specify" responses can be shown inline below the question they belong to.
_ft_q_key  = None
_cb_set_ft = set(checkbox_cols)
free_text_associations: dict[str, str] = {}   # ft_col → question key
for _c in df.columns:
    if _c in system_cols or _c in county_skip_cols:
        continue
    if _c in _cb_set_ft:
        _ft_q_key = parent_question(_c)
    elif _c in named_free_text_cols:
        if _ft_q_key is not None:
            free_text_associations[_c] = _ft_q_key
    elif _c not in free_text_cols:
        _ft_q_key = complete_col_labels.get(_c, _c)

# Reverse map: question key → list of associated specify columns
q_free_text_cols: dict[str, list[str]] = {}
for _fc, _qk in free_text_associations.items():
    q_free_text_cols.setdefault(_qk, []).append(_fc)

# Ordered response categories for known question types
ORDERED_LIKERT     = [
    "Strongly agree", "Agree", "Neither agree nor disagree",
    "Disagree", "Strongly disagree", "Don't know",
]
ORDERED_CONFIDENCE = ["Very confident", "Somewhat confident", "Not confident", "Don't know"]

# Numeric scores for ordinal tests (Wilcoxon, Mann-Whitney U).
# "Don't know" is intentionally absent → becomes NaN → excluded from tests.
LIKERT_SCORE_MAP = {
    "strongly agree":             5,
    "agree":                      4,
    "neither agree nor disagree": 3,
    "disagree":                   2,
    "strongly disagree":          1,
}
CONFIDENCE_SCORE_MAP = {
    "very confident":     3,
    "somewhat confident": 2,
    "not confident":      1,
}

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

# Branching rules — questions only shown to respondents who answered a parent
# question a specific way.  Add one dict per parent→child group.
BRANCHING_RULES = [
    {
        "child_patterns": [
            r"in what context did you receive the training",
            r"was the training focused solely on cannabis",
            r"what did the training.+include related to cannabis",
        ],
        "parent_pattern": r"have you had any formal training related to cannabis",
        "parent_value":   "yes",
    },

    # ── Screening question (Q1) branches ──────────────────────────────────────
    # Q2 — "Who do you screen, if only some?" → only "No, only some patients"
    {
        "child_patterns": [
            r"who do you or others.+screen.+if only some",
        ],
        "parent_pattern": r"screen all patients who are pregnant or breastfeeding for cannabis",
        "parent_value":   "No, only some patients",
    },
    # Q3 — "Why don't you screen all patients?" → both "No" groups
    {
        "child_patterns": [
            r"why don.t you or others.+screen all patients",
        ],
        "parent_pattern": r"screen all patients who are pregnant or breastfeeding for cannabis",
        "parent_value":   ["No, only some patients",
                           "No, don't screen any patients"],
    },
    # Q4 — "Why do you screen all patients?" → only "Yes, all patients"
    {
        "child_patterns": [
            r"why do you or others.+screen all patients for cannabis",
        ],
        "parent_pattern": r"screen all patients who are pregnant or breastfeeding for cannabis",
        "parent_value":   "Yes, all patients (i.e., universal screening)",
    },
    # Q5 — "Which approaches do you use to screen?" → all who screen (Yes + No, only some)
    {
        "child_patterns": [
            r"which.+approaches.+most commonly use to screen patients for cannabis",
        ],
        "parent_pattern": r"screen all patients who are pregnant or breastfeeding for cannabis",
        "parent_value":   ["Yes, all patients (i.e., universal screening)",
                           "No, only some patients"],
    },

    # ── Q14 documentation question branches ───────────────────────────────────
    # Q15 + Q16 — shown to those who answered "Yes" to Q14
    {
        "child_patterns": [
            r"how is this information recorded within the electronic health record",
            r"has your behavior around documenting this information changed since legalization",
        ],
        "parent_pattern": r"consistently record information in patient records about cannabis use",
        "parent_value":   "Yes",
    },
    # Q17 — shown to those who answered "No" to Q14
    {
        "child_patterns": [
            r"why don.t you or others.+consistently document this information",
        ],
        "parent_pattern": r"consistently record information in patient records about cannabis use",
        "parent_value":   "No",
    },

    # ── Q18 branch — Q19, Q20, Q21 shown to "Yes" ─────────────────────────────
    # Q19 and Q20 are open-text (free_text sheet only) so this rule only fires
    # for Q21 in the frequency tables.
    # Q21 has its own nested branch: Yes→Q22 (open text), No/Don't specify→Q23.
    # Q22 is open text so no rule needed for it. Q23 is shown to all.
    {
        "child_patterns": [
            r"what messages do you typically give.+patients who are pregnant",
            r"what messages do you typically give.+patients who are breastfeeding",
            r"does the message given to patients.+differ whether it is medical",
        ],
        "parent_pattern": r"do you ever talk to patients about cannabis use",
        "parent_value":   "Yes",
    },
]

def ordered_for_col(col):
    if any(re.search(p, col, re.IGNORECASE) for p in LIKERT_PATTERNS):
        return ORDERED_LIKERT
    if any(re.search(p, col, re.IGNORECASE) for p in CONFIDENCE_PATTERNS):
        return ORDERED_CONFIDENCE
    return None

def _branching_eligibility(df_sub, question_label):
    """Return (boolean_mask, n) for respondents eligible for question_label.
    Mask is None and n = len(df_sub) when no branching rule applies."""
    for rule in BRANCHING_RULES:
        if any(re.search(p, question_label, re.IGNORECASE) for p in rule["child_patterns"]):
            parent_col = next(
                (c for c in df_sub.columns
                 if re.search(rule["parent_pattern"], c, re.IGNORECASE)
                 and "(choice=" not in c),
                None
            )
            if parent_col:
                col_lower = df_sub[parent_col].str.strip().str.lower()
                non_empty = col_lower[col_lower != ""]
                if len(non_empty) == 0:
                    return None, len(df_sub)
                pv = rule["parent_value"]
                if isinstance(pv, list):
                    mask = col_lower.isin([v.lower() for v in pv])
                else:
                    mask = col_lower == pv.lower()
                n = int(mask.sum())
                if n == 0:
                    return None, 0
                return mask, n
    return None, len(df_sub)


def get_branching_denom(df_sub, question_label):
    """Return parent-eligible count if question_label matches a branching rule, else None."""
    for rule in BRANCHING_RULES:
        if any(re.search(p, question_label, re.IGNORECASE) for p in rule["child_patterns"]):
            parent_col = next(
                (c for c in df_sub.columns
                 if re.search(rule["parent_pattern"], c, re.IGNORECASE)
                 and "(choice=" not in c),
                None
            )
            if parent_col:
                col_lower = df_sub[parent_col].str.strip().str.lower()
                non_empty = col_lower[col_lower != ""]
                if len(non_empty) == 0:
                    # Parent column all blank for this group (e.g. ineligible
                    # respondents never saw this question) — skip silently.
                    return None
                pv = rule["parent_value"]
                if isinstance(pv, list):
                    mask = col_lower.isin([v.lower() for v in pv])
                else:
                    mask = col_lower == pv.lower()
                n = int(mask.sum())
                if n == 0:
                    pv_display = pv if isinstance(pv, str) else " / ".join(pv)
                    actual = sorted(non_empty.unique().tolist())
                    print(f"  WARNING: branching rule matched '{short_q(question_label, 60)}' "
                          f"but parent_value '{pv_display}' found 0 rows.\n"
                          f"    Parent column: {short_q(parent_col, 70)}\n"
                          f"    Actual values: {actual[:10]}\n"
                          f"    → update BRANCHING_RULES parent_value to match")
                    return None
                return n
    return None

print(f"\nColumn types detected:")
print(f"  Checkbox groups      : {len(checkbox_groups)}")
print(f"  Single-choice        : {len(single_cols)}")
print(f"  Free-text (skipped)  : {len(free_text_cols)}")

if checkbox_groups:
    print(f"\nCheckbox groups found:")
    for parent, cols in checkbox_groups.items():
        print(f"  [{len(cols):2d} choices] {short_q(parent, 90)}")

# ── Split eligible vs ineligible ──────────────────────────────────────────────
elig_col = next(
    (c for c in df.columns
     if re.search(r'prenatal.*delivery.*postpartum', c, re.IGNORECASE)
     or re.search(r'provide prenatal', c, re.IGNORECASE)),
    None
)

days_col = next(
    (c for c in df.columns
     if re.search(r'how many days.*see patients|average week.*how many', c, re.IGNORECASE)),
    None
)

if elig_col:
    blank_elig = df[elig_col].str.strip() == ""
    said_yes   = df[elig_col].str.strip().str.lower() == "yes"

    # Exclude respondents who answered "0 days" on the follow-up days-per-week
    # question — seeing 0 patients means they do not meet the clinical threshold.
    zero_days  = pd.Series(False, index=df.index)
    blank_days = pd.Series(False, index=df.index)
    if days_col:
        days_stripped = df[days_col].str.strip()
        zero_days  = days_stripped.str.lower().str.match(r'^0\s*days?$', na=False)
        blank_days = days_stripped == ""
        n_zero = zero_days.sum()
        if n_zero > 0:
            print(f"\nExcluding {n_zero} respondent(s) who said 'Yes' to prenatal care "
                  f"but answered '0 days' on: {short_q(days_col, 70)}")

    # Incomplete: left Q1 blank (answered no questions at all).
    # Ineligible: answered Q1="No", OR Q1="Yes" but Q2 blank or "0 days".
    # Eligible:   Q1="Yes" AND Q2 is a non-zero, non-blank value.
    if days_col:
        incomplete_mask = blank_elig
        eligible_mask   = said_yes & ~zero_days & ~blank_days
        ineligible_mask = (~said_yes & ~blank_elig) | zero_days | (said_yes & blank_days)
    else:
        incomplete_mask = blank_elig
        eligible_mask   = said_yes
        ineligible_mask = ~said_yes & ~blank_elig

    incomplete = df[ incomplete_mask].reset_index(drop=True)
    eligible   = df[ eligible_mask  ].reset_index(drop=True)
    ineligible = df[ ineligible_mask].reset_index(drop=True)

    print(f"\nEligibility split on: {short_q(elig_col, 80)}")
    if days_col:
        print(f"  + '0 days' exclusion on: {short_q(days_col, 70)}")
    print(f"  Eligible   (full survey)              : {len(eligible)}")
    print(f"  Ineligible (screener + demo only)     : {len(ineligible)}")
    print(f"  Incomplete (did not finish screener)  : {len(incomplete)} — excluded from all tables")
else:
    eligible   = df.copy()
    ineligible = pd.DataFrame(columns=df.columns)
    print("\nWARNING: eligibility column not found — treating all records as eligible")

# ── Analytic sample: completers only ─────────────────────────────────────────
# Eligible respondents who actually submitted the survey (Survey_complete = "Complete").
# Partial completers (started but never hit Submit) are excluded from all
# frequency tables and cross-tabulations.
_survey_complete_col = next(
    (c for c in eligible.columns
     if re.search(r'^survey_complete$', c.strip(), re.IGNORECASE)),
    None
)
if _survey_complete_col:
    _raw_completers = eligible[
        eligible[_survey_complete_col].str.strip().str.lower() == "complete"
    ].reset_index(drop=True)
    # Exclude "ghost completers" — records marked Complete but with fewer than
    # _MIN_MAIN_Q main-survey questions answered (data quality exclusion).
    _screener_col_set = {c for c in (elig_col, days_col) if c}
    _main_q_cols = [
        c for c in _raw_completers.columns
        if c not in _screener_col_set
        and not is_demo_col(c)
        and c not in free_text_cols
        and c not in system_cols
        and c not in county_skip_cols
        and not COMPLETE_RE.search(c.strip())
    ]
    _MIN_MAIN_Q = 2   # must answer at least this many main-survey questions
    if _main_q_cols:
        _n_main = _raw_completers[_main_q_cols].apply(
            lambda c: c.str.strip().ne("")
        ).sum(axis=1)
        completers = _raw_completers[_n_main >= _MIN_MAIN_Q].reset_index(drop=True)
        n_ghost = len(_raw_completers) - len(completers)
    else:
        completers = _raw_completers
        n_ghost = 0
    n_partial = len(eligible) - len(completers)
    print(f"\nAnalytic sample (Survey_complete = Complete): N = {len(completers)} "
          f"({len(eligible) - len(_raw_completers)} partial completes excluded"
          + (f", {n_ghost} near-blank ghost completer(s) excluded)" if n_ghost else ")"))
else:
    completers = eligible.copy()
    print(f"\nWARNING: Survey_complete column not found — "
          f"analytic sample = all eligible (N={len(eligible)})")

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

def _find_ts_col(role, pattern):
    """Find a timestamp column: named pattern takes priority, then TIMESTAMP_COL_OVERRIDES."""
    found = _find_col(df.columns, pattern)
    if found:
        return found
    override = TIMESTAMP_COL_OVERRIDES.get(role)
    return override if (override and override in df.columns) else None

start_col      = _find_ts_col("screener_start", r'^screener_start$')
end_elig_col   = _find_ts_col("survey_end",     r'^survey_end$')
end_inelig_col = _find_ts_col("inel_demo_end",  r'^inel_demo_end$')

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

    start    = pd.to_datetime(grp_df[start_col], errors="coerce", format="mixed")
    end      = pd.to_datetime(grp_df[end_col],   errors="coerce", format="mixed")
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
def freq_single(series, question_text, ordered=None, n_denom=None, include_missing=False):
    """Frequency table for a single-choice column.

    n_denom: override denominator (use for branching questions where eligible
             pool > those who answered).  When None, denominator = answered.
    include_missing: if True and n_denom is set, append a Missing (skipped) row.
    """
    vals = series.apply(norm_val).dropna()
    n_answered = len(vals)
    if n_answered == 0 and not (n_denom and n_denom > 0):
        return None
    denom = n_denom if n_denom is not None else n_answered
    counts = vals.value_counts(dropna=True)
    if ordered:
        idx = [o for o in ordered if o in counts.index]
        idx += [v for v in counts.index if v not in idx]
        counts = counts.reindex(idx).dropna()
    pct = (counts / denom * 100).round(1)
    out = pd.DataFrame({"Response": counts.index, "n": counts.values, "%": pct.values})
    out.insert(0, "Question", re.sub(r'\s+', ' ', str(question_text)).strip())
    out.insert(1, "Type", "Single choice")
    out["N (denominator)"] = denom
    if include_missing and n_denom is not None:
        n_missing = denom - n_answered
        if n_missing > 0:
            missing_row = pd.DataFrame([{
                "Question":        re.sub(r'\s+', ' ', str(question_text)).strip(),
                "Type":            "Single choice",
                "Response":        "Missing (skipped)",
                "n":               n_missing,
                "%":               round(n_missing / denom * 100, 1),
                "N (denominator)": denom,
            }])
            out = pd.concat([out, missing_row], ignore_index=True)
    return out

def freq_checkbox_group(df_sub, parent_q, child_cols,
                         n_denom_override=None, include_missing=False):
    """Frequency table for a checkbox group.

    Denominator = rows where at least one child column is checked (Checked/1/Yes).
    Falls back to non-empty string check for exports that don't use Checked/1
    (e.g. free-text positive values).  This order correctly excludes REDCap's
    "Unchecked" sentinel from the denominator.

    include_missing: if True and n_denom_override is set, append a Missing
                     (skipped) row for those in the eligible pool who left the
                     entire question blank.
    """
    # Primary: rows with at least one positively-checked option
    has_response = df_sub[child_cols].apply(
        lambda col: col.apply(is_checked)
    ).any(axis=1)
    n_answered = has_response.sum()

    if n_answered == 0:
        # Fallback: non-empty value that isn't an explicit negative sentinel
        has_response = df_sub[child_cols].apply(
            lambda col: col.str.strip().str.lower().isin(
                ["", "nan", "unchecked", "0", "false", "no"]
            ).eq(False)
        ).any(axis=1)
        n_answered = has_response.sum()
        if n_answered == 0:
            print(f"  WARNING: skipping checkbox group (all blank): "
                  f"{short_q(parent_q, 70)}")
            return None

    n_denom = n_denom_override if n_denom_override is not None else n_answered

    # Group child columns by their choice label so pandas dedup variants
    # (.1, .2, …) are combined rather than emitted as separate response rows.
    label_cols: dict[str, list[str]] = {}
    for col in child_cols:
        label_cols.setdefault(choice_label(col), []).append(col)

    rows = []
    for label, cols in label_cols.items():
        n_checked = sum(
            int(df_sub.loc[has_response, c].apply(is_checked).sum())
            for c in cols if c in df_sub.columns
        )
        rows.append({
            "Question":        re.sub(r'\s+', ' ', str(parent_q)).strip(),
            "Type":            "Select all that apply",
            "Response":        label,
            "n":               n_checked,
            "%":               round(n_checked / n_denom * 100, 1),
            "N (denominator)": int(n_denom),
        })

    if include_missing and n_denom_override is not None:
        n_missing = int(n_denom) - int(n_answered)
        if n_missing > 0:
            rows.append({
                "Question":        re.sub(r'\s+', ' ', str(parent_q)).strip(),
                "Type":            "Select all that apply",
                "Response":        "Missing (skipped)",
                "n":               n_missing,
                "%":               round(n_missing / n_denom * 100, 1),
                "N (denominator)": int(n_denom),
            })

    return pd.DataFrame(rows)

def free_text_for_question(df_sub, q_key):
    """Return a DataFrame of verbatim 'Other, please specify' responses for q_key, or None.

    For SATA questions, only includes rows where the corresponding 'Other' checkbox
    is checked — prevents dumping all respondents when only a few selected Other.
    """
    ft_cols = [c for c in q_free_text_cols.get(q_key, []) if c in df_sub.columns]
    if not ft_cols:
        return None

    # For SATA questions find the "Other, please specify" checkbox column so we
    # can filter to only respondents who actually selected that option.
    other_cb_col = None
    if q_key in checkbox_groups:
        for cb_col in checkbox_groups[q_key]:
            if (re.search(r'\bother\b', cb_col, re.IGNORECASE) and
                    re.search(r'\bspecify\b', cb_col, re.IGNORECASE)):
                other_cb_col = cb_col
                break
        if other_cb_col is None:
            for cb_col in checkbox_groups[q_key]:
                if re.search(r'\bother\b', cb_col, re.IGNORECASE):
                    other_cb_col = cb_col
                    break

    rows = []
    for ft_col in ft_cols:
        sub = (df_sub[df_sub[other_cb_col].apply(is_checked)]
               if other_cb_col and other_cb_col in df_sub.columns
               else df_sub)
        for val in sub[ft_col]:
            v = str(val).strip()
            if v and v.lower() not in ("nan", "", "checked", "unchecked", "0", "1"):
                rows.append({
                    "Question": re.sub(r'\s+', ' ', str(q_key)).strip(),
                    "Type":     "Open text",
                    "Response": v,
                })
    return pd.DataFrame(rows) if rows else None


def _preferred_single_cols(df_sub):
    """
    For single-choice columns that share the same resolved question label in
    df_sub, return the set of column names to USE — one per unique label,
    chosen as the column with the most non-empty values.

    Columns whose label matches a _DEMO_AND_SCREENER_PATTERNS pattern are
    grouped by their matched pattern (not by exact label).  This handles the
    common REDCap export case where the same demographic question lives in both
    the ineligible-demographics instrument and the main survey instrument under
    slightly different column names: for completers, the main-survey column
    (which has real answers) wins; for ineligibles, their form's column wins.

    Checkbox columns are unaffected (deduplicated via parent-question tracking).
    """
    label_best: dict[str, tuple[str, int]] = {}  # key → (col, n_answered)
    checkbox_set = set(checkbox_cols)
    for col in df_sub.columns:
        if col in free_text_cols or col in system_cols or col in county_skip_cols:
            continue
        if col in checkbox_set:
            continue
        label = complete_col_labels.get(col, col)
        n = int(df_sub[col].astype(str).str.strip().ne("").sum())
        # Group demographic questions by matched pattern so that the same concept
        # appearing in multiple instruments resolves to one winner per concept.
        key = label
        for p in _DEMO_AND_SCREENER_PATTERNS:
            if re.search(p, label, re.IGNORECASE):
                key = p
                break
        if key not in label_best or n > label_best[key][1]:
            label_best[key] = (col, n)
    return {col for col, _ in label_best.values()}


def build_all_frequencies(df_sub):
    """
    Build one combined frequency table for df_sub, processing columns in the
    order they appear in the original data (single-choice and checkbox groups
    are interleaved to match the survey flow).
    """
    parts = []
    emitted_checkbox_parents = set()
    checkbox_set = set(checkbox_cols)   # O(1) membership test
    preferred = _preferred_single_cols(df_sub)   # one col per label (most answers wins)

    for col in df_sub.columns:
        if col in free_text_cols or col in system_cols or col in county_skip_cols:
            continue

        if col in checkbox_set:
            # Emit the whole checkbox group when we hit the first child column
            parent = parent_question(col)
            if parent in emitted_checkbox_parents:
                continue
            emitted_checkbox_parents.add(parent)
            child_cols = [c for c in checkbox_groups.get(parent, []) if c in df_sub.columns]
            if child_cols:
                n_br = get_branching_denom(df_sub, parent)
                t = freq_checkbox_group(df_sub, parent, child_cols,
                                        n_denom_override=n_br,
                                        include_missing=(n_br is not None))
                if t is not None:
                    parts.append(t)
                    ft = free_text_for_question(df_sub, parent)
                    if ft is not None:
                        parts.append(ft)
        else:
            if col not in preferred:
                continue   # another column for this label has more answers — skip
            # Single-choice — use disambiguated label for Complete? columns
            label   = complete_col_labels.get(col, col)
            ordered = ordered_for_col(col)
            n_br    = get_branching_denom(df_sub, label)
            t = freq_single(df_sub[col], label, ordered,
                            n_denom=n_br, include_missing=(n_br is not None))
            if t is not None:
                parts.append(t)
                ft = free_text_for_question(df_sub, label)
                if ft is not None:
                    parts.append(ft)

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


_DEMO_AND_SCREENER_PATTERNS = [
    r"provide prenatal|prenatal.*delivery.*postpartum",
    r"how many days.*see patients|average week.*how many",
    r"what is your profession",
    r"do you primarily see pregnant",
    r"how long have you been practicing",
    r"what is your primary specialty",
    r"primary specialty",
    r"what is your secondary or sub.specialty",
    r"secondary.{0,30}specialty",
    r"sub.specialty",
    r"which of the following best describes your primary practice setting",
    r"how would you describe the insurance status",
    r"what is your gender",
    r"what is your age",
    r"what is your race",
    r"which county",
]


def detect_started_n(df_sub, screener_cols=None):
    """Count eligible respondents who answered at least one full-survey question.
    Excludes screener, demographic, system, and free-text columns so that
    respondents who only completed the screener/demographics section (and never
    opened the main survey) are not counted as starters."""
    skip = (free_text_cols | system_cols | county_skip_cols |
            (screener_cols or set()))
    main_cols = [
        c for c in df_sub.columns
        if c not in skip
        and not any(re.search(p, c, re.IGNORECASE) for p in _DEMO_AND_SCREENER_PATTERNS)
    ]
    if not main_cols:
        return len(df_sub)
    answered_any = df_sub[main_cols].apply(
        lambda c: c.str.strip().ne("")
    ).any(axis=1)
    n = int(answered_any.sum())
    print(f"  detect_started_n: {len(main_cols)} main-survey cols checked, "
          f"{n} respondents answered at least one")
    return n


def build_missingness(df_sub, started_n=None):
    """
    One row per survey question.  Call twice — once without started_n (uses
    len(df_sub) as the base) and once with started_n (e.g. 97) to produce a
    version that treats the non-starters as a separate 'never started' group
    rather than lumping them into 'not asked (branched)'.

    Columns:
      Question | Type | N total eligible | N never started | N not asked (branched) |
      N routed to question | N answered | N skipped (true missing) | % skipped | Flag

    'N never started'    : eligible respondents who never opened the main survey
                           (only present when started_n is supplied).
    'N not asked (branched)': respondents routed away by branch logic — not missing.
    'N skipped (true missing)': routed but gave no response — true missingness.
    Flag: 'NOT IN THIS FORM' when N answered = 0 with no branching (wrong instrument);
          'FLAG >20%' when % skipped > 20 among those routed.

    Output sorted by % skipped descending.
    """
    rows = []
    emitted = set()
    checkbox_set = set(checkbox_cols)
    total_n = len(df_sub)
    # When started_n is provided, unbranched questions use it as their denominator
    # (the 9 non-starters are surfaced as 'N never started', not as skips).
    base_n = started_n if started_n is not None else total_n
    n_never_started = (total_n - started_n) if started_n is not None else 0
    preferred = _preferred_single_cols(df_sub)

    for col in df_sub.columns:
        if col in free_text_cols or col in system_cols or col in county_skip_cols:
            continue

        if col in checkbox_set:
            parent = parent_question(col)
            if parent in emitted:
                continue
            emitted.add(parent)
            child_cols = [c for c in checkbox_groups.get(parent, []) if c in df_sub.columns]
            if not child_cols:
                continue
            mask, n_routed_raw = _branching_eligibility(df_sub, parent)
            # Cap routed N at base_n so non-starters don't inflate it
            n_routed = min(n_routed_raw, base_n) if mask is not None else base_n
            elig_sub = df_sub[mask] if mask is not None else df_sub
            has_any = elig_sub[child_cols].apply(lambda c: c.apply(is_checked)).any(axis=1)
            n_answered = int(has_any.sum())
            q_label = re.sub(r'\s+', ' ', str(parent)).strip()
            q_type = "Select all that apply"
        else:
            if col not in preferred:
                continue   # another column for this label has more answers — skip
            label = complete_col_labels.get(col, col)
            mask, n_routed_raw = _branching_eligibility(df_sub, label)
            n_routed = min(n_routed_raw, base_n) if mask is not None else base_n
            elig_sub = df_sub[mask] if mask is not None else df_sub
            n_answered = int(elig_sub[col].str.strip().ne("").sum())
            q_label = re.sub(r'\s+', ' ', str(label)).strip()
            q_type = "Single choice / Likert"

        n_not_asked = base_n - n_routed   # branched away — not true missing
        n_skipped   = max(0, n_routed - n_answered)
        pct_skipped = round(n_skipped / n_routed * 100, 1) if n_routed > 0 else None

        row = {
            "Question":                 q_label,
            "Type":                     q_type,
            "N total eligible":         base_n,
            "N routed to question":     n_routed,
            "N answered":               n_answered,
            "N skipped (true missing)": n_skipped,
            "% skipped":                pct_skipped,
            "Flag": ("NOT IN THIS FORM" if n_answered == 0 and n_not_asked == 0
                                            and n_never_started == 0
                     else "FLAG >20%"   if (pct_skipped or 0) > 20
                     else ""),
        }
        # Insert contextual N columns in a readable order
        if started_n is not None:
            row["N never started"] = n_never_started
        row["N not asked (branched)"] = n_not_asked
        rows.append(row)

    df_out = pd.DataFrame(rows) if rows else pd.DataFrame()
    # Reorder columns for readability
    col_order = (
        ["Question", "Type", "N total eligible"]
        + (["N never started"] if started_n is not None else [])
        + ["N not asked (branched)", "N routed to question",
           "N answered", "N skipped (true missing)", "% skipped", "Flag"]
    )
    df_out = df_out.reindex(columns=[c for c in col_order if c in df_out.columns])
    if not df_out.empty and "% skipped" in df_out.columns:
        df_out = df_out.sort_values("% skipped", ascending=False, na_position="last"
                                    ).reset_index(drop=True)
    return df_out


def build_respondent_completeness(df_sub):
    """
    Respondent-level completion analysis.

    For each respondent, counts:
      N questions routed   — questions this person was eligible to answer
      N answered           — of those, how many they answered
      N skipped            — routed but no response (true missing)

    Skips columns where nobody in df_sub answered (wrong-instrument columns)
    and demographic / screener columns so the count reflects main-survey
    questions only.

    Returns (summary_df, per_respondent_df):
      summary_df       — distribution of skip counts across respondents
      per_respondent_df— one row per respondent
    """
    total_n = len(df_sub)
    n_routed_arr  = np.zeros(total_n, dtype=int)
    n_skipped_arr = np.zeros(total_n, dtype=int)

    emitted = set()
    checkbox_set = set(checkbox_cols)
    preferred = _preferred_single_cols(df_sub)
    n_questions = 0

    for col in df_sub.columns:
        if col in free_text_cols or col in system_cols or col in county_skip_cols:
            continue

        if col in checkbox_set:
            parent = parent_question(col)
            if parent in emitted:
                continue
            emitted.add(parent)
            child_cols = [c for c in checkbox_groups.get(parent, []) if c in df_sub.columns]
            if not child_cols:
                continue
            answered_series = df_sub[child_cols].apply(
                lambda c: c.apply(is_checked)
            ).any(axis=1)
            if answered_series.sum() == 0:
                continue   # nobody answered — wrong instrument
            mask, _ = _branching_eligibility(df_sub, parent)
            routed_series = mask if mask is not None else pd.Series(True, index=df_sub.index)

        else:
            if col not in preferred:
                continue   # another column for this label has more answers — skip
            label = complete_col_labels.get(col, col)
            if any(re.search(p, label, re.IGNORECASE) for p in _DEMO_AND_SCREENER_PATTERNS):
                continue
            answered_series = df_sub[col].str.strip().ne("")
            if answered_series.sum() == 0:
                continue   # nobody answered — wrong instrument
            mask, _ = _branching_eligibility(df_sub, label)
            routed_series = mask if mask is not None else pd.Series(True, index=df_sub.index)

        routed   = routed_series.reindex(df_sub.index).fillna(False).to_numpy(dtype=bool)
        answered = answered_series.reindex(df_sub.index).fillna(False).to_numpy(dtype=bool)

        n_routed_arr  += routed.astype(int)
        n_skipped_arr += (routed & ~answered).astype(int)
        n_questions   += 1

    print(f"  Respondent completeness: {n_questions} questions assessed, "
          f"{total_n} respondents")

    n_answered_arr = n_routed_arr - n_skipped_arr
    per_resp = pd.DataFrame({
        "N questions routed": n_routed_arr,
        "N answered":         n_answered_arr,
        "N skipped":          n_skipped_arr,
        "% of routed answered": np.where(
            n_routed_arr > 0,
            np.round(n_answered_arr / n_routed_arr * 100, 1),
            np.nan
        ),
    })

    # Distribution: how many respondents had 0 skips, 1 skip, etc.
    skip_dist = (
        per_resp["N skipped"]
        .value_counts()
        .sort_index()
        .reset_index()
    )
    skip_dist.columns = ["N questions skipped", "N respondents"]
    skip_dist["% of respondents"] = round(
        skip_dist["N respondents"] / total_n * 100, 1
    )
    skip_dist["Cumulative N"] = skip_dist["N respondents"].cumsum()
    skip_dist["Cumulative %"] = round(
        skip_dist["Cumulative N"] / total_n * 100, 1
    )

    return skip_dist, per_resp


def build_mcar_analysis(df_sub, min_pct_missing=5.0):
    """
    Test whether question-level missingness is associated with respondent
    demographics (MCAR = Missing Completely At Random analysis).

    For each main-survey question where ≥min_pct_missing% of routed respondents
    gave no response, builds a contingency table of [answered vs skipped] ×
    [demographic category] and runs chi-square or Fisher's exact.

    Demographic variables tested: Gender, Age, Profession, Tenure,
    Practice setting (whichever columns are present in df_sub).

    Returns a DataFrame sorted significant-first, then by p-value.
    Columns: Question | % skipped | N skipped | N routed |
             Demographic var | Test | p-value | Significant (p<.05)
    """
    if not _SCIPY:
        print("  MCAR analysis skipped — scipy not installed")
        return pd.DataFrame()

    DEMO_TEST_PATTERNS = [
        (r"what is your gender",                                              "Gender"),
        (r"what is your age",                                                 "Age"),
        (r"what is your profession",                                          "Profession"),
        (r"how long have you been practicing",                                "Tenure"),
        (r"which of the following best describes your primary practice setting", "Practice setting"),
    ]
    demo_cols = []
    for pattern, display_name in DEMO_TEST_PATTERNS:
        col = next(
            (c for c in df_sub.columns
             if re.search(pattern, c, re.IGNORECASE) and "(choice=" not in c),
            None,
        )
        if col:
            demo_cols.append((col, display_name))

    if not demo_cols:
        print("  MCAR analysis: no demographic columns found")
        return pd.DataFrame()

    print(f"  MCAR analysis: testing against {len(demo_cols)} demographic variable(s)")

    rows = []
    emitted = set()
    checkbox_set = set(checkbox_cols)
    preferred = _preferred_single_cols(df_sub)

    for col in df_sub.columns:
        if col in free_text_cols or col in system_cols or col in county_skip_cols:
            continue

        if col in checkbox_set:
            parent = parent_question(col)
            if parent in emitted:
                continue
            emitted.add(parent)
            child_cols = [c for c in checkbox_groups.get(parent, []) if c in df_sub.columns]
            if not child_cols:
                continue
            q_label = re.sub(r'\s+', ' ', str(parent)).strip()
            if any(re.search(p, q_label, re.IGNORECASE) for p in _DEMO_AND_SCREENER_PATTERNS):
                continue
            mask, _ = _branching_eligibility(df_sub, parent)
            elig_sub = df_sub[mask] if mask is not None else df_sub
            answered_series = elig_sub[child_cols].apply(
                lambda c: c.apply(is_checked)
            ).any(axis=1)
            if answered_series.sum() == 0:
                continue  # wrong instrument — nobody answered
        else:
            if col not in preferred:
                continue   # another column for this label has more answers — skip
            label = complete_col_labels.get(col, col)
            if any(re.search(p, label, re.IGNORECASE) for p in _DEMO_AND_SCREENER_PATTERNS):
                continue
            mask, _ = _branching_eligibility(df_sub, label)
            q_label = re.sub(r'\s+', ' ', str(label)).strip()
            elig_sub = df_sub[mask] if mask is not None else df_sub
            answered_series = elig_sub[col].str.strip().ne("")
            if answered_series.sum() == 0:
                continue  # wrong instrument

        n_routed   = len(elig_sub)
        n_answered = int(answered_series.sum())
        n_skipped  = n_routed - n_answered
        pct_skipped = round(n_skipped / n_routed * 100, 1) if n_routed > 0 else 0.0

        if pct_skipped < min_pct_missing:
            continue

        # Binary missing indicator (index = elig_sub.index)
        missing_indicator = (~answered_series.astype(bool)).astype(int)
        routed_idx = elig_sub.index

        for demo_col, demo_name in demo_cols:
            if demo_col not in df_sub.columns:
                continue
            routed_demo = df_sub.loc[routed_idx, demo_col].str.strip()
            has_demo    = routed_demo.ne("")
            valid_idx   = routed_demo[has_demo].index
            if len(valid_idx) < 4:
                continue

            demo_sub    = routed_demo.loc[valid_idx]
            missing_sub = missing_indicator.reindex(valid_idx).fillna(0).astype(int)
            categories  = sorted(demo_sub.unique())
            if len(categories) < 2:
                continue

            table = []
            for cat in categories:
                cat_mask = demo_sub == cat
                n_miss   = int(missing_sub[cat_mask].sum())
                n_ans    = int(cat_mask.sum()) - n_miss
                table.append([n_ans, n_miss])

            p_val, test_name, _, _ = _stat_test(table)
            if p_val is None:
                continue

            rows.append({
                "Question":             q_label,
                "% skipped":            pct_skipped,
                "N skipped":            n_skipped,
                "N routed":             n_routed,
                "Demographic var":      demo_name,
                "Test":                 test_name,
                "p-value":              round(p_val, 4),
                "Significant (p<.05)": "Yes" if p_val < 0.05 else "No",
            })

    if not rows:
        print(f"  MCAR analysis: no questions with ≥{min_pct_missing}% missingness found")
        return pd.DataFrame()

    mcar_df = pd.DataFrame(rows).sort_values(
        ["Significant (p<.05)", "p-value"],
        ascending=[False, True],  # "Yes" before "No"; smallest p first within each group
    ).reset_index(drop=True)
    n_sig = int((mcar_df["Significant (p<.05)"] == "Yes").sum())
    print(f"  MCAR analysis: {len(mcar_df)} tests, {n_sig} significant at p < .05")
    return mcar_df


def _col_scores(series, score_map):
    """Map a raw response Series to numeric Likert/confidence scores.
    Values absent from score_map (e.g. 'Don't know') become NaN and are excluded."""
    def to_num(v):
        if pd.isna(v) or str(v).strip() in ("", "nan"):
            return np.nan
        key = str(v).strip().lower().replace("’", "'")
        return score_map.get(key, np.nan)
    return pd.to_numeric(series.apply(to_num), errors="coerce").dropna()


def _two_prop_z(n1, N1, n0, N0):
    """Two-proportion Z-test. Returns two-tailed p-value or None if undefined."""
    if N1 == 0 or N0 == 0:
        return None
    p_pool = (n1 + n0) / (N1 + N0)
    if p_pool in (0.0, 1.0):
        return None
    se = (p_pool * (1 - p_pool) * (1 / N1 + 1 / N0)) ** 0.5
    if se == 0:
        return None
    z = (n1 / N1 - n0 / N0) / se
    return float(2 * (1 - _norm.cdf(abs(z))))


def _holm_correct(p_values):
    """Holm step-down correction. Input list may contain None (skipped).
    Returns list of adjusted p-values in the same order."""
    valid = [(i, p) for i, p in enumerate(p_values) if p is not None]
    if not valid:
        return list(p_values)
    k = len(valid)
    adj = {}
    cummax = 0.0
    for rank, (i, p) in enumerate(sorted(valid, key=lambda x: x[1])):
        a = max(min(p * (k - rank), 1.0), cummax)
        adj[i] = round(a, 4)
        cummax = a
    result = list(p_values)
    for i, a in adj.items():
        result[i] = a
    return result


def _stat_test(table):
    """
    Return (p, test_name, stat_value, stat_df) for a contingency table.
    - 2×2 with any cell < 5 → Fisher's exact (stat_value=None, stat_df=None)
    - all other cases → chi-square (stat_value=χ², stat_df=degrees of freedom)
    - degenerate table → (None, None, None, None)
    """
    if not _SCIPY:
        return None, None, None, None
    rows = [[int(a), int(b)] for a, b in table if (a + b) > 0]
    if len(rows) < 2:
        return None, None, None, None
    col_sums = [sum(r[i] for r in rows) for i in range(2)]
    if any(s == 0 for s in col_sums):
        return None, None, None, None
    flat = [v for row in rows for v in row]
    if len(rows) == 2 and min(flat) < 5:
        # Fisher's exact is valid even with zero cells
        _, p = fisher_exact(rows)
        return float(p), "Fisher", None, None
    stat, p, dof, expected = chi2_contingency(rows)
    note = " (small n)" if expected.min() < 5 else ""
    return float(p), f"χ²{note}", round(float(stat), 2), int(dof)


def _add_pvalues(df, N1, N0, col_n1, col_n0, col_N1=None, col_N0=None):
    """
    Append statistical test columns to a cross-tab DataFrame.

    Single-choice: omnibus chi-square / Fisher's on the full m×2 table.
      p_value and test are placed on the first response row.
      When omnibus p < 0.05 and chi-square was used, adjusted standardised
      residuals are computed for each row; their two-tailed p-values are
      Holm-corrected and stored in post_hoc_p alongside adj_residual.

    Select-all-that-apply: an omnibus chi-square / Fisher's exact is run on the
      full k×2 table (one row per option) as the significance gate.  Only when
      that omnibus p < 0.05 are per-option two-proportion Z-tests computed and
      Holm-corrected into post_hoc_p.

    Columns added: p_value, test, stat_value, stat_df, adj_residual, post_hoc_p
    """
    if not _SCIPY:
        return df
    df = df.copy()
    for col in ("p_value", "stat_value", "stat_df", "adj_residual", "post_hoc_p"):
        df[col] = np.nan   # float NaN; avoids pd.NA dtype coercion issues in pandas 2.x
    df["test"] = None

    for (q, qtype), grp in df.groupby(["Question", "Type"], sort=False):
        data = grp[grp["Response"] != "Missing (skipped)"]
        n1s = pd.to_numeric(data[col_n1], errors="coerce").fillna(0).astype(int).tolist()
        n0s = pd.to_numeric(data[col_n0], errors="coerce").fillna(0).astype(int).tolist()

        # Use question-specific branching-aware N if provided (cross-tab), else fixed totals
        q_N1, q_N0 = N1, N0
        if col_N1 and col_N1 in data.columns:
            v = pd.to_numeric(data[col_N1], errors="coerce").dropna()
            if len(v) > 0: q_N1 = int(v.iloc[0])
        if col_N0 and col_N0 in data.columns:
            v = pd.to_numeric(data[col_N0], errors="coerce").dropna()
            if len(v) > 0: q_N0 = int(v.iloc[0])

        table = list(zip(n1s, n0s))
        p_omni, tname, s_val, s_df = _stat_test(table)

        for i, idx in enumerate(data.index):
            df.at[idx, "p_value"]   = p_omni if i == 0 else pd.NA
            df.at[idx, "test"]      = tname  if i == 0 else pd.NA
            df.at[idx, "stat_value"] = s_val if i == 0 else pd.NA
            df.at[idx, "stat_df"]   = s_df   if i == 0 else pd.NA

        if qtype == "Select all that apply":
            # Per-option Holm-corrected two-proportion Z-tests — always run for SATA.
            # SATA responses are not mutually exclusive so the omnibus chi-square is
            # conceptually inappropriate as a gate; per-option tests are the primary analysis.
            raw_ps = [_two_prop_z(c1, q_N1, c0, q_N0) for c1, c0 in zip(n1s, n0s)]
            for idx, adj_p in zip(data.index, _holm_correct(raw_ps)):
                df.at[idx, "post_hoc_p"] = adj_p
        else:
            # Post-hoc adjusted residuals when omnibus chi-square is significant
            if p_omni is not None and p_omni < 0.05 and tname and tname.startswith("χ"):
                rows_v = [[int(a), int(b)] for a, b in table if (a + b) > 0]
                col_s  = [sum(r[i] for r in rows_v) for i in range(2)]
                if len(rows_v) >= 2 and not any(s == 0 for s in col_s):
                    obs      = np.array(rows_v, dtype=float)
                    _, _, _, expected = chi2_contingency(obs)
                    n_total  = obs.sum()
                    row_sums = obs.sum(axis=1)
                    cs0      = obs.sum(axis=0)[0]
                    adj_res  = []
                    for j in range(len(rows_v)):
                        o, e, rs = obs[j, 0], expected[j, 0], row_sums[j]
                        denom = (e * (1 - rs / n_total) * (1 - cs0 / n_total)) ** 0.5
                        adj_res.append(round(float((o - e) / denom), 3) if denom > 0 else None)
                    raw_ps = [float(2 * (1 - _norm.cdf(abs(r)))) if r is not None else None
                              for r in adj_res]
                    adj_ps = _holm_correct(raw_ps)
                    valid_idx = [idx for idx, (a, b) in zip(data.index, table) if (a + b) > 0]
                    for idx, res, ap in zip(valid_idx, adj_res, adj_ps):
                        df.at[idx, "adj_residual"] = res
                        df.at[idx, "post_hoc_p"]   = ap

    return df


def _score_map_for_col(col):
    """Return (score_map, null_median) for a Likert/confidence column, or (None, None)."""
    if any(re.search(p, col, re.IGNORECASE) for p in LIKERT_PATTERNS):
        return LIKERT_SCORE_MAP, 3          # neutral = "Neither agree nor disagree"
    if any(re.search(p, col, re.IGNORECASE) for p in CONFIDENCE_PATTERNS):
        return CONFIDENCE_SCORE_MAP, 2      # midpoint = "Somewhat confident"
    return None, None


def _add_likert_tests_crosstab(merged, g1, g0,
                               col_n1=None, col_n0=None, col_N1=None, col_N0=None):
    """For Likert/confidence questions, overwrite chi-square stats with Mann-Whitney U.
    Mann-Whitney with two groups is already the complete pairwise test — no post-hoc
    decomposition is needed or appropriate. post_hoc_p is left blank for these rows.
    'Don't know' responses are excluded from scoring."""
    if not _SCIPY:
        return merged
    for col in g1.columns:
        score_map, _ = _score_map_for_col(col)
        if score_map is None:
            continue
        label = complete_col_labels.get(col, col)
        q_mask = merged["Question"] == label
        if not q_mask.any():
            print(f"  [MW skip] no Question match for label: {short_q(label, 80)}")
            continue
        s1 = _col_scores(g1[col], score_map)
        s0 = _col_scores(g0[col], score_map)
        if len(s1) < 2 or len(s0) < 2:
            print(f"  [MW skip] too few scores — s1={len(s1)}, s0={len(s0)} for: {short_q(label, 80)}")
            continue
        try:
            u_stat, p = mannwhitneyu(s1, s0, alternative="two-sided")
            p = float(p)
            tname = "Mann-Whitney U"
        except Exception as _mw_err:
            print(f"  [MW skip] exception for {short_q(label, 80)}: {_mw_err}")
            continue
        q_idx = merged.index[q_mask].tolist()
        for i, idx in enumerate(q_idx):
            merged.at[idx, "p_value"]      = p                       if i == 0 else pd.NA
            merged.at[idx, "test"]         = tname                   if i == 0 else pd.NA
            merged.at[idx, "stat_value"]   = round(float(u_stat), 2) if i == 0 else pd.NA
            merged.at[idx, "stat_df"]      = pd.NA
            merged.at[idx, "adj_residual"] = pd.NA
            merged.at[idx, "post_hoc_p"]   = pd.NA
    return merged


def _add_likert_tests_main(freq_df, raw_df):
    """Add one-sample Wilcoxon signed rank test to the main frequency table for
    Likert/confidence questions.  Tests whether the median differs from the neutral
    midpoint (3 for 5-point Likert, 2 for 3-point confidence).  'Don't know'
    responses are excluded.  p_value is placed on the first response row."""
    if not _SCIPY:
        return freq_df
    freq_df = freq_df.copy()
    for col_name in ("p_value", "test", "stat_value"):
        if col_name not in freq_df.columns:
            freq_df[col_name] = pd.NA
    for col in raw_df.columns:
        score_map, null_median = _score_map_for_col(col)
        if score_map is None:
            continue
        scores = _col_scores(raw_df[col], score_map)
        if len(scores) < 5:
            continue
        label = complete_col_labels.get(col, col)
        q_mask = freq_df["Question"] == label
        if not q_mask.any():
            continue
        diffs = scores - null_median
        diffs = diffs[diffs != 0]       # Wilcoxon requires non-zero differences
        if len(diffs) < 5:
            continue
        try:
            w_stat, p = wilcoxon(diffs, alternative="two-sided")
            p = float(p)
        except Exception:
            continue
        q_idx = freq_df.index[q_mask].tolist()
        for i, idx in enumerate(q_idx):
            freq_df.at[idx, "p_value"]   = p                       if i == 0 else pd.NA
            freq_df.at[idx, "test"]      = "Wilcoxon"              if i == 0 else pd.NA
            freq_df.at[idx, "stat_value"] = round(float(w_stat), 2) if i == 0 else pd.NA
    return freq_df


def build_crosstab(df_with_group, group_col, label1, label0):
    """
    Split df_with_group on binary group_col (1 vs 0), run build_all_frequencies
    on each half, and merge side-by-side.

    % and N use the branching-aware denominator from build_all_frequencies so
    cross-tab percentages match the main eligible table and h-bar charts.
    Single-choice / SATA: chi-square or Fisher's exact + Holm post-hoc.
    Likert / Confidence: Mann-Whitney U; Don't Know excluded; no post-hoc.
    """
    g1 = df_with_group[df_with_group[group_col] == 1].drop(columns=[group_col]).reset_index(drop=True)
    g0 = df_with_group[df_with_group[group_col] == 0].drop(columns=[group_col]).reset_index(drop=True)

    if g1.empty or g0.empty:
        print(f"  WARNING: one group is empty for {group_col} — skipping cross-tab")
        return pd.DataFrame()

    N1, N0 = len(g1), len(g0)
    print(f"  {label1}: N={N1}  |  {label0}: N={N0}")

    f1 = build_all_frequencies(g1)
    f0 = build_all_frequencies(g0)

    if f1.empty or f0.empty:
        return pd.DataFrame()

    for f in (f1, f0):
        f["n"] = pd.to_numeric(f["n"], errors="coerce").fillna(0).astype(int)

    col_n1, col_n0 = f"n ({label1})", f"n ({label0})"
    col_N1, col_N0 = f"N ({label1})", f"N ({label0})"

    # Rename columns before merge; branching-aware N (denominator) becomes per-group N column
    f1 = f1.rename(columns={"n": col_n1, "%": f"% ({label1})", "N (denominator)": col_N1})
    f0 = f0.rename(columns={"n": col_n0, "%": f"% ({label0})", "N (denominator)": col_N0})
    # Drop any residual columns that would cause merge key collisions
    for fi in (f1, f0):
        fi.drop(columns=[c for c in fi.columns
                          if c not in ["Question", "Type", "Response"]
                          and c not in [col_n1, f"% ({label1})", col_N1,
                                        col_n0, f"% ({label0})", col_N0]],
                errors="ignore", inplace=True)

    merged = f1.merge(f0, on=["Question", "Type", "Response"], how="outer")

    # After an outer merge, rows present in only one side have NaN in "Type".
    # Groupby in _add_pvalues silently drops NaN-key groups, so SATA Z-tests
    # would never run for those rows.  Fill from the non-NaN sibling within
    # each Question block.
    if "Type" in merged.columns:
        merged["Type"] = (
            merged.groupby("Question", sort=False)["Type"]
            .transform(lambda x: x.ffill().bfill())
        )

    merged = _add_pvalues(merged, N1, N0, col_n1, col_n0, col_N1=col_N1, col_N0=col_N0)
    merged = _add_likert_tests_crosstab(merged, g1, g0)

    col_order = [
        "Question", "Type", "Response",
        col_n1, f"% ({label1})",
        col_n0, f"% ({label0})",
        f"N ({label1})", f"N ({label0})",
        "p_value", "test", "stat_value", "stat_df", "adj_residual", "post_hoc_p",
    ]
    return merged[[c for c in col_order if c in merged.columns]]

# ── Build sheets ──────────────────────────────────────────────────────────────
print("\nBuilding frequency tables ...")

elig_freqs   = build_all_frequencies(completers)
elig_freqs   = _add_likert_tests_main(elig_freqs, completers)
inelig_freqs = build_all_frequencies(ineligible)

# Missingness — three versions (all use `eligible` N=106 as the baseline so
# the full attrition picture is visible alongside the analytic sample):
#   missingness         : denominator = all 106 eligible (screener-qualified)
#   missingness_started : denominator = those who actually started the full survey
#   missingness_complete: denominator = 82 completers (analytic sample)
_screener_cols = {c for c in (elig_col, days_col) if c}
_started_n = detect_started_n(eligible, screener_cols=_screener_cols)
print(f"  Detected started N = {_started_n}  (eligible N = {len(eligible)},"
      f"  non-starters = {len(eligible) - _started_n})")

missingness_df           = build_missingness(eligible)
missingness_started_df   = build_missingness(eligible, started_n=_started_n)
missingness_complete_df  = build_missingness(completers)

# MCAR analysis — is missingness associated with respondent demographics?
print("\nRunning MCAR analysis ...")
mcar_df = build_mcar_analysis(completers)

# Respondent-level completeness — how many questions each person skipped
print("\nBuilding respondent-level completeness ...")
resp_complete_summary_df, resp_complete_detail_df = build_respondent_completeness(completers)
print(f"  Respondents with 0 skips: "
      f"{int((resp_complete_detail_df['N skipped'] == 0).sum())} "
      f"of {len(completers)}")

# ── Diagnose demographic question coverage ────────────────────────────────────
# Print which demo-pattern columns exist in the ineligible group and whether
# they have any answered data — helps spot form/column label mismatches.
print("\nDemographic column check (ineligible group):")
for col in ineligible.columns:
    if not any(re.search(p, col, re.IGNORECASE) for p in DEMO_PATTERNS):
        continue
    n_answered = ineligible[col].apply(norm_val).dropna().__len__()
    in_freqs   = (inelig_freqs is not None and not inelig_freqs.empty
                  and any(re.search(p, str(q), re.IGNORECASE)
                          for q in inelig_freqs["Question"].unique()
                          for p in DEMO_PATTERNS
                          if re.search(p, col, re.IGNORECASE)))
    print(f"  [{n_answered:>3} answered] {'IN freqs' if in_freqs else 'MISSING from freqs'} | {col[:90]}")

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
    """Basic standardization: strip, collapse spaces, St. → St, title case,
    and strip trailing ' County' suffix (e.g. 'Washington County' → 'Washington')."""
    val = val.strip()
    val = re.sub(r'\s+', ' ', val)
    val = re.sub(r'\bSt\.', 'St', val, flags=re.IGNORECASE)
    val = re.sub(r'\s+County\s*$', '', val, flags=re.IGNORECASE)  # strip trailing " County"
    val = val.title()
    return val

# ── County recode rules ───────────────────────────────────────────────────────
# Applied after normalize_county (which title-cases and strips trailing " County").
# Comparisons are case-insensitive (keys stored lower-case).
# "" (blank)  = invalid entry, excluded from county frequency table.
# "N/A"       = respondent does not practice in a Minnesota county
#               (out-of-state, ineligible practice type, or non-specific response).
COUNTY_RECODES = {
    # Invalid / uninterpretable entries
    "usa":                  "",      # not a county name

    # Health system name entered instead of county
    "mille lacs health":    "Mille Lacs",

    # Spelling / capitalization fixes
    "ottertail":            "Otter Tail",   # common misspelling of Otter Tail County
    "mcleod":               "McLeod",       # title-case produces Mcleod; correct to McLeod
    "lesueur":              "Le Sueur",     # title-case produces Lesueur; correct to Le Sueur

    # Out-of-state responses
    "pima":                 "N/A",   # Pima County, AZ — out of state
    "pima az":              "N/A",   # Pima County AZ — out of state (trailing "County" already stripped)
    "pima county az":       "N/A",   # Pima County Az — out of state (explicit case)
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
        METRO_COUNTIES = {
            "hennepin", "washington", "carver", "scott",
            "ramsey", "anoka", "dakota",
        }

        county_df = (
            pd.DataFrame(county_rows)
            .sort_values(["group", "county_recoded"], key=lambda s: s.str.lower())
            .reset_index(drop=True)
        )
        county_df["metro"] = county_df["county_recoded"].apply(
            lambda c: 1 if c.strip().lower() in METRO_COUNTIES else 0
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

# ── Cross-tabulations ─────────────────────────────────────────────────────────
print("\nBuilding cross-tabulations ...")

# ── Metro cross-tab ────────────────────────────────────────────────────────────
crosstab_metro_df = pd.DataFrame()
if not county_df.empty and record_id_col:
    elig_county = (
        county_df[county_df["group"] == "eligible"][["record_id", "metro"]]
        .rename(columns={"record_id": record_id_col})
    )
    eligible_geo = completers.merge(elig_county, on=record_id_col, how="left")
    eligible_geo["metro"] = pd.to_numeric(eligible_geo["metro"], errors="coerce")
    eligible_geo = eligible_geo[eligible_geo["metro"].isin([0, 1])].copy()
    eligible_geo["metro"] = eligible_geo["metro"].astype(int)
    n_metro    = (eligible_geo["metro"] == 1).sum()
    n_nonmetro = (eligible_geo["metro"] == 0).sum()
    print(f"  Metro cross-tab: metro={n_metro}, non-metro={n_nonmetro} "
          f"({len(completers) - len(eligible_geo)} excluded — no county recorded)")
    crosstab_metro_df = build_crosstab(eligible_geo, "metro", "Metro", "Non-metro")
else:
    print("  Skipping metro cross-tab — county data not available")

# ── Tenure cross-tab ───────────────────────────────────────────────────────────
crosstab_tenure_df = pd.DataFrame()
years_col = next(
    (c for c in completers.columns
     if re.search(r"how long have you been practicing", c, re.IGNORECASE)),
    None
)
if years_col:
    eligible_tenure = completers.copy()
    eligible_tenure["tenure"] = eligible_tenure[years_col].apply(
        lambda v: 1 if re.search(TENURE_HIGH_PATTERN, str(v), re.IGNORECASE)
                  else (0 if str(v).strip() not in ("", "nan") else pd.NA)
    )
    eligible_tenure = eligible_tenure[eligible_tenure["tenure"].notna()].copy()
    eligible_tenure["tenure"] = eligible_tenure["tenure"].astype(int)
    n_high = (eligible_tenure["tenure"] == 1).sum()
    n_low  = (eligible_tenure["tenure"] == 0).sum()
    print(f"  Tenure cross-tab: 20+ yrs={n_high}, <20 yrs={n_low} "
          f"({len(completers) - len(eligible_tenure)} excluded — did not answer)")
    crosstab_tenure_df = build_crosstab(eligible_tenure, "tenure", "20+ years", "<20 years")
else:
    print("  WARNING: years of practice column not found — skipping tenure cross-tab")

# ── No-screen sub-group cross-tab ─────────────────────────────────────────────
# Split eligible respondents by Q1.3.1 response, then cross-tab Q1.3.3 only.
#   Group 1 = "No, don't screen any patients"
#   Group 0 = "No, only some patients"
crosstab_no_screen_df = pd.DataFrame()
q131_col = next(
    (c for c in completers.columns
     if re.search(r"screen all patients who are pregnant or breastfeeding for cannabis",
                  c, re.IGNORECASE)
     and "(choice=" not in c),
    None
)
if q131_col:
    dont_mask  = completers[q131_col].str.strip().str.lower().str.startswith("no, don", na=False)
    some_mask  = completers[q131_col].str.strip().str.lower().str.contains("only some", na=False)
    dont_grp   = completers[dont_mask].copy()
    some_grp   = completers[some_mask].copy()
    print(f"  No-screen cross-tab: 'No, don't screen any'={len(dont_grp)}, "
          f"'No, only some patients'={len(some_grp)}")
    if len(dont_grp) > 0 and len(some_grp) > 0:
        dont_grp["_no_screen_grp"] = 1
        some_grp["_no_screen_grp"] = 0
        eligible_no_screen = pd.concat([dont_grp, some_grp], ignore_index=True)
        raw_no_screen = build_crosstab(
            eligible_no_screen, "_no_screen_grp",
            "No, don't screen any", "No, only some patients"
        )
        q133_pat = r"why don.t you or others.+screen all patients"
        if not raw_no_screen.empty:
            crosstab_no_screen_df = raw_no_screen[
                raw_no_screen["Question"].apply(
                    lambda q: bool(re.search(q133_pat, str(q), re.IGNORECASE))
                )
            ].reset_index(drop=True)
            print(f"    → {len(crosstab_no_screen_df)} rows for Q1.3.3")

        # Collect Q1.3.3 open-text responses by group (separate sheet)
        q133_q_key = next(
            (k for k in checkbox_groups
             if re.search(q133_pat, k, re.IGNORECASE)),
            None
        )
        no_screen_freetext_rows = []
        for grp_label, grp_df in [
            ("No, don't screen any patients", dont_grp),
            ("No, only some patients",        some_grp),
        ]:
            if q133_q_key:
                ft = free_text_for_question(grp_df, q133_q_key)
                if ft is not None:
                    ft.insert(1, "Group", grp_label)
                    no_screen_freetext_rows.append(ft)
        no_screen_freetext_df = (
            pd.concat(no_screen_freetext_rows, ignore_index=True)
            if no_screen_freetext_rows else pd.DataFrame()
        )
        if not no_screen_freetext_df.empty:
            print(f"    → {len(no_screen_freetext_df)} Q1.3.3 open-text responses collected")
    else:
        no_screen_freetext_df = pd.DataFrame()
        print("  WARNING: one or both sub-groups empty — skipping no-screen cross-tab")
else:
    no_screen_freetext_df = pd.DataFrame()
    print("  WARNING: Q1.3.1 column not found — skipping no-screen cross-tab")

# ── Eligibility cross-tab ─────────────────────────────────────────────────────
# Compare eligible vs ineligible on demographic questions only.
# Both groups answered these; full-survey questions are intentionally excluded.
crosstab_elig_df = pd.DataFrame()
demo_keep = [c for c in df.columns if is_demo_col(c)]
if demo_keep:
    combined_elig = pd.concat(
        [eligible.assign(eligibility=1), ineligible.assign(eligibility=0)],
        ignore_index=True,
    )
    # Retain only demographic columns (single-choice and checkbox) + group indicator
    keep_set = set(demo_keep) | {"eligibility"}
    combined_demo = combined_elig[[c for c in combined_elig.columns if c in keep_set]].copy()
    n_elig   = len(eligible)
    n_inelig = len(ineligible)
    print(f"  Eligibility cross-tab: eligible={n_elig}, ineligible={n_inelig}")
    crosstab_elig_df = build_crosstab(combined_demo, "eligibility", "Eligible", "Ineligible")
else:
    print("  WARNING: no demographic columns found — skipping eligibility cross-tab")

# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")

sheets = {
    "eligible":             elig_freqs,
    "ineligible":           inelig_freqs,
    "missingness":            missingness_df,
    "missingness_started":    missingness_started_df,
    "missingness_complete":   missingness_complete_df,
    "mcar_analysis":          mcar_df,
    "resp_complete_summary":  resp_complete_summary_df,
    "resp_complete_detail":   resp_complete_detail_df,
    "crosstab_metro":       crosstab_metro_df,
    "crosstab_tenure":      crosstab_tenure_df,
    "crosstab_no_screen":   crosstab_no_screen_df,
    "no_screen_freetext":   no_screen_freetext_df,
    "crosstab_eligibility": crosstab_elig_df,
    "county_freq":          county_freq_df,
    "county":               county_df,
    "free_text":            free_text_df,
    "completion_time":      completion_time_df,
    "completion_summary":   completion_summary_df,
}

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for sheet_name, data in sheets.items():
        if data is not None and not data.empty:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {sheet_name:<12}: {len(data)} rows")
        else:
            print(f"  {sheet_name:<12}: (empty — skipped)")

print("\nDone.")
