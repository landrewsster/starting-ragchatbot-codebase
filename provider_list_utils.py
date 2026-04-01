#!/usr/bin/env python3
"""
provider_list_utils.py — Clean, sort, and cross-reference healthcare provider lists.

Designed to work with CSV output from npi_batch_query_free.py, but can accept
any CSV with provider data. Cross-reference matching uses normalized
name + zip5 and/or name + city to find providers across lists.

Usage examples:
    # Clean and deduplicate one or more NPI CSV files
    python3 provider_list_utils.py clean npi_results_physicians.csv

    # Clean and sort by practice address (for health system annotation)
    python3 provider_list_utils.py clean npi_results_physicians.csv --sort-by-address

    # Merge all three NPI output files into one deduplicated file, sorted by address
    python3 provider_list_utils.py merge \\
        npi_results_physicians.csv \\
        npi_results_nurses.csv \\
        npi_results_physician_assistants.csv \\
        -o merged_providers.csv --sort-by-address

    # Add county name and FIPS from an HRSA ZIP-to-county crosswalk
    python3 provider_list_utils.py add-county \\
        merged_providers.csv \\
        --crosswalk hrsa_zip_county.csv

    # Cross-reference NPI CSV(s) against an external Excel roster
    python3 provider_list_utils.py crossref \\
        --npi npi_results_physicians.csv npi_results_nurses.csv \\
        --external my_roster.xlsx

    # Cross-reference with explicit column mappings for the external file
    python3 provider_list_utils.py crossref \\
        --npi npi_results_physicians.csv \\
        --external my_roster.xlsx \\
        --sheet "Sheet1" \\
        --ext-last "Last Name" \\
        --ext-first "First Name" \\
        --ext-zip  "Zip" \\
        --ext-city "City"

Requirements:
    pip3 install pandas openpyxl
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency — run:  pip3 install pandas openpyxl")

# ── Column names used by npi_batch_query_free.py ──────────────────────────────

NPI_LAST   = "last_name"
NPI_FIRST  = "first_name"
NPI_ZIP    = "zip5"
NPI_CITY   = "primary_city"
NPI_ID     = "npi"
NPI_STATUS = "status"

# Default sort order for NPI data
DEFAULT_SORT_KEYS   = [NPI_LAST, NPI_FIRST, NPI_CITY]

# Sort order for grouping by practice address (useful for health-system annotation)
ADDRESS_SORT_KEYS   = ["zip5", "primary_city", "primary_address_1", NPI_LAST, NPI_FIRST]

# ── Name / field normalisation helpers ────────────────────────────────────────

_SUFFIXES = re.compile(
    r"\b(jr\.?|sr\.?|ii|iii|iv|v|md|do|np|pa|rn|lpn|crnp|cnm|aprn)\b",
    re.IGNORECASE,
)
_WHITESPACE = re.compile(r"\s+")
_NON_ALPHA  = re.compile(r"[^a-z0-9 ]")


def _norm_str(value) -> str:
    """Lowercase, strip punctuation and suffixes, collapse whitespace."""
    if pd.isna(value) or not str(value).strip():
        return ""
    s = str(value).lower()
    s = _SUFFIXES.sub("", s)
    s = _NON_ALPHA.sub(" ", s)
    s = _WHITESPACE.sub(" ", s).strip()
    return s


def _norm_zip(value) -> str:
    """Return first 5 digits of a zip code string."""
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9]", "", str(value))[:5]


def _norm_phone(value) -> str:
    """Normalise to NNN-NNN-NNNN or return original if not 10 digits."""
    if pd.isna(value):
        return ""
    digits = re.sub(r"[^0-9]", "", str(value))
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return str(value).strip()


def _title(value) -> str:
    """Title-case a string; leave empty values alone."""
    if pd.isna(value) or not str(value).strip():
        return ""
    return str(value).strip().title()


def _parse_map(map_args: list[str]) -> dict[str, str]:
    """
    Parse --map arguments of the form 'Source Column Name=target_col'.

    Example:
        ["Provider Last Name=last_name", "Zip Code=zip5"]
        → {"Provider Last Name": "last_name", "Zip Code": "zip5"}
    """
    result = {}
    for item in (map_args or []):
        if "=" not in item:
            sys.exit(
                f"Invalid --map value: '{item}'\n"
                f"  Expected format:  'Source Column=target_column'\n"
                f"  Example:          'Provider Last Name=last_name'"
            )
        src, _, tgt = item.partition("=")
        result[src.strip()] = tgt.strip()
    return result


# ── Cleaning ───────────────────────────────────────────────────────────────────

def clean_df(
    df: pd.DataFrame,
    active_only: bool = False,
    dedup: bool = True,
    dedup_on: str = "npi",
    sort_keys: list[str] | None = None,
) -> pd.DataFrame:
    """
    Normalise and optionally deduplicate a provider DataFrame.

    Parameters
    ----------
    df          : DataFrame loaded from a provider CSV or Excel file.
    active_only : If True, keep only rows where status == 'A'.
    dedup       : If True, deduplicate using the strategy in dedup_on.
    dedup_on    : Dedup strategy — 'npi' (default), 'name', 'name+zip',
                  'name+city', or 'none'.
    sort_keys   : Column names to sort by. Defaults to last/first/city.
    """
    df = df.copy()

    # ── Normalise string fields
    name_cols = ["last_name", "first_name", "middle_name",
                 "primary_city", "mailing_city"]
    for col in name_cols:
        if col in df.columns:
            df[col] = df[col].apply(_title)

    phone_cols = [c for c in df.columns if "telephone" in c or "fax" in c]
    for col in phone_cols:
        df[col] = df[col].apply(_norm_phone)

    if "zip5" in df.columns:
        df["zip5"] = df["zip5"].apply(_norm_zip)
    if "primary_postal_code" in df.columns:
        df["primary_postal_code"] = df["primary_postal_code"].apply(
            lambda v: str(v).strip() if not pd.isna(v) else ""
        )

    # Strip leading/trailing whitespace from all remaining string columns
    str_cols = df.select_dtypes(include=["object", "str"]).columns
    for col in str_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # ── Optional filter: active providers only
    if active_only and NPI_STATUS in df.columns:
        before = len(df)
        df = df[df[NPI_STATUS].str.upper() == "A"].copy()
        print(f"    active_only: kept {len(df)} of {before} rows")

    # ── Deduplicate
    if dedup and dedup_on != "none":
        if dedup_on == "npi":
            if NPI_ID in df.columns:
                before = len(df)
                df = df.drop_duplicates(subset=[NPI_ID], keep="first").copy()
                removed = before - len(df)
                if removed:
                    print(f"    dedup (npi): removed {removed} duplicate NPI row(s)")
        elif dedup_on in ("name", "name+zip", "name+city"):
            key_parts: dict[str, pd.Series] = {}
            if NPI_LAST in df.columns and NPI_FIRST in df.columns:
                key_parts["_dk_name"] = (
                    df[NPI_LAST].apply(_norm_str) + "|" + df[NPI_FIRST].apply(_norm_str)
                )
            if dedup_on == "name+zip" and NPI_ZIP in df.columns:
                key_parts["_dk_zip"] = df[NPI_ZIP].apply(_norm_zip)
            if dedup_on == "name+city" and NPI_CITY in df.columns:
                key_parts["_dk_city"] = df[NPI_CITY].apply(_norm_str)
            if key_parts:
                for col, series in key_parts.items():
                    df[col] = series
                before = len(df)
                df = df.drop_duplicates(subset=list(key_parts.keys()), keep="first").copy()
                removed = before - len(df)
                if removed:
                    print(f"    dedup ({dedup_on}): removed {removed} duplicate row(s)")
                df = df.drop(columns=list(key_parts.keys()))

    # ── Sort
    keys = sort_keys or DEFAULT_SORT_KEYS
    valid_keys = [k for k in keys if k in df.columns]
    if valid_keys:
        df = df.sort_values(valid_keys, ignore_index=True)

    return df


# ── Loading helpers ────────────────────────────────────────────────────────────

def _load_one(path: Path, sheet: str | None = None) -> pd.DataFrame:
    """Load a single CSV or Excel file into a DataFrame."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    elif suffix == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        sys.exit(f"Unsupported file type '{suffix}' for {path}. Use .csv, .xlsx, or .xls.")
    df.columns = df.columns.astype(str).str.strip()
    # Ensure zip5 is always present
    if "zip5" not in df.columns and "primary_postal_code" in df.columns:
        df["zip5"] = df["primary_postal_code"].apply(_norm_zip)
    return df


def load_provider_files(
    paths: list[Path],
    sheet: str | None = None,
    col_map: dict[str, str] | None = None,
) -> pd.DataFrame:
    """
    Load one or more CSV or Excel provider files and concatenate them.

    A 'source_file' column is always added so every row can be traced back to
    its origin file.  Files with different column sets are stacked; missing
    columns are filled with empty strings rather than raising an error.

    col_map : optional dict of {source_column: target_column} renames applied
              to any file that contains the source column.  Columns not present
              in a given file are silently skipped, so the same map can be
              passed for all files without affecting NPI files that don't have
              those source column names.
    """
    frames = []
    for p in paths:
        print(f"  Loading {p} …")
        df = _load_one(p, sheet=sheet)
        if col_map:
            rename = {src: tgt for src, tgt in col_map.items() if src in df.columns}
            if rename:
                print(f"    Column remap: {rename}")
                df = df.rename(columns=rename)
        df["source_file"] = p.name
        print(f"    {len(df)} rows, {len(df.columns)} columns")
        frames.append(df)
    if not frames:
        sys.exit("No provider files loaded.")
    combined = pd.concat(frames, ignore_index=True).fillna("")
    print(f"  Total: {len(combined)} rows from {len(paths)} file(s).")
    return combined


def flag_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add duplicate-inspection columns without removing any rows.

    Adds:
        _name_key    — normalised 'last|first' used for matching
        _dupe_count  — how many rows share this name key (1 = unique)
        _dupe_group  — integer group ID shared by all rows with the same name
                       (0 for rows that appear only once)
        _zip_match   — 'Y' if another row in the same dupe group shares the
                       same zip5 but comes from a different source file,
                       indicating a strong cross-file match; 'N' otherwise
        _files_in_group — comma-separated list of source files that contain
                          this name, so you can see at a glance whether a
                          name appears in just one file or across multiple

    Rows with _dupe_count > 1 are potential duplicates for review.
    In the dupes review file they are sorted by _dupe_group then source_file
    so every group of matching rows appears together.
    """
    df = df.copy()
    if NPI_LAST not in df.columns or NPI_FIRST not in df.columns:
        print("  WARNING: Cannot flag duplicates — last_name or first_name column missing.")
        return df

    df["_name_key"] = (
        df[NPI_LAST].apply(_norm_str) + "|" + df[NPI_FIRST].apply(_norm_str)
    )
    counts = df["_name_key"].map(df["_name_key"].value_counts())
    df["_dupe_count"] = counts.astype(int)

    dupe_keys = df.loc[df["_dupe_count"] > 1, "_name_key"].unique()
    group_map = {k: i + 1 for i, k in enumerate(sorted(dupe_keys))}
    df["_dupe_group"] = df["_name_key"].map(group_map).fillna(0).astype(int)

    # ── _files_in_group: which source files contain this name
    if "source_file" in df.columns:
        files_per_key = (
            df[df["_dupe_count"] > 1]
            .groupby("_name_key")["source_file"]
            .apply(lambda s: ", ".join(sorted(s.unique())))
        )
        df["_files_in_group"] = df["_name_key"].map(files_per_key).fillna("")
    else:
        df["_files_in_group"] = ""

    # ── _zip_match: Y if another row in the same group from a different file
    #               shares this row's zip5
    df["_zip_match"] = "N"
    if NPI_ZIP in df.columns and "source_file" in df.columns:
        dupes_df = df[df["_dupe_count"] > 1].copy()
        dupes_df["_zip_norm"] = dupes_df[NPI_ZIP].apply(_norm_zip)
        for idx, row in dupes_df.iterrows():
            group_rows = dupes_df[
                (dupes_df["_dupe_group"] == row["_dupe_group"]) &
                (dupes_df["source_file"] != row["source_file"]) &
                (dupes_df["_zip_norm"] == row["_zip_norm"]) &
                (dupes_df["_zip_norm"] != "")
            ]
            if not group_rows.empty:
                df.at[idx, "_zip_match"] = "Y"

    # ── Cross-file only: flag rows whose name appears in >1 distinct source file
    if "source_file" in df.columns:
        cross_keys = set(
            df[df["_dupe_count"] > 1]
            .groupby("_name_key")["source_file"]
            .filter(lambda s: s.nunique() > 1)
            .index
        )
        # Recompute: only count as dupe if name spans multiple files
        cross_file_mask = df["_name_key"].isin(
            df.groupby("_name_key")["source_file"]
            .transform("nunique") > 1
            if "source_file" in df.columns else pd.Series(False, index=df.index)
        )

    n_groups     = len(dupe_keys)
    n_rows       = (df["_dupe_count"] > 1).sum()
    n_zip_match  = (df["_zip_match"] == "Y").sum()
    print(f"  Duplicate check: {n_groups} name(s) appear more than once "
          f"({n_rows} total rows); {n_zip_match} also share a zip code.")
    return df


# Keep old name as an alias so any existing callers still work.
load_npi_csvs = load_provider_files


def load_external_excel(
    path: Path,
    sheet: str | None = None,
    col_last:  str | None = None,
    col_first: str | None = None,
    col_zip:   str | None = None,
    col_city:  str | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Load an external provider list from an Excel file.

    Returns (df, col_map) where col_map maps role→actual_column_name for the
    four match fields (last, first, zip, city).  Auto-detects column names
    if explicit names are not supplied.
    """
    print(f"  Loading external file {path} …")
    df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    df.columns = df.columns.astype(str).str.strip()
    print(f"  Found {len(df)} rows, columns: {list(df.columns)}")

    def _detect(explicit: str | None, patterns: list[str]) -> str | None:
        if explicit:
            if explicit in df.columns:
                return explicit
            sys.exit(f"Column '{explicit}' not found in {path}. "
                     f"Available: {list(df.columns)}")
        lower_cols = {c.lower(): c for c in df.columns}
        for pat in patterns:
            if pat in lower_cols:
                return lower_cols[pat]
        return None

    col_map = {
        "last":  _detect(col_last,  ["last name", "lastname", "last_name", "lname", "surname"]),
        "first": _detect(col_first, ["first name", "firstname", "first_name", "fname", "given name"]),
        "zip":   _detect(col_zip,   ["zip", "zip5", "zip code", "postal code", "postalcode", "postal_code"]),
        "city":  _detect(col_city,  ["city", "primary_city", "location city"]),
    }

    missing = [role for role, col in col_map.items() if col is None]
    if missing:
        print(
            f"\n  WARNING: Could not auto-detect external columns for: {missing}\n"
            f"  Use --ext-last, --ext-first, --ext-zip, --ext-city to specify them.\n"
            f"  Available columns: {list(df.columns)}\n"
        )

    return df, col_map


# ── Cross-reference ────────────────────────────────────────────────────────────

def _make_match_keys(
    df: pd.DataFrame,
    col_last: str | None,
    col_first: str | None,
    col_zip: str | None,
    col_city: str | None,
) -> pd.DataFrame:
    """
    Add helper columns _key_name_zip and _key_name_city to df for matching.
    """
    df = df.copy()

    last  = df[col_last].apply(_norm_str)  if col_last  and col_last  in df.columns else pd.Series("", index=df.index)
    first = df[col_first].apply(_norm_str) if col_first and col_first in df.columns else pd.Series("", index=df.index)
    zip_  = df[col_zip].apply(_norm_zip)   if col_zip   and col_zip   in df.columns else pd.Series("", index=df.index)
    city  = df[col_city].apply(_norm_str)  if col_city  and col_city  in df.columns else pd.Series("", index=df.index)

    df["_key_name_zip"]  = last + "|" + first + "|" + zip_
    df["_key_name_city"] = last + "|" + first + "|" + city
    return df


def crossref(
    npi_df: pd.DataFrame,
    ext_df: pd.DataFrame,
    ext_col_map: dict,
    match_zip: bool = True,
    match_city: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Cross-reference NPI providers against an external list.

    Matching strategy (applied in order, a provider matched by zip is not
    re-matched by city):
        1. Name + zip5  (if match_zip is True)
        2. Name + city  (if match_city is True, for rows not yet matched)

    Returns
    -------
    matched        : Rows present in both lists (NPI columns + external columns).
    npi_only       : NPI rows with no match in external list.
    external_only  : External rows with no match in NPI list.
    """
    # Add match keys to NPI data
    npi = _make_match_keys(
        npi_df,
        col_last=NPI_LAST, col_first=NPI_FIRST,
        col_zip=NPI_ZIP, col_city=NPI_CITY,
    )

    # Add match keys to external data
    ext = _make_match_keys(
        ext_df,
        col_last=ext_col_map.get("last"),
        col_first=ext_col_map.get("first"),
        col_zip=ext_col_map.get("zip"),
        col_city=ext_col_map.get("city"),
    )

    # Prefix external columns to avoid collisions (except key cols)
    ext_renamed = ext.rename(
        columns={c: f"ext_{c}" for c in ext.columns if not c.startswith("_key_")}
    )

    matched_parts   = []
    npi_matched_idx = set()
    ext_matched_idx = set()

    def _merge_on_key(key_col):
        nonlocal npi_matched_idx, ext_matched_idx
        npi_unmatched = npi[~npi.index.isin(npi_matched_idx)]
        ext_unmatched = ext_renamed[~ext_renamed.index.isin(ext_matched_idx)]

        # Remove empty keys (can't match on blank)
        npi_keyed = npi_unmatched[npi_unmatched[key_col].str.len() > 2]
        ext_keyed = ext_unmatched[ext_unmatched[key_col].str.len() > 2]

        merged = pd.merge(
            npi_keyed, ext_keyed,
            left_on=key_col, right_on=key_col,
            how="inner",
            suffixes=("", "_ext_dup"),
        )
        if len(merged):
            matched_parts.append(merged)
            npi_matched_idx.update(merged.index.tolist())
            # Recover original ext indices via the key
            matched_ext_keys = set(merged[key_col])
            ext_matched_idx.update(
                ext_renamed[ext_renamed[key_col].isin(matched_ext_keys)].index.tolist()
            )
            print(f"    Matched {len(merged)} rows on {key_col}")

    if match_zip:
        _merge_on_key("_key_name_zip")
    if match_city:
        _merge_on_key("_key_name_city")

    if matched_parts:
        # Drop internal key columns from output
        matched = pd.concat(matched_parts, ignore_index=True)
        drop_cols = [c for c in matched.columns if c.startswith("_key_")]
        matched = matched.drop(columns=drop_cols)
    else:
        matched = pd.DataFrame()

    npi_only = npi_df[~npi_df.index.isin(npi_matched_idx)].copy()
    external_only = ext_df[~ext_df.index.isin(ext_matched_idx)].copy()

    return matched, npi_only, external_only


# ── HRSA ZIP-to-county crosswalk ──────────────────────────────────────────────

def load_hrsa_crosswalk(path: Path) -> pd.DataFrame:
    """
    Load an HRSA ZIP-to-county crosswalk file (CSV or Excel).

    Accepts the standard HRSA formats as well as common variants.  The
    returned DataFrame always has three normalised columns:
        _xw_zip        – 5-digit ZIP string
        county_name    – county name (title-cased)
        county_fips    – 5-digit FIPS code string (state 2 + county 3)

    Typical HRSA crosswalk column names handled automatically:
        ZIP / ZIP_CODE / ZIPCODE / ZIP5
        COUNTY / COUNTY_NAME / CNTY_NAME / CNTY / CO_NAME
        STCOU / COUNTY_FIPS / FIPS / CO_FIPS / FIPS_CODE
    """
    suffix = path.suffix.lower()
    print(f"  Loading HRSA crosswalk {path} …")
    if suffix in (".xlsx", ".xls"):
        xw = pd.read_excel(path, dtype=str)
    else:
        xw = pd.read_csv(path, dtype=str)
    xw.columns = xw.columns.astype(str).str.strip()
    print(f"  Crosswalk columns: {list(xw.columns)}")

    lower = {c.lower(): c for c in xw.columns}

    def _find(candidates: list[str]) -> str | None:
        for name in candidates:
            if name in lower:
                return lower[name]
        return None

    col_zip   = _find(["zip_code", "zip", "zipcode", "zip5", "zcta", "zcta5"])
    col_county = _find(["county_name", "county", "cnty_name", "cnty", "co_name", "countyname"])
    col_fips   = _find(["stcou", "county_fips", "fips", "co_fips", "fips_code", "countyfips", "geoid"])

    missing = [role for role, col in [("zip", col_zip), ("county", col_county), ("fips", col_fips)] if col is None]
    if missing:
        print(
            f"\n  WARNING: Could not auto-detect crosswalk columns for: {missing}\n"
            f"  Available: {list(xw.columns)}\n"
            f"  county_name and county_fips will be blank for undetected fields.\n"
        )

    result = pd.DataFrame()
    result["_xw_zip"]     = xw[col_zip].apply(_norm_zip) if col_zip else ""
    result["county_name"] = xw[col_county].apply(_title) if col_county else ""
    result["county_fips"] = (
        xw[col_fips].apply(lambda v: re.sub(r"[^0-9]", "", str(v)).zfill(5) if not pd.isna(v) else "")
        if col_fips else ""
    )

    # A single ZIP may appear in multiple counties (split ZIPs); keep the
    # row with the longest county name as a simple tiebreaker, then dedup.
    result["_name_len"] = result["county_name"].str.len()
    result = (
        result.sort_values("_name_len", ascending=False)
              .drop_duplicates(subset=["_xw_zip"], keep="first")
              .drop(columns=["_name_len"])
              .reset_index(drop=True)
    )
    print(f"  Crosswalk: {len(result)} unique ZIP entries loaded.")
    return result


def add_county(df: pd.DataFrame, crosswalk: pd.DataFrame) -> pd.DataFrame:
    """
    Join county_name and county_fips onto a provider DataFrame using zip5.

    Rows whose zip5 is not found in the crosswalk receive empty strings.
    If county columns already exist they are overwritten.
    """
    for col in ("county_name", "county_fips"):
        if col in df.columns:
            df = df.drop(columns=[col])

    merged = df.merge(
        crosswalk.rename(columns={"_xw_zip": "zip5"}),
        on="zip5",
        how="left",
    )
    merged["county_name"] = merged["county_name"].fillna("")
    merged["county_fips"] = merged["county_fips"].fillna("")

    matched = (merged["county_name"] != "").sum()
    print(f"  County lookup: {matched} of {len(merged)} rows matched.")
    return merged


# ── CLI subcommands ────────────────────────────────────────────────────────────

def cmd_inspect(args):
    """Print column names and one example value per column for each file."""
    for fname in args.files:
        path = Path(fname)
        print(f"\n{'═'*60}")
        print(f"  File   : {path}")
        df = _load_one(path, sheet=getattr(args, "sheet", None))
        print(f"  Rows   : {len(df)}")
        print(f"  Columns: {len(df.columns)}")
        print()
        for col in df.columns:
            sample = df[col].dropna().astype(str)
            sample = sample[sample.str.strip() != ""]
            example = sample.iloc[0] if len(sample) else "(empty)"
            # Truncate long examples for readability
            if len(example) > 60:
                example = example[:57] + "..."
            print(f"    {col!r:<45}  e.g. {example!r}")
    print(f"\n{'═'*60}")


def cmd_clean(args):
    paths = [Path(f) for f in args.files]
    if args.sort_by_address:
        sort_keys = ADDRESS_SORT_KEYS
    else:
        sort_keys = args.sort.split(",") if args.sort else None
    for path in paths:
        print(f"\n{'─'*60}")
        print(f"  Cleaning: {path}")
        df = _load_one(path, sheet=getattr(args, "sheet", None))
        df = clean_df(df, active_only=args.active_only, dedup=not args.no_dedup, sort_keys=sort_keys)
        out = args.output or path.with_stem(f"cleaned_{path.stem}").with_suffix(".csv")
        df.to_csv(out, index=False)
        print(f"  Saved {len(df)} rows → {Path(out).resolve()}")


def cmd_merge(args):
    print(f"\n{'─'*60}")
    paths = [Path(f) for f in args.files]
    if args.sort_by_address:
        sort_keys = ADDRESS_SORT_KEYS
    else:
        sort_keys = args.sort.split(",") if args.sort else None
    col_map  = _parse_map(getattr(args, "map", None) or [])
    dedup_on = getattr(args, "dedup_on", "none") or "none"

    df = load_provider_files(paths, sheet=getattr(args, "sheet", None), col_map=col_map or None)

    # ── Flag duplicates BEFORE any dedup so the review file captures all matches
    df = flag_duplicates(df)
    out = args.output or "merged_providers.csv"
    dupes = df[df["_dupe_count"] > 1].sort_values(["_dupe_group", "source_file"])
    if not dupes.empty:
        dupe_out = Path(out).with_name(f"dupes_{Path(out).name}")
        dupes.to_csv(dupe_out, index=False)
        print(f"  Review file    : {len(dupes)} rows → {dupe_out.resolve()}")
        print(f"  Columns to review: source_file | _dupe_group | _zip_match | _files_in_group")
        print(f"  (All rows kept in main file — delete unwanted rows after review.)")

    # ── Then clean/normalise/dedup/sort
    df = clean_df(df, active_only=args.active_only, dedup=not args.no_dedup,
                  dedup_on=dedup_on, sort_keys=sort_keys)
    df.to_csv(out, index=False)
    print(f"\n  Merged: {len(df)} rows → {Path(out).resolve()}")


def compare_datasets(
    npi_df: pd.DataFrame,
    mn_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Split providers into three mutually exclusive output sets based on how
    their name appears across the NPI dataset and the MN licensing list.

    Categories (from the MN list's perspective):
        matched_1to1    mn_count == 1  AND  npi_count == 1
                        Clean 1:1 match. Both the NPI row and MN row are
                        included, sorted adjacent for easy comparison.

        multi_specialty mn_count > 1
                        Provider appears more than once in the MN list
                        (multiple specialties/licenses). All their MN rows
                        are included, plus their NPI row(s) if any exist.

        mn_only         mn_count == 1  AND  npi_count == 0
                        In the MN list but no matching name in NPI data.

        npi_only        npi_count > 0  AND  mn_count == 0
                        In the NPI data but no matching name in MN list.
                        Returned for completeness; not written by default.

    All DataFrames include a '_name_key' column (normalised last|first) and
    a 'source_file' column so every row can be traced to its origin file.
    """
    npi = npi_df.copy()
    mn  = mn_df.copy()

    npi["_name_key"] = npi[NPI_LAST].apply(_norm_str) + "|" + npi[NPI_FIRST].apply(_norm_str)
    mn["_name_key"]  = mn[NPI_LAST].apply(_norm_str)  + "|" + mn[NPI_FIRST].apply(_norm_str)

    npi_counts = npi["_name_key"].value_counts().to_dict()
    mn_counts  = mn["_name_key"].value_counts().to_dict()

    # Tag each MN row with counts from both datasets
    mn["_npi_count"] = mn["_name_key"].map(npi_counts).fillna(0).astype(int)
    mn["_mn_count"]  = mn["_name_key"].map(mn_counts).fillna(0).astype(int)
    npi["_npi_count"] = npi["_name_key"].map(npi_counts).fillna(0).astype(int)
    npi["_mn_count"]  = npi["_name_key"].map(mn_counts).fillna(0).astype(int)

    # ── Category keys
    matched_keys      = {k for k in mn_counts if mn_counts[k] == 1 and npi_counts.get(k, 0) == 1}
    multi_spec_keys   = {k for k in mn_counts if mn_counts[k] > 1 and npi_counts.get(k, 0) >= 1}
    mn_only_keys      = {k for k in mn_counts if npi_counts.get(k, 0) == 0}
    npi_only_keys     = {k for k in npi_counts if mn_counts.get(k, 0) == 0}

    # ── File 1: 1:1 matched — show NPI row + MN row adjacent, sorted by name
    npi_matched = npi[npi["_name_key"].isin(matched_keys)].copy()
    mn_matched  = mn[mn["_name_key"].isin(matched_keys)].copy()
    matched_1to1 = (
        pd.concat([npi_matched, mn_matched], ignore_index=True)
        .sort_values(["_name_key", "source_file"])
        .reset_index(drop=True)
    )

    # ── File 2: multi-specialty — all MN rows + NPI rows if name exists in NPI
    npi_multi = npi[npi["_name_key"].isin(multi_spec_keys)].copy()
    mn_multi  = mn[mn["_name_key"].isin(multi_spec_keys)].copy()
    multi_specialty = (
        pd.concat([npi_multi, mn_multi], ignore_index=True)
        .sort_values(["_name_key", "source_file"])
        .reset_index(drop=True)
    )

    # ── File 3: MN list only, no NPI match
    mn_only = mn[mn["_name_key"].isin(mn_only_keys)].copy().reset_index(drop=True)

    # ── File 4 (informational): NPI only, no MN match
    npi_only = npi[npi["_name_key"].isin(npi_only_keys)].copy().reset_index(drop=True)

    return matched_1to1, multi_specialty, mn_only, npi_only


def cmd_compare(args):
    print(f"\n{'─'*60}")
    col_map = _parse_map(getattr(args, "map", None) or [])

    # Load and clean NPI files
    npi_paths = [Path(f) for f in args.npi]
    npi_df = load_provider_files(npi_paths, col_map=col_map or None)
    npi_df = clean_df(npi_df, active_only=False, dedup=True, dedup_on="npi")

    # Load and clean MN list
    mn_path = Path(args.mn_list)
    mn_df = load_provider_files([mn_path], sheet=getattr(args, "sheet", None), col_map=col_map or None)
    mn_df = clean_df(mn_df, active_only=False, dedup=False)

    if NPI_LAST not in npi_df.columns or NPI_LAST not in mn_df.columns:
        sys.exit("Both datasets must have 'last_name' and 'first_name' columns. "
                 "Use --map to rename columns if needed.")

    print(f"\n  Comparing {len(npi_df)} NPI rows vs {len(mn_df)} MN list rows …")
    matched, multi, mn_only, npi_only = compare_datasets(npi_df, mn_df)

    stem = args.output_prefix or "compare"
    out_matched = f"{stem}_matched_1to1.csv"
    out_multi   = f"{stem}_multi_specialty.csv"
    out_mn_only = f"{stem}_mn_only.csv"
    out_npi_only = f"{stem}_npi_only.csv"

    matched.to_csv(out_matched,  index=False)
    multi.to_csv(out_multi,      index=False)
    mn_only.to_csv(out_mn_only,  index=False)
    npi_only.to_csv(out_npi_only, index=False)

    n_matched_providers = matched["_name_key"].nunique()
    n_multi_providers   = multi["_name_key"].nunique()

    print(f"\n  {'─'*56}")
    print(f"  1:1 matched providers : {n_matched_providers:>6}  ({len(matched)} rows)  → {out_matched}")
    print(f"  Multi-specialty (MN)  : {n_multi_providers:>6}  ({len(multi)} rows)  → {out_multi}")
    print(f"  MN list only          : {len(mn_only):>6}  (no NPI match)     → {out_mn_only}")
    print(f"  NPI only              : {len(npi_only):>6}  (no MN match)      → {out_npi_only}")
    print(f"  {'─'*56}")
    print(f"  In each file, rows with the same provider are adjacent.")
    print(f"  Use 'source_file' to see which row came from which dataset.")


def cmd_add_county(args):
    print(f"\n{'─'*60}")
    path = Path(args.file)
    print(f"  Loading provider file {path} …")
    df = pd.read_csv(path, dtype=str)
    if "zip5" not in df.columns and "primary_postal_code" in df.columns:
        df["zip5"] = df["primary_postal_code"].apply(_norm_zip)

    xw = load_hrsa_crosswalk(Path(args.crosswalk))
    df = add_county(df, xw)

    out = args.output or path.with_name(f"county_{path.name}")
    df.to_csv(out, index=False)
    print(f"  Saved {len(df)} rows → {Path(out).resolve()}")


def cmd_crossref(args):
    print(f"\n{'─'*60}")
    npi_paths = [Path(f) for f in args.npi]
    npi_df = load_npi_csvs(npi_paths)
    npi_df = clean_df(npi_df, active_only=args.active_only, dedup=True)

    ext_path = Path(args.external)
    ext_df, col_map = load_external_excel(
        ext_path,
        sheet=args.sheet,
        col_last=args.ext_last,
        col_first=args.ext_first,
        col_zip=args.ext_zip,
        col_city=args.ext_city,
    )

    print(f"\n  Matching {len(npi_df)} NPI providers against "
          f"{len(ext_df)} external records …")

    matched, npi_only, ext_only = crossref(
        npi_df, ext_df, col_map,
        match_zip=True,
        match_city=True,
    )

    stem = ext_path.stem
    out_matched  = f"crossref_{stem}_matched.csv"
    out_npi_only = f"crossref_{stem}_npi_only.csv"
    out_ext_only = f"crossref_{stem}_external_only.csv"

    if not matched.empty:
        matched.to_csv(out_matched, index=False)
        print(f"\n  Matched        : {len(matched):>5} rows → {out_matched}")
    else:
        print("\n  No matches found.")

    npi_only.to_csv(out_npi_only, index=False)
    ext_only.to_csv(out_ext_only, index=False)
    print(f"  NPI only       : {len(npi_only):>5} rows → {out_npi_only}")
    print(f"  External only  : {len(ext_only):>5} rows → {out_ext_only}")


# ── Argument parser ────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Clean, sort, and cross-reference healthcare provider lists.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # ── inspect ────────────────────────────────────────────────────────────────
    p_inspect = sub.add_parser("inspect", help="Show column names and sample values for one or more files.")
    p_inspect.add_argument("files", nargs="+", help="File(s) to inspect (.csv or .xlsx).")
    p_inspect.add_argument("--sheet", default=None, help="Excel sheet name or index (default: first sheet).")

    # ── clean ──────────────────────────────────────────────────────────────────
    p_clean = sub.add_parser("clean", help="Clean and deduplicate one or more provider files (CSV or Excel).")
    p_clean.add_argument("files", nargs="+", help="Provider file(s) to clean (.csv or .xlsx).")
    p_clean.add_argument("-o", "--output", help="Output CSV path (only used when a single file is given).")
    p_clean.add_argument("--sheet", default=None, help="Excel sheet name or index (default: first sheet).")
    p_clean.add_argument("--active-only", action="store_true", help="Keep only active providers (status=A).")
    p_clean.add_argument("--no-dedup", action="store_true", help="Skip NPI deduplication.")
    p_clean.add_argument("--sort", help="Comma-separated sort keys (default: last_name,first_name,primary_city).")
    p_clean.add_argument("--sort-by-address", action="store_true",
                         help="Sort by zip5, city, address_1 to group co-located providers (useful for health-system annotation).")

    # ── merge ──────────────────────────────────────────────────────────────────
    p_merge = sub.add_parser("merge", help="Merge multiple provider files (CSV or Excel) into one deduplicated CSV.")
    p_merge.add_argument("files", nargs="+", help="Provider file(s) to merge (.csv or .xlsx).")
    p_merge.add_argument("-o", "--output", default="merged_providers.csv", help="Output CSV path.")
    p_merge.add_argument("--sheet", default=None, help="Excel sheet name or index applied to all Excel inputs (default: first sheet).")
    p_merge.add_argument("--active-only", action="store_true", help="Keep only active providers (status=A).")
    p_merge.add_argument("--no-dedup", action="store_true", help="Skip deduplication (default behaviour — all rows kept).")
    p_merge.add_argument("--dedup-on", default="none", dest="dedup_on",
                         choices=["npi", "name", "name+zip", "name+city", "none"],
                         help="Dedup strategy (default: none — keep all rows and flag dupes for review). "
                              "Set to 'name' to auto-remove duplicates by last+first name after you have reviewed them.")
    p_merge.add_argument("--map", nargs="+", metavar="SRC=TGT", dest="map",
                         help="Rename columns before merging, e.g. --map 'Provider Last Name=last_name' 'Zip=zip5'. "
                              "Applied only to files that contain the source column name.")
    p_merge.add_argument("--sort", help="Comma-separated sort keys.")
    p_merge.add_argument("--sort-by-address", action="store_true",
                         help="Sort by zip5, city, address_1 to group co-located providers.")

    # ── compare ────────────────────────────────────────────────────────────────
    p_cmp = sub.add_parser(
        "compare",
        help="Split providers into 3 files: 1:1 NPI+MN matches, MN multi-specialty, MN-only."
    )
    p_cmp.add_argument("--npi",     nargs="+", required=True,
                       help="Cleaned NPI CSV file(s).")
    p_cmp.add_argument("--mn-list", required=True, dest="mn_list",
                       help="MN Physician and PA list (.xlsx or .csv).")
    p_cmp.add_argument("--sheet",   default=None,
                       help="Excel sheet name for the MN list (default: first sheet).")
    p_cmp.add_argument("--map",     nargs="+", metavar="SRC=TGT", dest="map",
                       help="Column renames applied to MN list before comparing, "
                            "e.g. --map 'address_line1=primary_address_1' 'zip=zip5'.")
    p_cmp.add_argument("--output-prefix", default="compare", dest="output_prefix",
                       help="Prefix for output filenames (default: 'compare').")

    # ── add-county ─────────────────────────────────────────────────────────────
    p_ac = sub.add_parser("add-county",
                          help="Join county name and FIPS from an HRSA ZIP-to-county crosswalk.")
    p_ac.add_argument("file", help="Provider CSV file to enrich.")
    p_ac.add_argument("--crosswalk", required=True,
                      help="HRSA ZIP-to-county crosswalk file (.csv or .xlsx). "
                           "Download from: https://data.hrsa.gov/  or the HUD USPS crosswalk.")
    p_ac.add_argument("-o", "--output", help="Output CSV path (default: county_<input>.csv).")

    # ── crossref ───────────────────────────────────────────────────────────────
    p_cr = sub.add_parser("crossref", help="Cross-reference NPI data against an external Excel roster.")
    p_cr.add_argument("--npi",      nargs="+", required=True, help="NPI CSV file(s).")
    p_cr.add_argument("--external", required=True,             help="External Excel (.xlsx) file.")
    p_cr.add_argument("--sheet",    default=None,              help="Sheet name (default: first sheet).")
    p_cr.add_argument("--active-only", action="store_true",    help="Keep only active NPI providers (status=A).")
    p_cr.add_argument("--ext-last",  default=None, help="Column name for last name in external file.")
    p_cr.add_argument("--ext-first", default=None, help="Column name for first name in external file.")
    p_cr.add_argument("--ext-zip",   default=None, help="Column name for zip code in external file.")
    p_cr.add_argument("--ext-city",  default=None, help="Column name for city in external file.")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    if args.command == "inspect":
        cmd_inspect(args)
    elif args.command == "clean":
        cmd_clean(args)
    elif args.command == "merge":
        cmd_merge(args)
    elif args.command == "compare":
        cmd_compare(args)
    elif args.command == "add-county":
        cmd_add_county(args)
    elif args.command == "crossref":
        cmd_crossref(args)
    print(f"\n{'─'*60}\n  Done.\n")


if __name__ == "__main__":
    main()
