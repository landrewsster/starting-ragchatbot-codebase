#!/usr/bin/env python3
"""
provider_list_utils.py — Clean, sort, and cross-reference healthcare provider lists.

Designed to work with CSV output from npi_batch_query_free.py, but can accept
any CSV with provider data. Cross-reference matching uses normalized
name + zip5 and/or name + city to find providers across lists.

Usage examples:
    # Clean and deduplicate one or more NPI CSV files
    python3 provider_list_utils.py clean npi_results_physicians.csv

    # Merge all three NPI output files into one deduplicated CSV
    python3 provider_list_utils.py merge \\
        npi_results_physicians.csv \\
        npi_results_nurses.csv \\
        npi_results_physician_assistants.csv \\
        -o merged_providers.csv

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
DEFAULT_SORT_KEYS = [NPI_LAST, NPI_FIRST, NPI_CITY]

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


# ── Cleaning ───────────────────────────────────────────────────────────────────

def clean_df(
    df: pd.DataFrame,
    active_only: bool = False,
    dedup: bool = True,
    sort_keys: list[str] | None = None,
) -> pd.DataFrame:
    """
    Normalise and optionally deduplicate a provider DataFrame.

    Parameters
    ----------
    df          : DataFrame loaded from an NPI CSV.
    active_only : If True, keep only rows where status == 'A'.
    dedup       : If True, drop rows with duplicate NPI numbers (keep first).
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
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # ── Optional filter: active providers only
    if active_only and NPI_STATUS in df.columns:
        before = len(df)
        df = df[df[NPI_STATUS].str.upper() == "A"].copy()
        print(f"    active_only: kept {len(df)} of {before} rows")

    # ── Deduplicate by NPI
    if dedup and NPI_ID in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset=[NPI_ID], keep="first").copy()
        removed = before - len(df)
        if removed:
            print(f"    dedup: removed {removed} duplicate NPI row(s)")

    # ── Sort
    keys = sort_keys or DEFAULT_SORT_KEYS
    valid_keys = [k for k in keys if k in df.columns]
    if valid_keys:
        df = df.sort_values(valid_keys, ignore_index=True)

    return df


# ── Loading helpers ────────────────────────────────────────────────────────────

def load_npi_csvs(paths: list[Path]) -> pd.DataFrame:
    """Load one or more NPI CSV files and concatenate them."""
    frames = []
    for p in paths:
        print(f"  Loading {p} …")
        df = pd.read_csv(p, dtype=str)
        # Ensure zip5 is always present (derive from postal code if missing)
        if "zip5" not in df.columns and "primary_postal_code" in df.columns:
            df["zip5"] = df["primary_postal_code"].apply(_norm_zip)
        frames.append(df)
    if not frames:
        sys.exit("No NPI files loaded.")
    combined = pd.concat(frames, ignore_index=True)
    print(f"  Loaded {len(combined)} total rows from {len(paths)} file(s).")
    return combined


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


# ── CLI subcommands ────────────────────────────────────────────────────────────

def cmd_clean(args):
    paths = [Path(f) for f in args.files]
    sort_keys = args.sort.split(",") if args.sort else None
    for path in paths:
        print(f"\n{'─'*60}")
        print(f"  Cleaning: {path}")
        df = pd.read_csv(path, dtype=str)
        if "zip5" not in df.columns and "primary_postal_code" in df.columns:
            df["zip5"] = df["primary_postal_code"].apply(_norm_zip)
        df = clean_df(df, active_only=args.active_only, dedup=not args.no_dedup, sort_keys=sort_keys)
        out = args.output or path.with_name(f"cleaned_{path.name}")
        df.to_csv(out, index=False)
        print(f"  Saved {len(df)} rows → {Path(out).resolve()}")


def cmd_merge(args):
    print(f"\n{'─'*60}")
    paths = [Path(f) for f in args.files]
    sort_keys = args.sort.split(",") if args.sort else None
    df = load_npi_csvs(paths)
    df = clean_df(df, active_only=args.active_only, dedup=True, sort_keys=sort_keys)
    out = args.output or "merged_providers.csv"
    df.to_csv(out, index=False)
    print(f"\n  Merged: {len(df)} unique rows → {Path(out).resolve()}")


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

    # ── clean ──────────────────────────────────────────────────────────────────
    p_clean = sub.add_parser("clean", help="Clean and deduplicate one or more NPI CSV files.")
    p_clean.add_argument("files", nargs="+", help="NPI CSV file(s) to clean.")
    p_clean.add_argument("-o", "--output", help="Output path (only used when a single file is given).")
    p_clean.add_argument("--active-only", action="store_true", help="Keep only active providers (status=A).")
    p_clean.add_argument("--no-dedup", action="store_true", help="Skip NPI deduplication.")
    p_clean.add_argument("--sort", help="Comma-separated sort keys (default: last_name,first_name,primary_city).")

    # ── merge ──────────────────────────────────────────────────────────────────
    p_merge = sub.add_parser("merge", help="Merge multiple NPI CSV files into one deduplicated file.")
    p_merge.add_argument("files", nargs="+", help="NPI CSV file(s) to merge.")
    p_merge.add_argument("-o", "--output", default="merged_providers.csv", help="Output CSV path.")
    p_merge.add_argument("--active-only", action="store_true", help="Keep only active providers (status=A).")
    p_merge.add_argument("--sort", help="Comma-separated sort keys.")

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
    if args.command == "clean":
        cmd_clean(args)
    elif args.command == "merge":
        cmd_merge(args)
    elif args.command == "crossref":
        cmd_crossref(args)
    print(f"\n{'─'*60}\n  Done.\n")


if __name__ == "__main__":
    main()
