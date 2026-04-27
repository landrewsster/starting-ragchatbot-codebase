#!/usr/bin/env python3
"""
check_round2.py

Inspect MailingListRound2.xlsx for:
  1. Duplicate entries (by NPI, and by name)
  2. Incomplete addresses (missing address, city, state, or zip)
  3. Name column issues:
       - Full name in column B instead of first name only
       - Likely first name in column A instead of last name

Usage:
    python3 check_round2.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
INPUT_FILE  = BASE / "MailingListRound2.xlsx"
OUTPUT_FILE = BASE / "round2_check.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
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

STRIP_SUFFIXES = re.compile(
    r"\b(jr|sr|ii|iii|iv|md|do|dpm|dds|phd|np|pa|rn|aprn|cnp|cnm|cns|esq)\.?\s*$",
    re.IGNORECASE,
)

# Common first names — used to flag likely first names in the last-name column
COMMON_FIRST_NAMES = {
    "james","john","robert","michael","william","david","richard","joseph",
    "thomas","charles","christopher","daniel","matthew","anthony","mark",
    "donald","steven","paul","andrew","joshua","kenneth","kevin","brian",
    "george","timothy","ronald","edward","jason","jeffrey","ryan","gary",
    "jacob","nicholas","eric","jonathan","stephen","larry","justin","scott",
    "brandon","benjamin","samuel","raymond","frank","gregory","raymond",
    "mary","patricia","linda","barbara","elizabeth","jennifer","maria",
    "susan","margaret","dorothy","lisa","nancy","karen","betty","helen",
    "sandra","donna","carol","ruth","sharon","michelle","laura","sarah",
    "kimberly","deborah","jessica","shirley","cynthia","angela","melissa",
    "brenda","amy","anna","rebecca","virginia","kathleen","pamela","martha",
    "debra","amanda","stephanie","carolyn","christine","marie","janet",
    "catherine","frances","ann","joyce","diane","alice","julie","heather",
    "teresa","doris","gloria","evelyn","jean","cheryl","mildred","katherine",
    "joan","ashley","judith","rose","janice","kelly","nicole","judy",
    "christina","kathy","theresa","beverly","denise","tammy","irene","jane",
    "lori","rachel","marilyn","andrea","kathryn","louise","sara","anne",
    "jacqueline","wanda","bonnie","julia","ruby","lois","tina","phyllis",
    "norma","paula","diana","annie","lillian","emily","robin","peggy",
    "crystal","gladys","rita","dawn","connie","florence","tracy","edna",
    "tiffany","carmen","rosa","cindy","grace","wendy","victoria","edith",
    "kim","sherry","sylvia","josephine","thelma","shannon","sheila","ethel",
    "ellen","elaine","marjorie","carrie","charlotte","monica","esther",
    "pauline","emma","juanita","anita","rhonda","hazel","amber","eva",
    "debbie","april","leslie","clara","lucille","jamie","joanne","eleanor",
    "valerie","danielle","megan","alicia","suzanne","michele","gail","bertha",
    "darlene","veronica","jill","erin","geraldine","lauren","cathy","joann",
    "lorraine","lynn","sally","regina","erica","beatrice","dolores","bernice",
    "audrey","yvonne","annette","june","samantha","marion","dana","stacy",
    "ana","renee","ida","vivian","roberta","holly","brittany","melanie",
    "loretta","yolanda","jeanette","laurie","katie","kristen","vanessa",
    "alma","sue","elsie","beth","jeanne",
}

# ── Load ──────────────────────────────────────────────────────────────────────
print(f"\nLoading: {INPUT_FILE.name}")
try:
    xl = pd.ExcelFile(INPUT_FILE)
    print(f"  Sheets: {xl.sheet_names}")
    df = xl.parse(xl.sheet_names[0], dtype=str).fillna("")
except FileNotFoundError:
    raise SystemExit(f"ERROR: {INPUT_FILE} not found")

print(f"  {len(df)} rows | {len(df.columns)} columns")
print(f"  Columns: {list(df.columns)}")

# ── Identify columns ──────────────────────────────────────────────────────────
# Column A is the first column, B is the second
col_a = df.columns[0]   # expected: last name
col_b = df.columns[1] if len(df.columns) > 1 else None  # expected: first name

npi_col  = find_col(df, ["npi", "NPI"])
addr_col = find_col(df, ["address", "address1", "address_1", "addr1", "addr",
                          "Delivery Address", "primary_address_1", "AddressLine1",
                          "mailing_address_1"])
city_col = find_col(df, ["city", "City", "primary_city", "mailing_city"])
state_col= find_col(df, ["state", "State", "primary_state", "mailing_state"])
zip_col  = find_col(df, ["zip", "Zip", "zip5", "postal_code", "mailing_postal_code",
                          "primary_zip", "mailing_zip"])

print(f"\n  col_A (last?)={col_a}  col_B (first?)={col_b}")
print(f"  npi={npi_col}  addr={addr_col}  city={city_col}  state={state_col}  zip={zip_col}")

# ── Flag issues ───────────────────────────────────────────────────────────────
issues = []

for i, row in df.iterrows():
    row_issues = []
    val_a = norm(row[col_a])
    val_b = norm(row[col_b]) if col_b else ""

    # Full name in column B (contains a space → likely "First Last" or "First M Last")
    if val_b and " " in val_b.strip():
        row_issues.append("fullname_in_col_B")

    # First name in column A (single word matching common first names)
    if val_a and " " not in val_a and val_a in COMMON_FIRST_NAMES:
        row_issues.append("firstname_in_col_A")

    # Incomplete address
    addr  = norm(row[addr_col])  if addr_col  else ""
    city  = norm(row[city_col])  if city_col  else ""
    state = norm(row[state_col]) if state_col else ""
    zipv  = norm(row[zip_col])   if zip_col   else ""

    if not addr:
        row_issues.append("missing_address")
    if not city:
        row_issues.append("missing_city")
    if not state:
        row_issues.append("missing_state")
    if not zipv:
        row_issues.append("missing_zip")

    if row_issues:
        issues.append((i, "|".join(row_issues)))

# ── Duplicates by NPI ─────────────────────────────────────────────────────────
dup_npi_rows = pd.DataFrame()
if npi_col:
    npi_series = df[npi_col].apply(norm)
    valid_npi  = npi_series[npi_series != ""]
    dup_mask   = valid_npi.duplicated(keep=False)
    dup_npi_rows = df[npi_series.isin(valid_npi[dup_mask])].copy()
    dup_npi_rows = dup_npi_rows.sort_values(npi_col)
    print(f"\nDuplicate NPI rows    : {len(dup_npi_rows)}")

# ── Duplicates by name ────────────────────────────────────────────────────────
def name_norm(row):
    a = norm(row[col_a])
    b = norm(row[col_b]) if col_b else ""
    return STRIP_SUFFIXES.sub("", f"{a}|{b}").strip()

df["_name_key"] = df.apply(name_norm, axis=1)
valid_names = df["_name_key"][df["_name_key"] != "|"]
dup_name_mask = valid_names.duplicated(keep=False)
dup_name_rows = df[df["_name_key"].isin(valid_names[dup_name_mask])].copy()
dup_name_rows = dup_name_rows.sort_values("_name_key")
print(f"Duplicate name rows   : {len(dup_name_rows)}")

# ── Build issue dataframe ─────────────────────────────────────────────────────
issue_df = df.loc[[i for i, _ in issues]].copy()
issue_df["_issues"] = [f for _, f in issues]
issue_df = issue_df.sort_values("_issues")

name_issues_df = issue_df[issue_df["_issues"].str.contains(
    "fullname_in_col_B|firstname_in_col_A")].copy()
addr_issues_df = issue_df[issue_df["_issues"].str.contains(
    "missing_address|missing_city|missing_state|missing_zip")].copy()

print(f"Name column issues    : {len(name_issues_df)}")
print(f"Incomplete address    : {len(addr_issues_df)}")
print(f"Total flagged rows    : {len(issue_df)}")

# ── Issue summary ─────────────────────────────────────────────────────────────
from collections import Counter
all_flags = []
for _, flags in issues:
    all_flags.extend(flags.split("|"))
print(f"\nIssue breakdown:")
for flag, cnt in sorted(Counter(all_flags).items(), key=lambda x: -x[1]):
    print(f"  {cnt:4d}  {flag}")

# ── Write output ──────────────────────────────────────────────────────────────
df_clean = df.drop(columns=["_name_key"])
print(f"\nWriting: {OUTPUT_FILE.name}")
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    name_issues_df.drop(columns=["_name_key"], errors="ignore").to_excel(
        writer, sheet_name="name_issues", index=False)
    addr_issues_df.drop(columns=["_name_key"], errors="ignore").to_excel(
        writer, sheet_name="address_issues", index=False)
    dup_npi_rows.to_excel(
        writer, sheet_name="duplicate_npi",  index=False)
    dup_name_rows.drop(columns=["_name_key"], errors="ignore").to_excel(
        writer, sheet_name="duplicate_name", index=False)
    print(f"  name_issues    : {len(name_issues_df)} rows")
    print(f"  address_issues : {len(addr_issues_df)} rows")
    print(f"  duplicate_npi  : {len(dup_npi_rows)} rows")
    print(f"  duplicate_name : {len(dup_name_rows)} rows")

print("\nDone.")
