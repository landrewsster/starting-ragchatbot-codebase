#!/usr/bin/env python3
"""
reprocess_hs.py

Apply address overrides and propagation to an existing mailed_providers_hs.xlsx
without making any API calls.  Run this after health_system_lookup.py has
already completed a full API run.

Usage:
    python3 reprocess_hs.py
"""

import re
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE        = Path.home() / "Downloads" / "CRC MDH Project" / "Current Mailing Files"
INPUT_FILE  = BASE / "mailed_providers_hs_reprocessed.xlsx"
OUTPUT_FILE = BASE / "mailed_providers_hs_reprocessed.xlsx"

HS_COL      = "health_system"
HS_CITY_COL = "health_system_city"

# ── Manual address overrides ──────────────────────────────────────────────────
# Add rows here as you discover gaps.  Key = normalized street address.
MANUAL_ADDR_OVERRIDES = {
    # 420 Delaware St SE — UMN research/clinical offices
    "420 delaware street se":          ("University of Minnesota Health", "Minneapolis"),
    "420 delaware st se":              ("University of Minnesota Health", "Minneapolis"),
    "420 delaware street":             ("University of Minnesota Health", "Minneapolis"),
    "420 delaware st":                 ("University of Minnesota Health", "Minneapolis"),
    # 717 Delaware St SE — UMN (Division of Nephrology etc.)
    "717 delaware street se":          ("University of Minnesota Health", "Minneapolis"),
    "717 delaware st se":              ("University of Minnesota Health", "Minneapolis"),
    # 401 East River Parkway — VCRC / UMN research buildings
    "401 east river parkway":          ("University of Minnesota Health", "Minneapolis"),
    "401 e river parkway":             ("University of Minnesota Health", "Minneapolis"),
    "401 east river pkwy":             ("University of Minnesota Health", "Minneapolis"),
    # 920 E 28th St — Allina Health Bandana Square Clinic
    "920 e 28th street":               ("Allina Health Bandana Square Clinic", "Minneapolis"),
    "920 e. 28th st.":                 ("Allina Health Bandana Square Clinic", "Minneapolis"),
    "920 east 28th st.":               ("Allina Health Bandana Square Clinic", "Minneapolis"),
    "920 e 28th st":                   ("Allina Health Bandana Square Clinic", "Minneapolis"),
    # 1021 Bandana Blvd E — Allina Health Bandana Square Clinic
    "1021 bandana blvd e":             ("Allina Health Bandana Square Clinic", "Saint Paul"),
    "1021 bandana boulevard e":        ("Allina Health Bandana Square Clinic", "Saint Paul"),
    # 8100 W 78th St — Abbott Northwestern General Medicine Associates (Edina)
    "8100 w 78th street":              ("Abbott Northwestern General Medicine Associates", "Edina"),
    "8100 west 78th street":           ("Abbott Northwestern General Medicine Associates", "Edina"),
    "8100 w 78th st":                  ("Abbott Northwestern General Medicine Associates", "Edina"),
    "8100 west 78th st.":              ("Abbott Northwestern General Medicine Associates", "Edina"),
    # 2001 Blaisdell Ave S — Park Nicollet Clinic Minneapolis
    "2001 blaisdell avenue s":         ("Park Nicollet Clinic", "Minneapolis"),
    "2001 blaisdell ave s":            ("Park Nicollet Clinic", "Minneapolis"),
    # 333 Smith Ave N — United Hospital (Allina Health)
    "333 smith avenue n":              ("United Hospital", "Saint Paul"),
    "333 smith ave n":                 ("United Hospital", "Saint Paul"),
    # 1 / One Veterans Dr — VA Medical Center Minneapolis
    "1 veterans drive":                ("VA Medical Center", "Minneapolis"),
    "1 veterans dr":                   ("VA Medical Center", "Minneapolis"),
    "one veterans drive":              ("VA Medical Center", "Minneapolis"),
    "one veterans dr":                 ("VA Medical Center", "Minneapolis"),
    # Lakeview Clinic — Waconia
    "333 third street sw":             ("Lakeview Clinic", "Waconia"),
    "333 3rd street sw":               ("Lakeview Clinic", "Waconia"),
    "333 3rd st sw":                   ("Lakeview Clinic", "Waconia"),
    # Welia Health — Mora
    "301 highway 65 n":                ("Welia Health", "Mora"),
    "301 hwy 65 n":                    ("Welia Health", "Mora"),
    # CentraCare Health — St. Cloud Hospital
    "1406 6th avenue n":               ("CentraCare Health - St. Cloud Hospital", "Saint Cloud"),
    "1406 6th ave n":                  ("CentraCare Health - St. Cloud Hospital", "Saint Cloud"),
    # Grand Itasca Clinic and Hospital
    "1601 golf course road":           ("Grand Itasca Clinic and Hospital", "Grand Rapids"),
    "1601 golf course rd":             ("Grand Itasca Clinic and Hospital", "Grand Rapids"),
    # Perham Health
    "1000 coney street w":             ("Perham Health", "Perham"),
    "1000 coney st w":                 ("Perham Health", "Perham"),
    # Lakeview Hospital — Stillwater (HealthPartners)
    "927 w churchill street":          ("Lakeview Hospital", "Stillwater"),
    "927 w churchill st":              ("Lakeview Hospital", "Stillwater"),
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s).lower().strip())

def safe_addr(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def find_col(df, candidates):
    low = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns:
            return c
        if c.lower() in low:
            return low[c.lower()]
    return None

_SUITE_RE = re.compile(
    r"\b(suite|ste|floor|fl\b|room|rm\b|apt|unit\b|bldg|#)\s*[\w-]+",
    re.IGNORECASE,
)
_SUITE_FLAG = re.compile(
    r"\b(suite|ste|floor|fl\b|room|rm\b|apt|unit\b|#)\b", re.IGNORECASE
)
_BAD_HS_RE = re.compile(
    r"^(nan|none|n/?a|home)\s*$|"
    r"^(suite|ste|floor|fl\b|room|rm\b|apt|unit\b|bldg|building|"
    r"c/o|attention|attn|united states postal|usps|post office)\b|"
    # internal mail/room codes
    r"^(mmc|mail route|mail stop|vcrc)\b|"
    r"mr#\s*\d+|"
    r"\bmail\s+(delivery\s+)?code\b|"
    # department names at start OR embedded after a comma
    r"^(department of|division of|dept\.? of)\b|"
    r",\s*(department of|division of)\b|"
    # USPS bulk mail sort codes — start with 3+ asterisks
    r"^\*{3,}|"
    # street address fragments
    r"^at\s+\w",
    re.IGNORECASE,
)
# Matches names that look like individual providers rather than health systems.
# "PA" is intentionally excluded — it also means Professional Association (org suffix).
_INDIVIDUAL_RE = re.compile(
    r"^dr\.?\s|"                                   # starts with Dr.
    r",\s*(md|do|dpm|dds|np|rn|aprn|cnp|cnm|cns|phd|dnp|facc|fhrs|mbbs|msph)[\s,]*$|"
    r"\s[-–]\s*(md|do|np|rn|aprn)\s*$|"           # " - MD" or " – DO"
    r"\b(md|do|dnp|facc|fhrs)\s*$",               # ends with bare credential
    re.IGNORECASE,
)

# Regex-based address overrides — checked against addr1 AND addr2.
# Each entry: (compiled pattern, health_system, city)
ADDR_PATTERN_OVERRIDES = [
    # Delaware St addresses → UMN campus
    (re.compile(r"\bdelaware\b", re.IGNORECASE),       "University of Minnesota Health", "Minneapolis"),
    # VCRC (Variety Club Research Center) in any address field → UMN
    (re.compile(r"\bvcrc\b", re.IGNORECASE),           "University of Minnesota Health", "Minneapolis"),
    # East River Parkway → UMN campus buildings
    (re.compile(r"\beast river parkway\b", re.IGNORECASE), "University of Minnesota Health", "Minneapolis"),
]

_DEPT_NAME_RE = re.compile(r"^(department of|division of|dept\.? of)\b", re.IGNORECASE)

# ── Load ──────────────────────────────────────────────────────────────────────
print(f"\nLoading: {INPUT_FILE.name}")
df = pd.read_excel(INPUT_FILE, sheet_name="mailed_providers_hs", dtype=str)
df = df.fillna("")
print(f"  {len(df)} rows")

if HS_COL not in df.columns:
    df[HS_COL] = ""
if HS_CITY_COL not in df.columns:
    df[HS_CITY_COL] = ""

addr1_col = find_col(df, ["_check_addr", "primary_address_1", "Delivery Address",
                           "work_address_1", "Alternate 1 Address"])
addr2_col = find_col(df, ["primary_address_2", "work_address_2"])
city_col  = find_col(df, ["_check_city", "primary_city", "City", "work_city"])
npi_col   = find_col(df, ["npi", "NPI", "national_provider_identifier"])
last_col  = find_col(df, ["last_name", "Last Name", "lname", "_check_last"])
first_col = find_col(df, ["first_name", "First Name", "fname", "_check_first"])

print(f"  addr1={addr1_col}  addr2={addr2_col}  city={city_col}  npi={npi_col}")

# ── Step 0: apply user annotations from review column ────────────────────────
# "Opt. Endorsement Line" (column K) written during review → confirmed
# health_system.  Also merges annotations written into any review sheet
# (solo_group_review, likely_clinic_review) back into the main df.
# Matching priority: NPI → last|first|addr1 composite key.
USER_OVERRIDE_COL  = "Opt. Endorsement Line"
REVIEW_SHEETS      = ["solo_group_review", "likely_clinic_review"]
user_col = find_col(df, [USER_OVERRIDE_COL])
df["_user_reviewed"] = ""

def _row_key(row, npi_c, last_c, first_c, addr_c):
    """Return a unique key for a row: NPI if present, else name+addr."""
    npi = safe_addr(row.get(npi_c)) if npi_c else ""
    if npi:
        return "npi:" + npi
    last  = norm(row.get(last_c))  if last_c  else ""
    first = norm(row.get(first_c)) if first_c else ""
    addr  = norm(row.get(addr_c))  if addr_c  else ""
    return f"name:{last}|{first}|{addr}" if (last or first) and addr else ""

# Build key → index map for main df
main_key_to_idx: dict[str, int] = {}
for i, row in df.iterrows():
    k = _row_key(row, npi_col, last_col, first_col, addr1_col)
    if k:
        main_key_to_idx[k] = i

# Merge annotations from review sheets
xl = pd.ExcelFile(INPUT_FILE)
for sheet in REVIEW_SHEETS:
    if sheet not in xl.sheet_names:
        continue
    rdf    = xl.parse(sheet, dtype=str).fillna("")
    r_npi  = find_col(rdf, ["npi", "NPI", "national_provider_identifier"])
    r_last = find_col(rdf, ["last_name", "Last Name", "lname", "_check_last"])
    r_frst = find_col(rdf, ["first_name", "First Name", "fname", "_check_first"])
    r_addr = find_col(rdf, ["_check_addr", "primary_address_1", "Delivery Address",
                             "work_address_1", "Alternate 1 Address"])
    r_user = find_col(rdf, [USER_OVERRIDE_COL])
    if not r_user:
        continue
    merged = 0
    for _, rrow in rdf.iterrows():
        ann = safe_addr(rrow.get(r_user))
        if not ann:
            continue
        k = _row_key(rrow, r_npi, r_last, r_frst, r_addr)
        if not k or k not in main_key_to_idx:
            continue
        i = main_key_to_idx[k]
        if safe_addr(df.at[i, user_col] if user_col else ""):
            continue  # already annotated in main sheet — don't overwrite
        if user_col:
            df.at[i, user_col] = ann
        merged += 1
    if merged:
        print(f"  Merged {merged} annotations from '{sheet}' sheet")

def _apply_user_col(df):
    if not user_col:
        return 0
    hits = 0
    for i, row in df.iterrows():
        val = safe_addr(row.get(user_col))
        if not val:
            continue
        low = val.lower().strip()
        if low in ("solo practice", "solo"):
            val = "Solo Practice"
        elif low in ("group practice", "group"):
            val = "Group Practice"
        df.at[i, HS_COL]           = val
        df.at[i, "_user_reviewed"] = "yes"
        city = safe_addr(row.get(city_col)) if city_col else ""
        if not norm(row.get(HS_CITY_COL)):
            df.at[i, HS_CITY_COL] = city
        hits += 1
    return hits

user_hits = _apply_user_col(df)
if user_col:
    print(f"User annotations applied ({USER_OVERRIDE_COL!r}): {user_hits}")
else:
    print(f"  (no '{USER_OVERRIDE_COL}' column found — skipping Step 0)")

# ── Step 1: clear bad and individual-provider classifications ─────────────────
cleared_bad = 0
cleared_ind = 0
for i, row in df.iterrows():
    if row.get("_user_reviewed") == "yes":
        continue  # never clear user-confirmed classifications
    hs = safe_addr(row[HS_COL])
    if not hs:
        continue
    if _BAD_HS_RE.match(hs):
        df.at[i, HS_COL]      = ""
        df.at[i, HS_CITY_COL] = ""
        cleared_bad += 1
    elif _INDIVIDUAL_RE.search(hs):
        df.at[i, HS_COL]      = ""
        df.at[i, HS_CITY_COL] = ""
        cleared_ind += 1
print(f"\nCleared {cleared_bad} bad classifications")
print(f"Cleared {cleared_ind} individual provider name classifications")

# ── Step 2: apply manual overrides ────────────────────────────────────────────
override_hits = 0
for i, row in df.iterrows():
    if norm(row[HS_COL]):
        continue
    addr1 = safe_addr(row.get(addr1_col)) if addr1_col else ""
    key   = norm(addr1)
    if key in MANUAL_ADDR_OVERRIDES:
        hs, hs_city = MANUAL_ADDR_OVERRIDES[key]
        df.at[i, HS_COL]      = hs
        df.at[i, HS_CITY_COL] = hs_city
        override_hits += 1
print(f"Manual overrides applied: {override_hits}")

# ── Step 2b: regex-based address pattern overrides ────────────────────────────
pattern_hits = 0
for i, row in df.iterrows():
    if norm(row[HS_COL]):
        continue
    addr1 = safe_addr(row.get(addr1_col)) if addr1_col else ""
    addr2 = safe_addr(row.get(addr2_col)) if addr2_col else ""
    # When addr1 is a department/division name the real street may be in addr2
    search_addrs = [addr1]
    if _DEPT_NAME_RE.match(addr1):
        search_addrs.append(addr2)
    for pattern, hs, hs_city in ADDR_PATTERN_OVERRIDES:
        if any(pattern.search(a) for a in search_addrs if a):
            df.at[i, HS_COL]      = hs
            df.at[i, HS_CITY_COL] = hs_city
            pattern_hits += 1
            break
print(f"Pattern overrides applied: {pattern_hits}")

# ── Step 3: address propagation ───────────────────────────────────────────────
def addr_key(row) -> str:
    addr = safe_addr(row.get(addr1_col)) if addr1_col else ""
    city = safe_addr(row.get(city_col))  if city_col  else ""
    addr_base = _SUITE_RE.sub("", addr).strip().rstrip(",").strip()
    return norm(addr_base) + "|" + norm(city)

_PRACTICE_LABELS = {"Solo Practice", "Group Practice"}

addr_votes_named: dict = {}  # named health systems only (not Solo/Group)
addr_votes_all:   dict = {}  # all health systems

for _, row in df[df[HS_COL].apply(norm) != ""].iterrows():
    key = addr_key(row)
    if not key or key == "|":
        continue
    hs      = str(row[HS_COL]).strip()
    hs_city = str(row[HS_CITY_COL]).strip()
    if _BAD_HS_RE.match(hs) or _INDIVIDUAL_RE.search(hs):
        continue  # don't let bad values spread via propagation
    addr_votes_all.setdefault(key, {})
    addr_votes_all[key][(hs, hs_city)] = addr_votes_all[key].get((hs, hs_city), 0) + 1
    if hs not in _PRACTICE_LABELS:
        addr_votes_named.setdefault(key, {})
        addr_votes_named[key][(hs, hs_city)] = addr_votes_named[key].get((hs, hs_city), 0) + 1

addr_best_named = {k: max(v, key=v.get) for k, v in addr_votes_named.items()}
addr_best_all   = {k: max(v, key=v.get) for k, v in addr_votes_all.items()}

prop_hits = 0
# Pass 1: fill rows with no health system yet
for i, row in df[df[HS_COL].apply(norm) == ""].iterrows():
    key = addr_key(row)
    if key in addr_best_all:
        hs, hs_city = addr_best_all[key]
        df.at[i, HS_COL]      = hs
        df.at[i, HS_CITY_COL] = hs_city
        prop_hits += 1

# Pass 2: upgrade Solo/Group Practice rows when a named system exists at the
# same address — propagates user annotations across all rows at an address
upgrade_hits = 0
for i, row in df.iterrows():
    if row.get("_user_reviewed") == "yes":
        continue
    if row[HS_COL] not in _PRACTICE_LABELS:
        continue
    key = addr_key(row)
    if key in addr_best_named:
        hs, hs_city = addr_best_named[key]
        df.at[i, HS_COL]      = hs
        df.at[i, HS_CITY_COL] = hs_city
        upgrade_hits += 1
        prop_hits += 1

print(f"Propagation classified : {prop_hits - upgrade_hits}")
print(f"Propagation upgraded   : {upgrade_hits}  (Solo/Group → named system)")


# ── Step 4: label remaining unclassified as Solo or Group Practice ────────────
# Count how many providers share each base address across the whole dataset
addr_provider_count: dict[str, int] = {}
for _, row in df.iterrows():
    k = addr_key(row)
    if k and k != "|":
        addr_provider_count[k] = addr_provider_count.get(k, 0) + 1

practice_hits = 0
for i, row in df[df[HS_COL].apply(norm) == ""].iterrows():
    k     = addr_key(row)
    count = addr_provider_count.get(k, 1)
    label = "Group Practice" if count >= 2 else "Solo Practice"
    city  = safe_addr(row.get(city_col)) if city_col else ""
    df.at[i, HS_COL]      = label
    df.at[i, HS_CITY_COL] = city
    practice_hits += 1

solo_count  = (df[HS_COL] == "Solo Practice").sum()
group_count = (df[HS_COL] == "Group Practice").sum()
print(f"Labeled as Solo Practice : {solo_count}")
print(f"Labeled as Group Practice: {group_count}")

# ── Step 5: flag suite addresses within solo/group rows ───────────────────────
# Suite in addr2 means office building — could be a clinic or a legitimate solo
# practice. Flagged here for reference; review via solo_group_review sheet.
if addr2_col:
    df["_likely_clinic"] = df.apply(
        lambda r: (
            r[HS_COL] in ("Solo Practice", "Group Practice") and
            bool(_SUITE_FLAG.search(safe_addr(r.get(addr2_col))))
        ),
        axis=1,
    ).map({True: "suite_address", False: ""})
else:
    df["_likely_clinic"] = ""

n_suite = (df["_likely_clinic"] == "suite_address").sum()

print(f"\nResults:")
print(f"  Total          : {len(df)}")
print(f"  Health system  : {(~df[HS_COL].isin(['Solo Practice','Group Practice',''])).sum()}")
print(f"  Group Practice : {group_count}")
print(f"  Solo Practice  : {solo_count}")
print(f"  Suite address (solo/group, for reference): {n_suite}")

for label in ["Group Practice", "Solo Practice"]:
    subset = df[df[HS_COL] == label]
    cities = subset[HS_CITY_COL].apply(norm).replace("", pd.NA).dropna().unique().tolist()
    cities = sorted([c.title() for c in cities if c])
    n_loc  = len(cities)
    city_str = ", ".join(cities[:8])
    if len(cities) > 8:
        city_str += f" (+{len(cities)-8} more)"
    print(f"\n  {label}: {len(subset)} providers across {n_loc} cities")
    if city_str:
        print(f"    {city_str}")

# ── Write ─────────────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
# solo_group_review includes the _likely_clinic column so the user can filter
# by "suite_address" to see office-building rows that may warrant a closer look.
solo_group_df = df[df[HS_COL].isin(["Solo Practice", "Group Practice"])].copy()
hs_df         = df[~df[HS_COL].isin(["Solo Practice", "Group Practice", ""])].copy()

# Build health system summary table
hs_summary_rows = []
for hs, cnt in hs_df[HS_COL].value_counts().items():
    subset = hs_df[hs_df[HS_COL] == hs]
    cities = sorted({c.title() for c in subset[HS_CITY_COL].apply(norm) if c})
    hs_summary_rows.append({
        "health_system":    hs,
        "provider_count":   cnt,
        "location_count":   len(cities),
        "cities":           ", ".join(cities),
    })
hs_summary_df = pd.DataFrame(hs_summary_rows)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="mailed_providers_hs", index=False)
    hs_summary_df.to_excel(writer, sheet_name="health_system_summary", index=False)
    solo_group_df.to_excel(writer, sheet_name="solo_group_review", index=False)
    print(f"  mailed_providers_hs   : {len(df)} rows")
    print(f"  health_system_summary : {len(hs_summary_df)} health systems")
    print(f"  solo_group_review     : {len(solo_group_df)} rows  ({n_suite} with suite address)")

# ── Top 30 health systems — printed last so it's always visible ───────────────
top30 = hs_df[HS_COL].value_counts().head(75)
print(f"\nTop 75 health systems  (providers | locations | cities):")
for hs, cnt in top30.items():
    subset = hs_df[hs_df[HS_COL] == hs]
    cities = subset[HS_CITY_COL].apply(norm).replace("", pd.NA).dropna().unique().tolist()
    cities = [c.title() for c in sorted(cities)]
    city_str = ", ".join(cities[:5])
    if len(cities) > 5:
        city_str += f" (+{len(cities)-5} more)"
    print(f"  {cnt:4d} providers | {len(cities):2d} locations | {city_str}")
    print(f"       {hs}")

print("\nDone.")
