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
INPUT_FILE  = BASE / "mailed_providers_hs.xlsx"
OUTPUT_FILE = BASE / "mailed_providers_hs.xlsx"   # overwrite in place

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
    # 920 E 28th St — Allina Minneapolis Heart Institute (Abbott Northwestern campus)
    "920 e 28th street":               ("Allina Health - Minneapolis Heart Institute", "Minneapolis"),
    "920 e 28th st":                   ("Allina Health - Minneapolis Heart Institute", "Minneapolis"),
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
    # Lakeview Clinic — Waconia (and satellite locations)
    "333 third street sw":             ("Lakeview Clinic", "Waconia"),
    "333 3rd street sw":               ("Lakeview Clinic", "Waconia"),
    "333 3rd st sw":                   ("Lakeview Clinic", "Waconia"),
    # Welia Health — Mora (formerly Kanabec Hospital area)
    "301 highway 65 n":                ("Welia Health", "Mora"),
    "301 hwy 65 n":                    ("Welia Health", "Mora"),
    # St. Cloud Hospital — CentraCare Health
    "1406 6th avenue n":               ("CentraCare Health - St. Cloud Hospital", "Saint Cloud"),
    "1406 6th ave n":                  ("CentraCare Health - St. Cloud Hospital", "Saint Cloud"),
    # Grand Itasca Clinic and Hospital — Grand Rapids
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
    # department names — not health system names
    r"^(department of|division of|dept\.? of)\b",
    re.IGNORECASE,
)
# Matches names that look like individual providers rather than health systems
_INDIVIDUAL_RE = re.compile(
    r"^dr\.?\s|"                                   # starts with Dr.
    r",\s*(md|do|dpm|dds|np|pa-?c?|rn|aprn|cnp|cnm|cns|phd|dnp|facc|fhrs|mbbs|msph)[\s,]*$|"
    r"\s[-–]\s*(md|do|np|pa|rn|aprn)\s*$|"        # " - MD" or " – DO"
    r"\b(md|do|dnp|facc|fhrs)\s*$",               # ends with bare credential
    re.IGNORECASE,
)

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

print(f"  addr1={addr1_col}  addr2={addr2_col}  city={city_col}")

# ── Step 1: clear bad and individual-provider classifications ─────────────────
cleared_bad = 0
cleared_ind = 0
for i, row in df.iterrows():
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

# ── Step 3: address propagation ───────────────────────────────────────────────
def addr_key(row) -> str:
    addr = safe_addr(row.get(addr1_col)) if addr1_col else ""
    city = safe_addr(row.get(city_col))  if city_col  else ""
    addr_base = _SUITE_RE.sub("", addr).strip().rstrip(",").strip()
    return norm(addr_base) + "|" + norm(city)

addr_hs_votes: dict = {}
for _, row in df[df[HS_COL].apply(norm) != ""].iterrows():
    key = addr_key(row)
    if not key or key == "|":
        continue
    hs      = str(row[HS_COL]).strip()
    hs_city = str(row[HS_CITY_COL]).strip()
    if key not in addr_hs_votes:
        addr_hs_votes[key] = {}
    addr_hs_votes[key][(hs, hs_city)] = addr_hs_votes[key].get((hs, hs_city), 0) + 1

addr_best = {k: max(v, key=v.get) for k, v in addr_hs_votes.items()}

prop_hits = 0
for i, row in df[df[HS_COL].apply(norm) == ""].iterrows():
    key = addr_key(row)
    if key in addr_best:
        hs, hs_city = addr_best[key]
        df.at[i, HS_COL]      = hs
        df.at[i, HS_CITY_COL] = hs_city
        prop_hits += 1
print(f"Propagation classified: {prop_hits}")

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

# ── Step 5: flag likely clinics ───────────────────────────────────────────────
if addr2_col:
    df["_likely_clinic"] = df.apply(
        lambda r: (
            r[HS_COL] in ("Solo Practice", "Group Practice") and
            bool(_SUITE_FLAG.search(safe_addr(r.get(addr2_col))))
        ),
        axis=1,
    ).map({True: "likely_clinic", False: ""})
else:
    df["_likely_clinic"] = ""

n_likely = (df["_likely_clinic"] == "likely_clinic").sum()
print(f"Likely clinic (solo/group with suite): {n_likely}")

print(f"\nResults:")
print(f"  Total          : {len(df)}")
print(f"  Health system  : {(~df[HS_COL].isin(['Solo Practice','Group Practice',''])).sum()}")
print(f"  Group Practice : {group_count}")
print(f"  Solo Practice  : {solo_count}")
print(f"  Likely clinic (needs review): {n_likely}")

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

# Summary — exclude Solo/Group from the health system list
hs_df = df[~df[HS_COL].isin(["Solo Practice", "Group Practice", ""])].copy()
all_systems = hs_df[HS_COL].value_counts()
print(f"\n  Health systems  (providers | locations | cities):")
for hs, cnt in all_systems.items():
    subset = hs_df[hs_df[HS_COL] == hs]
    cities = subset[HS_CITY_COL].apply(norm).replace("", pd.NA).dropna().unique().tolist()
    cities = [c.title() for c in sorted(cities)]
    n_loc  = len(cities)
    city_str = ", ".join(cities[:5])
    if len(cities) > 5:
        city_str += f" (+{len(cities)-5} more)"
    print(f"    {cnt:4d} providers | {n_loc:2d} locations | {city_str}")
    print(f"         {hs}")

# ── Write ─────────────────────────────────────────────────────────────────────
print(f"\nWriting: {OUTPUT_FILE.name}")
solo_group_df   = df[df[HS_COL].isin(["Solo Practice", "Group Practice"])].copy()
likely_clin_df  = df[df["_likely_clinic"] == "likely_clinic"].copy()
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="mailed_providers_hs", index=False)
    solo_group_df.to_excel(writer, sheet_name="solo_group_review", index=False)
    likely_clin_df.to_excel(writer, sheet_name="likely_clinic_review", index=False)
    print(f"  mailed_providers_hs   : {len(df)} rows")
    print(f"  solo_group_review     : {len(solo_group_df)} rows")
    print(f"  likely_clinic_review  : {len(likely_clin_df)} rows")

print("\nDone.")
