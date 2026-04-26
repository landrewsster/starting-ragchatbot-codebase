#!/usr/bin/env python3
"""
health_system_lookup.py

Classify providers by health system using a three-phase approach:

  Phase 1  — Reference file lookup (name match → address match → addr1 org detection)
  Phase 1b — Provider name Google Places search (disambiguates same-address/diff-org cases)
  Phase 2  — Google Places API for remaining unique unclassified addresses
  Phase 3  — Write results + summary table

Populates columns:
  health_system      (col O)
  health_system_city (col P)

API key:
  Set environment variable before running:
    export GOOGLE_PLACES_API_KEY="your-key-here"

Usage:
    python3 health_system_lookup.py
"""

import os
import re
import sys
import time
from pathlib import Path

import requests
import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE          = Path.home() / "Downloads" / "CRC MDH Project"
MAILING_BASE  = BASE / "Current Mailing Files"
PROVIDER_BASE = BASE / "ProviderDataFiles"

INPUT_FILE  = MAILING_BASE / "gold_vs_mailing_check.xlsx"
INPUT_SHEET = "mailed_providers"

# Both files contribute to the address → health system reference lookup
REF_FILES = [
    PROVIDER_BASE / "GHmatched_phyPA_LMH.xlsx",
    PROVIDER_BASE / "Mailing List for Printing Services copy_columns_corrected_orgs.xlsx",
]

OUTPUT_FILE = MAILING_BASE / "mailed_providers_hs.xlsx"

# ── API config ───────────────────────────────────────────────────────────────
API_KEY = os.environ.get("GOOGLE_PLACES_API_KEY", "")
PLACES_URL = "https://places.googleapis.com/v1/places:searchText"
API_DELAY  = 0.15   # seconds between API calls

# ── Classification patterns ──────────────────────────────────────────────────
MEDICAL_RE = re.compile(
    r"clinic|hospital|medical|health|care|wellness|rehab|therapy|therapist|"
    r"surgery|surgical|ortho|cardio|oncol|pediatr|neuro|psych|mental|dental|"
    r"vision|eye|family practice|urgent care|emergency|pharmacy|pharma|"
    r"clinica|salud|centro|community|associates|group|practice|services|urgent",
    re.IGNORECASE
)

REALESTATE_RE = re.compile(
    r"airbnb|vrbo|zillow|trulia|redfin|realtor|keller williams|coldwell|"
    r"century 21|re/max|remax|sotheby|real estate|realty|properties|"
    r"apartments|apartment|condo|rental|rentals|suites|inn|hotel|motel|"
    r"lodge|resort|vacation",
    re.IGNORECASE
)

# ── Helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if pd.isna(s) or s is None:
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

def is_po_box(addr: str) -> bool:
    return bool(re.match(r"^\s*p\.?\s*o\.?\s*box\b", addr, re.IGNORECASE))

def is_org_name(addr1: str) -> bool:
    """Return True if addr1 looks like an org name (starts with letter, not PO Box)."""
    if not addr1:
        return False
    if is_po_box(addr1):
        return False
    return bool(re.match(r"^[a-zA-Z]", addr1.strip()))

# ── Google Places API call ───────────────────────────────────────────────────
def places_search(query: str) -> dict | None:
    """Search Google Places API v1. Returns first result dict or None."""
    if not API_KEY:
        return None
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": API_KEY,
        "X-Goog-FieldMask": "places.displayName,places.formattedAddress,places.types,places.primaryType",
    }
    payload = {"textQuery": query, "maxResultCount": 1}
    try:
        resp = requests.post(PLACES_URL, json=payload, headers=headers, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        places = data.get("places", [])
        return places[0] if places else None
    except Exception as e:
        print(f"    API error for '{query}': {e}")
        return None

MEDICAL_TYPES = {
    "hospital", "doctor", "pharmacy", "health", "dentist",
    "physiotherapist", "medical_lab", "drugstore",
}

def classify_place(place: dict) -> tuple[str, str]:
    """Return (health_system_name, city) from a Places API result."""
    if not place:
        return "", ""
    name         = place.get("displayName", {}).get("text", "")
    address      = place.get("formattedAddress", "")
    types        = place.get("types", [])
    primary_type = place.get("primaryType", "")

    # Extract city from formatted address
    city = ""
    parts = address.split(",")
    if len(parts) >= 3:
        city = parts[-3].strip()
    elif len(parts) >= 2:
        city = parts[-2].strip()

    if REALESTATE_RE.search(name):
        return "REAL_ESTATE", city

    # Health system or medical facility by name keywords
    if MEDICAL_RE.search(name):
        return name, city

    # Catch orgs whose names don't include obvious medical words
    # but whose Places type is clearly medical
    if primary_type in MEDICAL_TYPES or any(t in MEDICAL_TYPES for t in types):
        return name, city

    return "", city

# ── Load reference files ──────────────────────────────────────────────────────
ref_lookup = {}   # norm_addr → (health_system, health_system_city)

print(f"\nLoading reference files ...")
for ref_path in REF_FILES:
    print(f"  {ref_path.name}")
    try:
        ref_df = pd.read_excel(ref_path, dtype=str)
    except FileNotFoundError:
        print(f"    WARNING: not found, skipping")
        continue

    ref_df = ref_df.fillna("")

    # Detect health system columns — GHmatched uses "Health System Name/City";
    # mailing list file uses "health_system" / "health_system_city" (or col index 14/15)
    hs_c   = find_col(ref_df, ["Health System Name", "health_system", "HealthSystem"])
    city_c = find_col(ref_df, ["Health System City", "health_system_city", "HealthSystemCity"])
    if not hs_c and len(ref_df.columns) > 14:
        hs_c   = ref_df.columns[14]
    if not city_c and len(ref_df.columns) > 15:
        city_c = ref_df.columns[15]

    if not hs_c:
        print(f"    WARNING: no health system column found, skipping")
        continue

    added = 0
    for addr_candidate in ["work_address_1", "mailing_address_1", "home_address_1",
                            "primary_address_1", "Alternate 1 Address", "npi_primary_address_1"]:
        col = find_col(ref_df, [addr_candidate])
        if not col:
            continue
        for _, row in ref_df.iterrows():
            hs   = str(row[hs_c]).strip()
            city = str(row[city_c]).strip() if city_c else ""
            if hs and hs.lower() not in ("nan", ""):
                key = norm(row[col])
                if key and key not in ref_lookup:
                    ref_lookup[key] = (hs, city)
                    added += 1
    print(f"    +{added} address entries  (total: {len(ref_lookup)})")

# ── Load input (mailed_providers sheet) ──────────────────────────────────────
print(f"\nLoading input: {INPUT_FILE.name}  sheet='{INPUT_SHEET}'")
try:
    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET, dtype=str)
except FileNotFoundError:
    sys.exit(f"ERROR: {INPUT_FILE} not found — run check_gold_vs_mailing.py first")
df = df.fillna("")
print(f"  {len(df)} rows | {len(df.columns)} columns")

# Add output columns if not already present
if "health_system" not in df.columns:
    df["health_system"] = ""
if "health_system_city" not in df.columns:
    df["health_system_city"] = ""

hs_col      = "health_system"
hs_city_col = "health_system_city"

# Detect key columns
last_col   = find_col(df, ["last_name", "LastName"])
first_col  = find_col(df, ["first_name", "FirstName"])
addr1_col  = find_col(df, ["primary_address_1", "work_address_1", "Alternate 1 Address",
                            "npi_primary_address_1", "_addr1"])
addr2_col  = find_col(df, ["primary_address_2", "work_address_2", "npi_primary_address_2"])
addr3_col  = find_col(df, ["address_line3", "primary_address_3"])
city_col   = find_col(df, ["primary_city", "work_city", "City", "npi_primary_city", "_city"])
zip_col    = find_col(df, ["zip5", "primary_postal_code", "ZIP+4", "npi_primary_zip", "_zip"])

print(f"  Name: last={last_col} first={first_col}")
print(f"  Addr: addr1={addr1_col} city={city_col} zip={zip_col}")
print(f"  Output cols: {hs_col}, {hs_city_col}")

# ── Phase 1: Reference file lookup ───────────────────────────────────────────
print(f"\nPhase 1: Reference file lookup ...")
phase1_hits = 0

for i, row in df.iterrows():
    # Skip if already classified
    if norm(row.get(hs_col, "")):
        continue

    addr1 = str(row.get(addr1_col, "")).strip() if addr1_col else ""
    addr2 = str(row.get(addr2_col, "")).strip() if addr2_col else ""
    addr3 = str(row.get(addr3_col, "")).strip() if addr3_col else ""
    city  = str(row.get(city_col,  "")).strip() if city_col  else ""

    # Step 1a: addr1 org detection
    if is_org_name(addr1):
        df.at[i, hs_col]      = addr1
        df.at[i, hs_city_col] = city
        phase1_hits += 1
        continue

    # Step 1b: addr2/addr3 org detection
    for extra in [addr2, addr3]:
        if is_org_name(extra):
            df.at[i, hs_col]      = extra
            df.at[i, hs_city_col] = city
            phase1_hits += 1
            break
    if norm(df.at[i, hs_col]):
        continue

    # Step 1c: reference file address lookup
    for addr in [addr1, addr2, addr3]:
        key = norm(addr)
        if key and key in ref_lookup:
            hs, hs_city = ref_lookup[key]
            df.at[i, hs_col]      = hs
            df.at[i, hs_city_col] = hs_city or city
            phase1_hits += 1
            break

print(f"  Phase 1 classified: {phase1_hits}")
unclassified_after_p1 = df[df[hs_col].apply(norm) == ""].shape[0]
print(f"  Still unclassified: {unclassified_after_p1}")

# ── Phase 1b: Provider name search ───────────────────────────────────────────
# Searches by provider name + city for providers still unclassified after Phase 1.
# Catches cases where multiple orgs share one address (e.g., two different
# clinics in the same building) — address-based lookup can't disambiguate these.
if not API_KEY:
    print(f"\nPhase 1b: Skipped (GOOGLE_PLACES_API_KEY not set)")
else:
    unclassified_mask_1b = df[hs_col].apply(norm) == ""
    p1b_rows = df[unclassified_mask_1b]
    print(f"\nPhase 1b: Name-based search for {len(p1b_rows)} unclassified providers ...")
    phase1b_hits = 0

    # Detect specialty/taxonomy and state columns if present
    taxonomy_col = find_col(df, ["taxonomy_descriptions", "primary_taxonomy_code",
                                  "license_type_desc", "specialty boards"])
    state_col    = find_col(df, ["primary_state", "state", "State"])

    for i, row in p1b_rows.iterrows():
        first = str(row.get(first_col, "")).strip() if first_col else ""
        last  = str(row.get(last_col,  "")).strip() if last_col  else ""
        city  = str(row.get(city_col,  "")).strip() if city_col  else ""
        state = str(row.get(state_col, "")).strip() if state_col else "MN"

        if not last:
            continue

        # Specialty: use up to 3 words (not just first word) to keep query useful
        specialty = ""
        if taxonomy_col:
            tax_val = str(row.get(taxonomy_col, "")).strip()
            if tax_val and tax_val.lower() not in ("nan", ""):
                specialty = " ".join(tax_val.split()[:3])

        # Build query: name + optional specialty + city + state + "physician"
        location = " ".join(p for p in [city, state] if p)
        parts = [p for p in [first, last, specialty, location, "physician"] if p]
        query = " ".join(parts)

        place = places_search(query)
        hs, hs_city = classify_place(place)

        if hs and hs != "REAL_ESTATE":
            df.at[i, hs_col]      = hs
            df.at[i, hs_city_col] = hs_city or city
            phase1b_hits += 1

        time.sleep(API_DELAY)

    print(f"  Phase 1b classified: {phase1b_hits}")
    print(f"  Still unclassified : {df[df[hs_col].apply(norm) == ''].shape[0]}")

# ── Phase 2: Google Places API for unique unclassified addresses ──────────────
if not API_KEY:
    print(f"\nPhase 2: Skipped (GOOGLE_PLACES_API_KEY not set)")
else:
    unclassified_mask = df[hs_col].apply(norm) == ""

    # Collect unique addresses to minimize API calls
    unique_addrs: dict[str, tuple] = {}  # norm_addr_key → (addr1, city, zip)
    for _, row in df[unclassified_mask].iterrows():
        addr1 = str(row.get(addr1_col, "")).strip() if addr1_col else ""
        city  = str(row.get(city_col,  "")).strip() if city_col  else ""
        z     = str(row.get(zip_col,   "")).strip() if zip_col   else ""
        key   = norm(addr1) + "|" + norm(city)
        if key not in unique_addrs:
            unique_addrs[key] = (addr1, city, z)

    print(f"\nPhase 2: API lookup for {len(unique_addrs)} unique addresses ...")
    addr_results: dict[str, tuple] = {}  # norm key → (health_system, hs_city)

    for idx, (key, (addr1, city, z)) in enumerate(unique_addrs.items()):
        if is_po_box(addr1):
            query = f"clinic hospital {city} {z}"
        else:
            query = f"{addr1} {city} {z}"

        place = places_search(query)
        hs, hs_city = classify_place(place)
        addr_results[key] = (hs, hs_city or city)

        if (idx + 1) % 50 == 0:
            print(f"  {idx + 1}/{len(unique_addrs)} addresses processed ...")
        time.sleep(API_DELAY)

    # Apply API results
    phase2_hits = 0
    for i, row in df[unclassified_mask].iterrows():
        addr1 = str(row.get(addr1_col, "")).strip() if addr1_col else ""
        city  = str(row.get(city_col,  "")).strip() if city_col  else ""
        key   = norm(addr1) + "|" + norm(city)
        hs, hs_city = addr_results.get(key, ("", ""))
        if hs and hs != "REAL_ESTATE":
            df.at[i, hs_col]      = hs
            df.at[i, hs_city_col] = hs_city
            phase2_hits += 1

    print(f"  Phase 2 classified: {phase2_hits}")
    print(f"  Still unclassified: {df[df[hs_col].apply(norm) == ''].shape[0]}")

# ── Phase 3: Write results ────────────────────────────────────────────────────
print(f"\nPhase 3: Writing results ...")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="mailed_providers_hs", index=False)
print(f"  Saved: {OUTPUT_FILE.name}")

# Summary
classified   = df[df[hs_col].apply(norm) != ""].shape[0]
unclassified = df[df[hs_col].apply(norm) == ""].shape[0]
print(f"\nSummary:")
print(f"  Total rows       : {len(df)}")
print(f"  Classified       : {classified}")
print(f"  Unclassified     : {unclassified}")

top_systems = df[df[hs_col].apply(norm) != ""][hs_col].value_counts().head(20)
print(f"\n  Top 20 health systems:")
for hs, cnt in top_systems.items():
    print(f"    {cnt:4d}  {hs}")

print("\nDone.")
