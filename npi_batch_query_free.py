#!/usr/bin/env python3
"""
Batch NPI Registry queries using the FREE public NPI Registry REST API.
No API key or credits required.
Run from your Mac's Terminal:
    python3 npi_batch_query_free.py
Requirements:
    pip3 install requests
Output files:
    npi_results_physicians.csv
    npi_results_nurses.csv
    npi_results_physician_assistants.csv
"""
import csv
import sys
import time
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("Missing dependency — run:  pip3 install requests")

# ── Nominatim organisation lookup ──────────────────────────────────────────────

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
_org_cache: dict = {}

def lookup_organization(address_1: str, city: str, state: str) -> str:
    """Return a place/building name for the given address via OSM Nominatim.
    Returns empty string if not found or on error.
    Nominatim policy: max 1 request/second, requires User-Agent header.
    """
    if not address_1 or not city:
        return ""
    key = (address_1.strip().lower(), city.strip().lower(), state.strip().lower())
    if key in _org_cache:
        return _org_cache[key]
    params = {
        "q":            f"{address_1}, {city}, {state}",
        "format":       "json",
        "limit":        1,
        "addressdetails": 0,
    }
    headers = {"User-Agent": "npi-batch-query/1.0 (research use)"}
    try:
        resp = requests.get(NOMINATIM_URL, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
        hits = resp.json()
        name = hits[0].get("name", "") if hits else ""
    except Exception:
        name = ""
    _org_cache[key] = name
    time.sleep(1.1)   # Nominatim hard rate limit: 1 req/sec
    return name

# ── Configuration ──────────────────────────────────────────────────────────────

NPI_API_URL       = "https://npiregistry.cms.hhs.gov/api/"
STATE             = "MN"
COUNTRY_CODE      = "US"
ENUM_TYPE         = "NPI-1"
PAGE_SIZE         = 200
PAUSE_SEC         = 0.3
LOOKUP_ORG        = False   # Set True to enable Nominatim org lookup (adds ~1s per provider)

TAXONOMY_SETS = [
    {
        "label":       "physicians",
        "output_file": Path("npi_results_physicians.csv"),
        "taxonomy_codes": [
            "207VX0000X",
            "207VG0400X",
            "2080N0001X",
            "207V00000X",
            "207VM0101X",
            "207VC0300X",
            "207VC0200X",
            "207VE0102X",
            "207RA0401X",
            "207QA0401X",
            "208D00000X",
            "2083A0300X",
        ],
    },
    {
        "label":       "nurses",
        "output_file": Path("npi_results_nurses.csv"),
        "taxonomy_codes": [
            "163WX0002X",
            "163WX0003X",
            "163WM0102X",
            "163WN0002X",
            "163WW0101X",
            "163WP1700X",
            "367A00000X",
            "363LX0001X",
            "363LW0102X",
            "163WN0003X",
            "163WA0400X",
            "163WC1500X",
            "163WL0100X",
            "163WR1000X",
            "364SC1501X",
            "364SF0001X",
            "364SP1700X",
            "364SW0102X",
            "363LC1500X",
            "363LF0000X",
            "363LP1700X",
            "363LP2300X",
            "164W00000X",
            "164X00000X",
        ],
    },
    {
        "label":       "physician_assistants",
        "output_file": Path("npi_results_physician_assistants.csv"),
        "taxonomy_codes": [
            "363A00000X",
            "363AM0700X",
            "363AS0400X",
        ],
    },
]

TAXONOMY_SEARCH_TERMS = {
    "207VX0000X": "Obstetrics",
    "207VG0400X": "Gynecology",
    "2080N0001X": "Neonatal-Perinatal Medicine",
    "207V00000X": "Obstetrics & Gynecology",
    "207VM0101X": "Maternal & Fetal Medicine",
    "207VC0300X": "Complex Family Planning",
    "207VC0200X": "Critical Care Medicine",
    "207VE0102X": "Reproductive Endocrinology",
    "207RA0401X": "Addiction Medicine",
    "207QA0401X": "Addiction Medicine",
    "208D00000X": "General Practice",
    "2083A0300X": "Addiction Medicine",
    "163WX0002X": "Obstetric, High-Risk",
    "163WX0003X": "Obstetric, Inpatient",
    "163WM0102X": "Maternal Newborn",
    "163WN0002X": "Neonatal Intensive Care",
    "163WN0003X": "Neonatal, Low-Risk",
    "163WW0101X": "Women's Health Care, Ambulatory",
    "163WP1700X": "Perinatal",
    "367A00000X": "Advanced Practice Midwife",  # also known as CNM
    "363LX0001X": "Obstetrics & Gynecology",
    "363LW0102X": "Women's Health",
    "163WA0400X": "Addiction",
    "163WC1500X": "Community Health",
    "163WL0100X": "Lactation Consultant",
    "163WR1000X": "Reproductive Endocrinology",
    "364SC1501X": "Community Health",
    "364SF0001X": "Family Health",
    "364SP1700X": "Perinatal",
    "364SW0102X": "Women's Health",
    "363LC1500X": "Community Health",
    "363LF0000X": "Family",
    "363LP1700X": "Perinatal",
    "363LP2300X": "Primary Care",
    "164W00000X": "Licensed Practical Nurse",
    "164X00000X": "Licensed Vocational Nurse",
    "363A00000X": "Physician Assistant",
    "363AM0700X": "Medical",
    "363AS0400X": "Surgical",
}

CSV_COLUMNS = [
    "npi", "enumeration_type", "status", "enumeration_date", "last_updated",
    "last_name", "first_name", "middle_name", "name_prefix", "name_suffix",
    "credential", "authorized_official_name", "authorized_official_title", "authorized_official_phone", "sole_proprietor",
    "primary_address_1", "primary_address_2", "primary_city", "primary_state",
    "primary_postal_code", "primary_country_code", "primary_telephone", "primary_fax",
    "mailing_address_1", "mailing_address_2", "mailing_city", "mailing_state",
    "mailing_postal_code", "mailing_country_code", "mailing_telephone", "mailing_fax",
    "taxonomy_codes", "taxonomy_descriptions", "taxonomy_licenses", "taxonomy_states",
    "primary_taxonomy_code", "identifiers", "other_names", "query_taxonomy_code",
    "organization_lookup",
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _addr(addresses, purpose):
    for a in addresses:
        if isinstance(a, dict) and a.get("address_purpose", "").upper() == purpose:
            return a
    return {}

def flatten_provider(record, query_taxonomy_code):
    basic       = record.get("basic", {})
    addresses   = record.get("addresses", [])
    taxonomies  = record.get("taxonomies", [])
    identifiers = record.get("identifiers", [])
    other_names = record.get("other_names", [])
    loc  = _addr(addresses, "LOCATION")
    mail = _addr(addresses, "MAILING")
    return {
        "npi":                  record.get("number", ""),
        "enumeration_type":     record.get("enumeration_type", ""),
        "status":               basic.get("status", ""),
        "enumeration_date":     basic.get("enumeration_date", ""),
        "last_updated":         basic.get("last_updated", ""),
        "last_name":            basic.get("last_name", ""),
        "first_name":           basic.get("first_name", ""),
        "middle_name":          basic.get("middle_name", ""),
        "name_prefix":          basic.get("name_prefix", ""),
        "name_suffix":          basic.get("name_suffix", ""),
        "credential":           basic.get("credential", ""),
        "authorized_official_name": " ".join(filter(None, [
            basic.get("authorized_official_first_name", ""),
            basic.get("authorized_official_last_name", ""),
        ])),
        "authorized_official_title": basic.get("authorized_official_title_or_position", ""),
        "authorized_official_phone": basic.get("authorized_official_telephone_number", ""),
        "sole_proprietor":      basic.get("sole_proprietor", ""),
        "primary_address_1":    loc.get("address_1", ""),
        "primary_address_2":    loc.get("address_2", ""),
        "primary_city":         loc.get("city", ""),
        "primary_state":        loc.get("state", ""),
        "primary_postal_code":  loc.get("postal_code", ""),
        "primary_country_code": loc.get("country_code", ""),
        "primary_telephone":    loc.get("telephone_number", ""),
        "primary_fax":          loc.get("fax_number", ""),
        "mailing_address_1":    mail.get("address_1", ""),
        "mailing_address_2":    mail.get("address_2", ""),
        "mailing_city":         mail.get("city", ""),
        "mailing_state":        mail.get("state", ""),
        "mailing_postal_code":  mail.get("postal_code", ""),
        "mailing_country_code": mail.get("country_code", ""),
        "mailing_telephone":    mail.get("telephone_number", ""),
        "mailing_fax":          mail.get("fax_number", ""),
        "taxonomy_codes":        "; ".join(t.get("code", "")    or "" for t in taxonomies),
        "taxonomy_descriptions": "; ".join(t.get("desc", "")    or "" for t in taxonomies),
        "taxonomy_licenses":     "; ".join(t.get("license", "") or "" for t in taxonomies),
        "taxonomy_states":       "; ".join(t.get("state", "")   or "" for t in taxonomies),
        "primary_taxonomy_code": next((t.get("code", "") for t in taxonomies if t.get("primary")), ""),
        "identifiers":           "; ".join(f"{i.get('type','')}:{i.get('identifier','')}" for i in identifiers if isinstance(i, dict)),
        "other_names":           "; ".join(" ".join(filter(None, [n.get("prefix",""), n.get("first_name",""), n.get("middle_name",""), n.get("last_name",""), n.get("suffix","")])) for n in other_names if isinstance(n, dict)),
        "query_taxonomy_code":   query_taxonomy_code,
        "organization_lookup":   lookup_organization(
                                     loc.get("address_1", ""),
                                     loc.get("city", ""),
                                     loc.get("state", ""),
                                 ) if LOOKUP_ORG else "",
    }

# ── NPI API fetch ──────────────────────────────────────────────────────────────

def fetch_all_for_taxonomy(taxonomy_code):
    search_term = TAXONOMY_SEARCH_TERMS.get(taxonomy_code, taxonomy_code)
    matched, skip = [], 0
    while True:
        params = {
            "version": "2.1", "enumeration_type": ENUM_TYPE,
            "taxonomy_description": search_term,
            "state": STATE, "country_code": COUNTRY_CODE,
            "limit": PAGE_SIZE, "skip": skip,
        }
        resp = requests.get(NPI_API_URL, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        results      = data.get("results", [])
        result_count = data.get("result_count", 0)
        page_matched = [
            r for r in results
            if any(t.get("code") == taxonomy_code for t in r.get("taxonomies", []))
        ]
        matched.extend(page_matched)
        skip += len(results)
        print(f"    fetched {skip}/{result_count} candidates — {len(matched)} match {taxonomy_code} so far")
        if not results or skip >= result_count:
            break
        time.sleep(PAUSE_SEC)
    return matched

# ── Per-set runner ─────────────────────────────────────────────────────────────

def run_taxonomy_set(taxonomy_set):
    label, output_file, tax_codes = taxonomy_set["label"], taxonomy_set["output_file"], taxonomy_set["taxonomy_codes"]
    print(f"\n{'═'*60}\n  Set : {label}\n  File: {output_file}\n{'═'*60}")

    done = set()
    # Check whether the existing file (if any) has the current header.
    # If the header differs (e.g. columns were added/removed), treat it as a
    # new file so it gets overwritten with the correct header.
    write_mode = "w"
    if output_file.exists():
        with open(output_file, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            if list(reader.fieldnames or []) == CSV_COLUMNS:
                # Header matches — safe to resume
                for row in reader:
                    if row.get("query_taxonomy_code"):
                        done.add(row["query_taxonomy_code"])
                write_mode = "a"
                print(f"  Resuming — {len(done)} taxonomy code(s) already done.")
            else:
                print(f"  Header mismatch — overwriting {output_file} with current columns.")

    with open(output_file, write_mode, newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        if write_mode == "w":
            writer.writeheader()
        for taxonomy_code in tax_codes:
            if taxonomy_code in done:
                print(f"\n  skip  {taxonomy_code}")
                continue
            print(f"\n  ── {taxonomy_code} ──")
            providers = fetch_all_for_taxonomy(taxonomy_code)
            print(f"  {len(providers)} provider(s) found.")
            for p in providers:
                writer.writerow(flatten_provider(p, taxonomy_code))
            csvfile.flush()
            print(f"  Saved → {output_file}")

    total = sum(1 for _ in open(output_file, encoding="utf-8")) - 1
    print(f"\n  {label} complete — {total} row(s) in {output_file.resolve()}")

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    for taxonomy_set in TAXONOMY_SETS:
        run_taxonomy_set(taxonomy_set)
    print(f"\n{'═'*60}\n  All sets complete.")
    for ts in TAXONOMY_SETS:
        print(f"    {ts['label']:30s} → {ts['output_file'].resolve()}")
    print(f"{'═'*60}")

if __name__ == "__main__":
    main()
