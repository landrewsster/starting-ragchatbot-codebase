#!/usr/bin/env python3
"""
Batch NPI Registry queries using the FREE public NPI Registry REST API.
No API key or credits required.

Run from your Mac's Terminal:
    python3 npi_batch_query_free.py

Requirements:
    pip3 install requests   (requests is the only dependency)

Output files:
    npi_results_physicians.csv
    npi_results_nurses_midwives.csv
"""

import csv
import sys
import time
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("Missing dependency — run:  pip3 install requests")

# ── Configuration ──────────────────────────────────────────────────────────────

NPI_API_URL  = "https://npiregistry.cms.hhs.gov/api/"
STATE        = "MN"           # Minnesota
COUNTRY_CODE = "US"           # United States
ENUM_TYPE    = "NPI-1"        # Individual providers only
PAGE_SIZE    = 200            # Max allowed by the API
PAUSE_SEC    = 0.3            # Polite delay between requests

TAXONOMY_SETS = [
    {
        "label":       "physicians",
        "output_file": Path("npi_results_physicians.csv"),
        "taxonomy_codes": [
            "207VX0000X",   # Obstetrics & Gynecology - Complex Family Planning
            "207VG0400X",   # Obstetrics & Gynecology - Gynecology
            "2080N0001X",   # Neonatal-Perinatal Medicine
            "207V00000X",   # Obstetrics & Gynecology
        ],
    },
    {
        "label":       "nurses_and_midwives",
        "output_file": Path("npi_results_nurses_midwives.csv"),
        "taxonomy_codes": [
            "163WX0002X",   # Obstetric/Gynecologic Registered Nurse
            "163WX0003X",   # Inpatient Obstetric Registered Nurse
            "163WM0102X",   # Maternal Newborn Registered Nurse
            "163WN0002X",   # Neonatal Intensive Care Registered Nurse
            "163WW0101X",   # Wound Care Registered Nurse
            "163WP1700X",   # Perinatal Registered Nurse
            "175M00000X",   # Midwife
            "176B00000X",   # Midwife, Lay
        ],
    },
]

# The NPI API's taxonomy_description parameter is a text search on the description,
# not the code. These search terms retrieve the right candidates; results are then
# filtered client-side to the exact taxonomy code.
TAXONOMY_SEARCH_TERMS = {
    "207VX0000X": "Complex Family Planning",
    "207VG0400X": "Gynecology",
    "2080N0001X": "Neonatal-Perinatal",
    "207V00000X": "Obstetrics & Gynecology",
    "163WX0002X": "Obstetric/Gynecologic",
    "163WX0003X": "Inpatient Obstetric",
    "163WM0102X": "Maternal Newborn",
    "163WN0002X": "Neonatal Intensive Care",
    "163WW0101X": "Wound Care",
    "163WP1700X": "Perinatal",
    "175M00000X": "Midwife",
    "176B00000X": "Midwife",   # filtered client-side to exact code
}

# ── CSV columns ────────────────────────────────────────────────────────────────

CSV_COLUMNS = [
    "npi",
    "enumeration_type",
    "status",
    "enumeration_date",
    "last_updated",
    "last_name",
    "first_name",
    "middle_name",
    "name_prefix",
    "name_suffix",
    "credential",
    "gender",
    "sole_proprietor",
    "primary_address_1",
    "primary_address_2",
    "primary_city",
    "primary_state",
    "primary_postal_code",
    "primary_country_code",
    "primary_telephone",
    "primary_fax",
    "mailing_address_1",
    "mailing_address_2",
    "mailing_city",
    "mailing_state",
    "mailing_postal_code",
    "mailing_country_code",
    "mailing_telephone",
    "mailing_fax",
    "taxonomy_codes",           # semicolon-separated
    "taxonomy_descriptions",    # semicolon-separated
    "taxonomy_licenses",        # semicolon-separated
    "taxonomy_states",          # semicolon-separated
    "primary_taxonomy_code",
    "identifiers",              # semicolon-separated "type:value"
    "other_names",              # semicolon-separated
    "query_taxonomy_code",
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _addr(addresses: list, purpose: str) -> dict:
    for a in addresses:
        if isinstance(a, dict) and a.get("address_purpose", "").upper() == purpose:
            return a
    return {}


def flatten_provider(record: dict, query_taxonomy_code: str) -> dict:
    basic      = record.get("basic", {})
    addresses  = record.get("addresses", [])
    taxonomies = record.get("taxonomies", [])
    identifiers = record.get("identifiers", [])
    other_names = record.get("other_names", [])

    loc  = _addr(addresses, "LOCATION")
    mail = _addr(addresses, "MAILING")

    tax_codes  = [t.get("code", "")    for t in taxonomies]
    tax_descs  = [t.get("desc", "")    for t in taxonomies]
    tax_lic    = [t.get("license", "") for t in taxonomies]
    tax_states = [t.get("state", "")   for t in taxonomies]
    primary_tx = next((t.get("code", "") for t in taxonomies if t.get("primary")), "")

    ident_strs = [
        f"{i.get('type', '')}:{i.get('identifier', '')}"
        for i in identifiers if isinstance(i, dict)
    ]

    other_name_strs = [
        " ".join(filter(None, [
            n.get("prefix", ""), n.get("first_name", ""),
            n.get("middle_name", ""), n.get("last_name", ""),
            n.get("suffix", ""),
        ]))
        for n in other_names if isinstance(n, dict)
    ]

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
        "gender":               basic.get("gender", ""),
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
        "taxonomy_codes":           "; ".join(tax_codes),
        "taxonomy_descriptions":    "; ".join(tax_descs),
        "taxonomy_licenses":        "; ".join(tax_lic),
        "taxonomy_states":          "; ".join(tax_states),
        "primary_taxonomy_code":    primary_tx,
        "identifiers":              "; ".join(ident_strs),
        "other_names":              "; ".join(other_name_strs),
        "query_taxonomy_code":      query_taxonomy_code,
    }

# ── NPI API pagination ─────────────────────────────────────────────────────────

def fetch_all_for_taxonomy(taxonomy_code: str) -> list[dict]:
    """
    Page through the NPI registry and return every provider record whose
    taxonomies list contains *taxonomy_code* exactly.

    The API's taxonomy_description param does text matching on descriptions,
    not codes, so we use a mapped search term to retrieve candidates and then
    filter client-side by the exact code.
    """
    search_term = TAXONOMY_SEARCH_TERMS.get(taxonomy_code, taxonomy_code)
    matched, skip = [], 0

    while True:
        params = {
            "version":              "2.1",
            "enumeration_type":     ENUM_TYPE,
            "taxonomy_description": search_term,
            "state":                STATE,
            "country_code":         COUNTRY_CODE,
            "limit":                PAGE_SIZE,
            "skip":                 skip,
        }

        resp = requests.get(NPI_API_URL, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        results      = data.get("results", [])
        result_count = data.get("result_count", 0)

        # Filter to providers that carry the exact taxonomy code
        page_matched = [
            r for r in results
            if any(t.get("code") == taxonomy_code for t in r.get("taxonomies", []))
        ]
        matched.extend(page_matched)
        skip += len(results)

        print(f"    fetched {skip}/{result_count} candidates — "
              f"{len(matched)} match code {taxonomy_code} so far")

        if not results or skip >= result_count:
            break

        time.sleep(PAUSE_SEC)

    return matched

# ── Per-set runner ─────────────────────────────────────────────────────────────

def run_taxonomy_set(taxonomy_set: dict) -> None:
    label       = taxonomy_set["label"]
    output_file = taxonomy_set["output_file"]
    tax_codes   = taxonomy_set["taxonomy_codes"]

    print(f"\n{'═' * 60}")
    print(f"  Set : {label}")
    print(f"  File: {output_file}")
    print(f"  Filters: {ENUM_TYPE} | {STATE} | {COUNTRY_CODE}")
    print(f"  Taxonomy codes: {', '.join(tax_codes)}")
    print(f"{'═' * 60}")

    # Detect already-processed taxonomy codes for resuming
    done: set[str] = set()
    file_exists = output_file.exists()
    if file_exists:
        with open(output_file, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("query_taxonomy_code"):
                    done.add(row["query_taxonomy_code"])
        print(f"  Resuming — {len(done)} taxonomy code(s) already in file.")

    with open(output_file, "a", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        if not file_exists:
            writer.writeheader()

        for taxonomy_code in tax_codes:
            if taxonomy_code in done:
                print(f"\n  skip  {taxonomy_code} (already in CSV)")
                continue

            print(f"\n  ── {taxonomy_code} ──")
            providers = fetch_all_for_taxonomy(taxonomy_code)
            print(f"  {len(providers)} provider(s) found.")

            for provider in providers:
                writer.writerow(flatten_provider(provider, taxonomy_code))
            csvfile.flush()
            print(f"  Written to {output_file}")

    total = sum(1 for _ in open(output_file, encoding="utf-8")) - 1
    print(f"\n  {label} complete — {total} row(s) in {output_file.resolve()}")

# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    for taxonomy_set in TAXONOMY_SETS:
        run_taxonomy_set(taxonomy_set)

    print(f"\n{'═' * 60}")
    print("  All sets complete.")
    for ts in TAXONOMY_SETS:
        print(f"    {ts['label']:30s} → {ts['output_file'].resolve()}")
    print(f"{'═' * 60}")


if __name__ == "__main__":
    main()
