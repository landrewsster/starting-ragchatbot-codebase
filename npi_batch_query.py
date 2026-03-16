#!/usr/bin/env python3
"""
Batch NPI Registry queries across two named taxonomy sets.

For each taxonomy code this script:
  1. Searches the NPI registry for Individual providers in Minnesota (Primary
     Location, United States) with that taxonomy.
  2. Fetches the complete record for every NPI returned.
  3. Writes one CSV row per provider to the set's output file.

Interrupted runs can be safely resumed — already-processed taxonomy codes
are detected from the existing CSV and skipped.

Usage:
    python npi_batch_query.py
"""

import csv
import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
import anthropic

# ── Shared filters ─────────────────────────────────────────────────────────────

STATE        = "Minnesota"
COUNTRY      = "United States"
ADDRESS_TYPE = "Primary Location"
NPI_TYPE     = "Individual"

# ── Taxonomy sets ──────────────────────────────────────────────────────────────

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

# ── Constants ──────────────────────────────────────────────────────────────────

NPI_MCP_SERVER = {
    "type": "url",
    "url": "https://mcp.deepsense.ai/npi_registry/mcp",
    "name": "npi-registry",
}

SEARCH_SYSTEM_PROMPT = """\
You are an NPI Registry data extraction assistant.
Use the NPI registry tools to find providers matching the request.
Return ONLY a JSON array of NPI number strings, one per provider found.
Example: ["1234567890", "0987654321"]
Output nothing else — no explanation, no markdown fences.\
"""

DETAIL_SYSTEM_PROMPT = """\
You are an NPI Registry data extraction assistant.
Use the NPI registry tools to retrieve full details for the given NPI number.
Return ONLY a single JSON object with ALL available fields exactly as returned
by the API (npi, enumerationType, basic, addresses, taxonomies, identifiers,
other_names, endpoints, etc.).
Output nothing else — no explanation, no markdown fences.\
"""

MODEL = "claude-sonnet-4-6"

# CSV columns — every provider row will contain these fields
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
    "taxonomy_codes",           # semicolon-separated list
    "taxonomy_descriptions",    # semicolon-separated list
    "taxonomy_licenses",        # semicolon-separated list
    "taxonomy_states",          # semicolon-separated list
    "primary_taxonomy_code",
    "identifiers",              # semicolon-separated "type:value" pairs
    "other_names",              # semicolon-separated
    "query_taxonomy_code",      # the taxonomy code used to find this provider
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _extract_json(text: str, *, array: bool):
    open_ch, close_ch = ('[', ']') if array else ('{', '}')
    start = text.find(open_ch)
    end   = text.rfind(close_ch) + 1
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end])
        except json.JSONDecodeError:
            pass
    return [] if array else {}


def _text_from_response(response) -> str:
    return next(
        (block.text for block in response.content if hasattr(block, "text")),
        "",
    )


def _addr(addresses: list, purpose: str) -> dict:
    """Return the first address block matching *purpose* ('LOCATION' or 'MAILING')."""
    for a in addresses:
        if isinstance(a, dict) and a.get("address_purpose", "").upper() == purpose:
            return a
    return {}


def flatten_provider(record: dict, query_taxonomy_code: str) -> dict:
    """Convert a raw NPI API record dict into a flat CSV row dict."""
    basic     = record.get("basic", {})
    addresses = record.get("addresses", [])
    taxonomies = record.get("taxonomies", [])
    identifiers = record.get("identifiers", [])
    other_names = record.get("other_names", [])

    loc  = _addr(addresses, "LOCATION")
    mail = _addr(addresses, "MAILING")

    tax_codes  = [t.get("code", "")        for t in taxonomies]
    tax_descs  = [t.get("desc", "")        for t in taxonomies]
    tax_lic    = [t.get("license", "")     for t in taxonomies]
    tax_states = [t.get("state", "")       for t in taxonomies]
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
        "npi":                  record.get("npi", ""),
        "enumeration_type":     record.get("enumerationType", ""),
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


# ── Core queries ───────────────────────────────────────────────────────────────

def search_npis(client: anthropic.Anthropic, taxonomy_code: str) -> list[str]:
    """Return NPI numbers for Individual providers in Minnesota matching *taxonomy_code*."""
    response = client.beta.messages.create(
        model=MODEL,
        max_tokens=2048,
        system=SEARCH_SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": (
                f"Search the NPI registry for healthcare providers with ALL of the following filters:\n"
                f"  - NPI Type: {NPI_TYPE}\n"
                f"  - State: {STATE}\n"
                f"  - Country: {COUNTRY}\n"
                f"  - Address Type: {ADDRESS_TYPE}\n"
                f"  - Taxonomy code: {taxonomy_code}\n"
                "Return a JSON array of NPI number strings."
            ),
        }],
        mcp_servers=[NPI_MCP_SERVER],
        betas=["mcp-client-2025-04-04"],
    )

    raw = _extract_json(_text_from_response(response), array=True)
    npis: list[str] = []
    for item in raw:
        if isinstance(item, str):
            npis.append(item)
        elif isinstance(item, dict) and "npi" in item:
            npis.append(str(item["npi"]))
    return npis


def fetch_npi_details(client: anthropic.Anthropic, npi: str) -> dict:
    """Return the complete NPI record for a single NPI number."""
    response = client.beta.messages.create(
        model=MODEL,
        max_tokens=2048,
        system=DETAIL_SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Retrieve all available information for NPI number {npi}.",
        }],
        mcp_servers=[NPI_MCP_SERVER],
        betas=["mcp-client-2025-04-04"],
    )

    result = _extract_json(_text_from_response(response), array=False)
    if not result:
        result = {"npi": npi, "error": "details not retrieved"}
    return result


# ── Per-set runner ─────────────────────────────────────────────────────────────

def run_taxonomy_set(client: anthropic.Anthropic, taxonomy_set: dict) -> None:
    label        = taxonomy_set["label"]
    output_file  = taxonomy_set["output_file"]
    tax_codes    = taxonomy_set["taxonomy_codes"]

    print(f"\n{'═' * 60}")
    print(f"  Set: {label}  →  {output_file}")
    print(f"  Filters: {NPI_TYPE} | {STATE}, {COUNTRY} | {ADDRESS_TYPE}")
    print(f"  Taxonomy codes: {', '.join(tax_codes)}")
    print(f"{'═' * 60}")

    # Detect already-processed taxonomy codes so the run can be resumed
    done: set[str] = set()
    file_exists = output_file.exists()
    if file_exists:
        with open(output_file, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("query_taxonomy_code"):
                    done.add(row["query_taxonomy_code"])
        print(f"  Resuming — {len(done)} taxonomy code(s) already in {output_file}.")

    # Open CSV for appending (write header only if file is new)
    with open(output_file, "a", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        if not file_exists:
            writer.writeheader()

        for taxonomy_code in tax_codes:
            if taxonomy_code in done:
                print(f"\n  skip  {taxonomy_code} (already in CSV)")
                continue

            print(f"\n  ── {taxonomy_code} ──")
            print("    Searching for NPI numbers...", end=" ", flush=True)
            npis = search_npis(client, taxonomy_code)
            print(f"{len(npis)} found.")

            for i, npi in enumerate(npis, 1):
                print(f"    [{i}/{len(npis)}] NPI {npi}...", end=" ", flush=True)
                details  = fetch_npi_details(client, npi)
                flat_row = flatten_provider(details, taxonomy_code)
                writer.writerow(flat_row)
                csvfile.flush()   # persist row immediately
                print("saved.")

    total_rows = sum(1 for _ in open(output_file, encoding="utf-8")) - 1  # minus header
    print(f"\n  {label} complete — {total_rows} provider row(s) in {output_file.resolve()}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    load_dotenv(Path(__file__).parent / ".env")
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY is not set.", file=sys.stderr)
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    for taxonomy_set in TAXONOMY_SETS:
        run_taxonomy_set(client, taxonomy_set)

    print(f"\n{'═' * 60}")
    print("  All sets complete.")
    for ts in TAXONOMY_SETS:
        print(f"    {ts['label']:25s} → {ts['output_file'].resolve()}")
    print(f"{'═' * 60}")


if __name__ == "__main__":
    main()
