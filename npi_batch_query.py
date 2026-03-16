#!/usr/bin/env python3
"""
Batch NPI Registry queries across multiple taxonomy codes.

For each taxonomy code this script:
  1. Searches the NPI registry for providers in STATE with that taxonomy.
  2. Fetches the complete record for every NPI returned.
  3. Appends the results to OUTPUT_FILE (JSON).

If OUTPUT_FILE already exists the script skips taxonomy codes that were
already processed, so interrupted runs can be safely resumed.

Usage:
    python npi_batch_query.py

Configuration:
    Edit STATE, TAXONOMY_CODES, and OUTPUT_FILE below, or override via
    environment variables NPI_STATE and NPI_OUTPUT.
"""

import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
import anthropic

# ── Configuration ──────────────────────────────────────────────────────────────

STATE           = os.getenv("NPI_STATE",        "Minnesota")
COUNTRY         = os.getenv("NPI_COUNTRY",      "United States")
ADDRESS_TYPE    = os.getenv("NPI_ADDRESS_TYPE",  "Primary Location")
NPI_TYPE        = os.getenv("NPI_TYPE",          "Individual")

TAXONOMY_CODES = [
    "207V00000X",   # Obstetrics & Gynecology
    "2080N0001X",   # Neonatal-Perinatal Medicine
    "163WX0002X",   # Obstetric/Gynecologic Registered Nurse
    "163WX0003X",   # Inpatient Obstetric Registered Nurse
    "163WM0102X",   # Maternal Newborn Registered Nurse
    "163WN0002X",   # Neonatal Intensive Care Registered Nurse
    "163WW0101X",   # Wound Care Registered Nurse
    "163WP1700X",   # Perinatal Registered Nurse
    "175M00000X",   # Midwife
    "176B00000X",   # Midwife, Lay
]

OUTPUT_FILE = Path(os.getenv("NPI_OUTPUT", "npi_results.json"))

# ── Constants ──────────────────────────────────────────────────────────────────

NPI_MCP_SERVER = {
    "type": "url",
    "url": "https://mcp.deepsense.ai/npi_registry/mcp",
    "name": "npi-registry",
}

# Ask Claude to return only a JSON array of NPI strings so parsing is reliable
SEARCH_SYSTEM_PROMPT = """\
You are an NPI Registry data extraction assistant.
Use the NPI registry tools to find providers matching the request.
Return ONLY a JSON array of NPI number strings, one per provider found.
Example: ["1234567890", "0987654321"]
Output nothing else — no explanation, no markdown fences.\
"""

# Ask Claude to return the full provider record as a JSON object
DETAIL_SYSTEM_PROMPT = """\
You are an NPI Registry data extraction assistant.
Use the NPI registry tools to retrieve full details for the given NPI number.
Return ONLY a single JSON object with ALL available fields exactly as returned
by the API (npi, enumerationType, basic, addresses, taxonomies, identifiers,
other_names, endpoints, etc.).
Output nothing else — no explanation, no markdown fences.\
"""

MODEL = "claude-sonnet-4-6"

# ── Helpers ────────────────────────────────────────────────────────────────────

def _extract_json(text: str, *, array: bool):
    """Pull the first JSON array or object out of an arbitrary text string."""
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


# ── Core queries ───────────────────────────────────────────────────────────────

def search_npis(
    client: anthropic.Anthropic,
    state: str,
    taxonomy_code: str,
    country: str,
    address_type: str,
    npi_type: str,
) -> list[str]:
    """Return NPI numbers for providers matching all supplied filters."""
    response = client.beta.messages.create(
        model=MODEL,
        max_tokens=2048,
        system=SEARCH_SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": (
                f"Search the NPI registry for healthcare providers with ALL of the following filters:\n"
                f"  - NPI Type: {npi_type}\n"
                f"  - State: {state}\n"
                f"  - Country: {country}\n"
                f"  - Address Type: {address_type}\n"
                f"  - Taxonomy code: {taxonomy_code}\n"
                "Return a JSON array of NPI number strings."
            ),
        }],
        mcp_servers=[NPI_MCP_SERVER],
        betas=["mcp-client-2025-04-04"],
    )

    raw = _extract_json(_text_from_response(response), array=True)

    # Accept both plain NPI strings and objects that contain an "npi" key
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


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    load_dotenv(Path(__file__).parent / ".env")
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY is not set.", file=sys.stderr)
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # Load existing output so interrupted runs can be resumed
    if OUTPUT_FILE.exists():
        with open(OUTPUT_FILE) as f:
            output = json.load(f)
        done = {q["taxonomy_code"] for q in output.get("queries", [])}
        print(f"Resuming — {len(done)} taxonomy code(s) already in {OUTPUT_FILE}.")
    else:
        output = {
            "state": STATE,
            "country": COUNTRY,
            "address_type": ADDRESS_TYPE,
            "npi_type": NPI_TYPE,
            "queries": [],
        }
        done = set()

    for taxonomy_code in TAXONOMY_CODES:
        if taxonomy_code in done:
            print(f"  skip  {taxonomy_code} (already processed)")
            continue

        print(f"\n── Taxonomy {taxonomy_code} ──")
        print(f"  Filters: {NPI_TYPE} | {STATE}, {COUNTRY} | {ADDRESS_TYPE}")
        print("  Searching for NPI numbers...", end=" ", flush=True)
        npis = search_npis(client, STATE, taxonomy_code, COUNTRY, ADDRESS_TYPE, NPI_TYPE)
        print(f"{len(npis)} found.")

        providers: list[dict] = []
        for i, npi in enumerate(npis, 1):
            print(f"  [{i}/{len(npis)}] Fetching full record for NPI {npi}...", end=" ", flush=True)
            details = fetch_npi_details(client, npi)
            providers.append(details)
            print("done.")

        output["queries"].append({
            "taxonomy_code": taxonomy_code,
            "state": STATE,
            "country": COUNTRY,
            "address_type": ADDRESS_TYPE,
            "npi_type": NPI_TYPE,
            "provider_count": len(providers),
            "providers": providers,
        })

        # Write after every taxonomy so progress is never lost on interruption
        with open(OUTPUT_FILE, "w") as f:
            json.dump(output, f, indent=2)
        print(f"  Saved {len(providers)} provider(s) → {OUTPUT_FILE}")

    total = sum(q["provider_count"] for q in output["queries"])
    print(f"\nComplete — {len(output['queries'])} taxonomy code(s), {total} total provider(s).")
    print(f"Output file: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
