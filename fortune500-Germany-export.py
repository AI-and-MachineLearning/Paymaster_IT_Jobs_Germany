#!/usr/bin/env python3
"""
Fortune 500 companies present in Germany -> Excel exporter

Inputs:
  - fortune500.csv with columns: fortune_rank, company_name
Outputs:
  - Excel file with company metadata (Wikidata) + Germany presence heuristics + careers URL guess

Data sources:
  - Wikidata Query Service (SPARQL endpoint) https://query.wikidata.org/sparql
    Docs: https://www.wikidata.org/wiki/Wikidata:SPARQL_query_service
"""

import argparse
import time
import re
from typing import Dict, Any, List, Optional, Tuple

import pandas as pd
import requests

WDQS_URL = "https://query.wikidata.org/sparql"
USER_AGENT = "Fortune500GermanyExporter/1.0 (contact: your-email@example.com)"  # set to your email/org

# Common careers paths to probe (best-effort; many companies redirect to Workday/Greenhouse/etc.)
CAREERS_PATHS = [
    "/careers", "/careers/", "/jobs", "/jobs/", "/career", "/career/",
    "/careers/jobs", "/careers/jobs/", "/careers/search", "/careers/search/",
    "/about/careers", "/about/careers/", "/company/careers", "/company/careers/"
]

def wdqs_query(query: str, timeout: int = 60) -> Dict[str, Any]:
    headers = {
        "Accept": "application/sparql-results+json",
        "User-Agent": USER_AGENT
    }
    r = requests.get(WDQS_URL, params={"query": query, "format": "json"}, headers=headers, timeout=timeout)
    r.raise_for_status()
    return r.json()

def first_binding_value(bindings: List[Dict[str, Any]], key: str) -> Optional[str]:
    if not bindings:
        return None
    if key not in bindings[0]:
        return None
    return bindings[0][key]["value"]

def safe_str(x: Optional[str]) -> str:
    return x if x is not None else ""

def extract_qid(uri: str) -> str:
    # e.g., http://www.wikidata.org/entity/Q2283 -> Q2283
    m = re.search(r"/(Q\d+)$", uri)
    return m.group(1) if m else ""

def entity_search_best_qid(company_name: str, lang: str = "en", limit: int = 5) -> Optional[str]:
    """
    Uses Wikidata API EntitySearch via SPARQL service wikibase:mwapi to find a best matching entity.
    """
    company_escaped = company_name.replace('"', '\\"')
    query = f"""
    SELECT ?item ?itemLabel ?desc WHERE {{
      SERVICE wikibase:mwapi {{
        bd:serviceParam wikibase:api "EntitySearch" ;
                        wikibase:search "{company_escaped}" ;
                        wikibase:language "{lang}" ;
                        wikibase:limit {limit} .
        ?item wikibase:apiOutputItem mwapi:item .
      }}
      OPTIONAL {{ ?item schema:description ?desc FILTER(LANG(?desc)="{lang}") }}
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "{lang}". }}
    }}
    """
    data = wdqs_query(query)
    bindings = data.get("results", {}).get("bindings", [])
    if not bindings:
        return None
    # Heuristic: choose first result (often correct for well-known companies)
    uri = bindings[0]["item"]["value"]
    qid = extract_qid(uri)
    return qid or None

def fetch_company_profile(qid: str, lang: str = "en") -> Dict[str, Any]:
    """
    Fetch core company properties from Wikidata: label, description, website, industry, HQ + HQ country.
    Industries may return multiple values; we take a concatenated list.
    """
    query = f"""
    SELECT ?item ?itemLabel ?desc ?website
           (GROUP_CONCAT(DISTINCT ?industryLabel; separator="; ") AS ?industries)
           ?hqLabel ?hqCountryLabel
    WHERE {{
      BIND(wd:{qid} AS ?item)
      OPTIONAL {{ ?item schema:description ?desc FILTER(LANG(?desc)="{lang}") }}
      OPTIONAL {{ ?item wdt:P856 ?website. }}
      OPTIONAL {{
        ?item wdt:P452 ?industry .
      }}
      OPTIONAL {{
        ?item wdt:P159 ?hq .
        OPTIONAL {{ ?hq wdt:P17 ?hqCountry. }}
      }}
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "{lang}". }}
    }}
    GROUP BY ?item ?itemLabel ?desc ?website ?hqLabel ?hqCountryLabel
    """
    data = wdqs_query(query)
    bindings = data.get("results", {}).get("bindings", [])
    if not bindings:
        return {}

    row = bindings[0]
    return {
        "wikidata_qid": qid,
        "wikidata_url": f"https://www.wikidata.org/wiki/{qid}",
        "description": row.get("desc", {}).get("value", ""),
        "global_website": row.get("website", {}).get("value", ""),
        "industry": row.get("industries", {}).get("value", ""),
        "hq": row.get("hqLabel", {}).get("value", ""),
        "hq_country": row.get("hqCountryLabel", {}).get("value", "")
    }

def detect_germany_presence(qid: str, lang: str = "en", max_entities: int = 20) -> Tuple[bool, str]:
    """
    Heuristic "presence in Germany" detector:
      - Find items that have parent organization (P749) == company
      - AND have country (P17) == Germany (Q183)
    These are often German subsidiaries like "X Deutschland GmbH".
    """
    query = f"""
    SELECT ?deEntity ?deEntityLabel WHERE {{
      ?deEntity wdt:P749 wd:{qid} ;
                wdt:P17 wd:Q183 .
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "{lang}". }}
    }}
    LIMIT {max_entities}
    """
    data = wdqs_query(query)
    bindings = data.get("results", {}).get("bindings", [])
    if not bindings:
        return False, ""

    entities = []
    for b in bindings:
        uri = b["deEntity"]["value"]
        de_qid = extract_qid(uri)
        label = b.get("deEntityLabel", {}).get("value", de_qid)
        entities.append(f"{label} ({de_qid})")

    return True, "; ".join(entities)

def guess_careers_url(base_url: str, timeout: int = 15) -> str:
    """
    Best-effort: probe common careers paths on the official website.
    If base_url is empty or not reachable, returns empty string.
    """
    if not base_url:
        return ""

    # normalize
    base = base_url.strip()
    if base.endswith("/"):
        base = base[:-1]

    headers = {"User-Agent": USER_AGENT}

    def ok_status(code: int) -> bool:
        return 200 <= code < 400

    # First try a HEAD to base (some sites block HEAD; then fallback to GET)
    try:
        r = requests.get(base, headers=headers, timeout=timeout, allow_redirects=True)
        if not ok_status(r.status_code):
            return ""
        final_base = r.url.rstrip("/")
    except Exception:
        return ""

    # Probe known paths
    for path in CAREERS_PATHS:
        url = f"{final_base}{path}"
        try:
            rr = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
            if ok_status(rr.status_code) and rr.url:
                return rr.url
        except Exception:
            continue

    return ""

def process_company(company_name: str, fortune_rank: Any, lang: str = "en", sleep_s: float = 0.2) -> Dict[str, Any]:
    out = {
        "fortune_rank": fortune_rank,
        "company_name": company_name,
        "wikidata_qid": "",
        "wikidata_url": "",
        "description": "",
        "industry": "",
        "hq": "",
        "hq_country": "",
        "global_website": "",
        "careers_website": "",
        "has_germany_presence": False,
        "germany_entities": "",
        "notes": ""
    }

    try:
        qid = entity_search_best_qid(company_name, lang=lang)
        if not qid:
            out["notes"] = "No Wikidata match found via EntitySearch"
            return out

        out["wikidata_qid"] = qid
        out["wikidata_url"] = f"https://www.wikidata.org/wiki/{qid}"

        profile = fetch_company_profile(qid, lang=lang)
        for k in ["description", "industry", "hq", "hq_country", "global_website"]:
            out[k if k != "global_website" else "global_website"] = profile.get(k, "")

        # Germany presence (heuristic)
        has_de, de_entities = detect_germany_presence(qid, lang=lang)
        out["has_germany_presence"] = has_de
        out["germany_entities"] = de_entities

        # Careers site (best-effort)
        out["careers_website"] = guess_careers_url(out["global_website"])

        if not out["industry"]:
            out["notes"] = (out["notes"] + "; " if out["notes"] else "") + "Industry missing in Wikidata"
        if not out["global_website"]:
            out["notes"] = (out["notes"] + "; " if out["notes"] else "") + "Official website missing in Wikidata"
        if not has_de:
            out["notes"] = (out["notes"] + "; " if out["notes"] else "") + "No German subsidiary detected (heuristic)"

        return out

    finally:
        # be polite to public endpoint
        time.sleep(sleep_s)

def main():
    parser = argparse.ArgumentParser(description="Export Fortune 500 companies present in Germany to Excel (via Wikidata).")
    parser.add_argument("--input", required=True, help="Input CSV (fortune_rank, company_name)")
    parser.add_argument("--output", default="fortune500_present_in_germany.xlsx", help="Output Excel file")
    parser.add_argument("--lang", default="en", help="Label language for Wikidata (default: en)")
    parser.add_argument("--limit", type=int, default=0, help="Process only first N rows (0 = all)")
    args = parser.parse_args()

    df_in = pd.read_csv(args.input)
    required_cols = {"company_name"}
    if not required_cols.issubset(df_in.columns):
        raise ValueError(f"Input must contain columns: {sorted(required_cols)}. Optional: fortune_rank")

    if "fortune_rank" not in df_in.columns:
        df_in["fortune_rank"] = ""

    if args.limit and args.limit > 0:
        df_in = df_in.head(args.limit)

    rows: List[Dict[str, Any]] = []
    for _, r in df_in.iterrows():
        company = str(r["company_name"]).strip()
        rank = r.get("fortune_rank", "")
        if not company:
            continue
        rows.append(process_company(company, rank, lang=args.lang))

    output_columns = [
        "fortune_rank", "company_name", "wikidata_qid", "wikidata_url",
        "description", "industry", "hq", "hq_country", "global_website",
        "careers_website", "has_germany_presence", "germany_entities", "notes"
    ]
    df_out = pd.DataFrame(rows, columns=output_columns)

    # Filter to "present in Germany" if you want:
    # df_out = df_out[df_out["has_germany_presence"] == True].copy()

    # Write Excel
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Fortune500_DE")
        # Add a second sheet with just the "present in Germany" subset
        df_out[df_out["has_germany_presence"] == True].to_excel(writer, index=False, sheet_name="Present_in_Germany")

    print(f"Done. Wrote: {args.output}")
    print("Tip: Check 'notes' column and validate edge cases. Wikidata coverage varies by company.")

if __name__ == "__main__":
    main()