"""
paymaster_jobs_germany_it_export.py

Client-side script to collect IT/Technology/Software job openings from selected
global paymaster/EOR vendors that expose public job feeds (Greenhouse, Lever, Workable),
filter for Germany (primary: Baden-Württemberg), and export to Excel.

Sources used:
- Greenhouse Job Board API (public GET) https://boards-api.greenhouse.io/v1/boards/{board_token}/jobs?content=true
- Lever Postings API (public)          https://api.lever.co/v0/postings/{site}?mode=json
- Workable Widget JSON (public)        https://apply.workable.com/api/v1/widget/accounts/{clientname}

Requirements:
  pip install requests pandas openpyxl
Python: 3.9+ recommended
"""

from __future__ import annotations

import re
import time
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

import requests
import pandas as pd


# =========================
# CONFIGURATION
# =========================

OUTPUT_XLSX = f"paymaster_IT_jobs_Germany_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# Germany focus
GERMANY_ALIASES = [
    "germany", "deutschland", "de", "remote - germany", "remote (germany)"
]

# Baden-Württemberg focus (cities + state name variants)
BW_KEYWORDS = [
    "baden-württemberg", "baden württemberg", "baden wuerttemberg", "bw",
    "stuttgart", "karlsruhe", "mannheim", "heidelberg", "freiburg", "ulm",
    "heilbronn", "reutlingen", "tübingen", "tuebingen", "pforzheim",
    "konstanz", "offenburg", "esslingen", "böblingen", "boeblingen",
    "ludwigsburg", "sinsheim", "ravensburg"
]

# Tech/IT filtering
TECH_INCLUDE_REGEX = re.compile(
    r"(?i)\b("
    r"software|engineer|engineering|developer|development|devops|sre|platform|"
    r"data|analytics|machine learning|ml|ai|artificial intelligence|"
    r"security|infosec|cyber|cloud|infrastructure|it\b|sysadmin|"
    r"qa\b|quality|test|automation|architect|site reliability|"
    r"backend|front ?end|full ?stack|mobile|ios|android|"
    r"database|dba|etl|bi\b|observability|kubernetes|"
    r")\b"
)

# Exclusions to avoid non-tech divisions (tune as you like)
TECH_EXCLUDE_REGEX = re.compile(
    r"(?i)\b("
    r"sales|account executive|business development|bd\b|marketing|"
    r"hr\b|people\b|recruit|talent acquisition|customer success|"
    r"payroll specialist\b|legal|finance|accounting|"
    r")\b"
)

# IMPORTANT:
# - We still allow “Payroll Engineer”, “Security Engineer”, “Data Engineer” etc.
# - Exclusion includes "payroll specialist" specifically, not "payroll engineer".


# =========================
# COMPANY CONNECTORS
# =========================

@dataclass
class SourceConfig:
    company: str
    source_type: str  # "greenhouse" | "lever" | "workable" | "manual"
    token_or_site: Optional[str] = None
    url: Optional[str] = None
    notes: Optional[str] = None


SOURCES: List[SourceConfig] = [
    # Greenhouse boards (public job posts)
    SourceConfig(
        company="Omnipresent",
        source_type="greenhouse",
        token_or_site="omnipresent",
        url="https://boards.greenhouse.io/omnipresent",
        notes="Greenhouse board token: omnipresent"
    ),
    SourceConfig(
        company="Remote",
        source_type="greenhouse",
        token_or_site="remotereferralboardinternaluseonly",
        url="https://job-boards.greenhouse.io/remotereferralboardinternaluseonly",
        notes="Greenhouse board visible publicly; filter for Germany at runtime"
    ),

    # Lever postings (public)
    SourceConfig(
        company="RemoFirst",
        source_type="lever",
        token_or_site="remofirst",
        url="https://jobs.lever.co/remofirst",
        notes="Lever site: remofirst"
    ),

    # Workable widget (public)
    SourceConfig(
        company="WorkMotion",
        source_type="workable",
        token_or_site="workmotion",
        url="https://apply.workable.com/workmotion",
        notes="Workable account: workmotion"
    ),

    # Manual registry (kept in Excel but not fetched automatically)
    SourceConfig(company="Deel", source_type="manual", url="https://www.deel.com/careers/", notes="Portal is dynamic; add API later if identified."),
    SourceConfig(company="Safeguard Global", source_type="manual", url="https://safeguardglobal.wd3.myworkdayjobs.com/External_Careers/", notes="Workday portal; scraping is unstable. Use manual search/export."),
    SourceConfig(company="CloudPay", source_type="manual", url="https://cloudpay.wd3.myworkdayjobs.com/CloudPay_External", notes="Workday portal; scraping is unstable. Use manual search/export."),
    SourceConfig(company="TMF Group", source_type="manual", url="https://www.tmf-group.com/en/careers/", notes="Uses PageUp Careers; better tracked manually."),
    SourceConfig(company="SD Worx", source_type="manual", url="https://careers.sdworx.com/jobs", notes="Teamtailor portal; API needs key. Track manually unless you have a key."),
    SourceConfig(company="Multiplier", source_type="manual", url="https://www.usemultiplier.com/careers", notes="Kula ATS; add API later if identified."),
    SourceConfig(company="Rippling", source_type="manual", url="https://www.rippling.com/careers", notes="Portal varies; add API later if identified."),
    SourceConfig(company="Velocity Global / Pebl", source_type="manual", url="https://hellopebl.com/company/careers/", notes="Portal varies; add API later if identified."),
    SourceConfig(company="Lano", source_type="manual", url="https://www.lano.io/careers", notes="Portal varies; add API later if identified."),
]


# =========================
# HTTP UTILITIES
# =========================

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; PaymasterJobCollector/1.0; +https://example.local)"
})

def http_get_json(url: str, params: Optional[Dict[str, Any]] = None, timeout: int = 30) -> Any:
    resp = SESSION.get(url, params=params, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


# =========================
# NORMALIZATION HELPERS
# =========================

def norm_text(s: Optional[str]) -> str:
    return (s or "").strip()

def classify_germany(location_blob: str) -> bool:
    t = location_blob.lower()
    return any(alias in t for alias in GERMANY_ALIASES) or " germany" in t or t.endswith(", de")

def classify_bw(location_blob: str) -> bool:
    t = location_blob.lower()
    return any(k in t for k in BW_KEYWORDS)

def classify_region(location_blob: str) -> str:
    if not classify_germany(location_blob):
        return "Non-Germany"
    if classify_bw(location_blob):
        return "Baden-Württemberg"
    if "remote" in location_blob.lower() and "germany" in location_blob.lower():
        return "Germany (Remote/Unspecified)"
    return "Germany (Other State/Unspecified)"

def is_tech_role(title: str, dept: str = "", team: str = "", description: str = "") -> bool:
    blob = " ".join([title, dept, team, description])
    if TECH_INCLUDE_REGEX.search(blob) is None:
        return False
    # Allow strongly tech titles even if exclusion matches
    if re.search(r"(?i)\b(payroll engineer|software engineer|data engineer|security engineer|platform engineer)\b", blob):
        return True
    if TECH_EXCLUDE_REGEX.search(blob):
        return False
    return True


# =========================
# FETCHERS
# =========================

def fetch_greenhouse(board_token: str, company: str) -> List[Dict[str, Any]]:
    url = f"https://boards-api.greenhouse.io/v1/boards/{board_token}/jobs"
    data = http_get_json(url, params={"content": "true"})
    jobs = data.get("jobs", []) if isinstance(data, dict) else []
    out = []
    for j in jobs:
        title = norm_text(j.get("title"))
        loc = norm_text((j.get("location") or {}).get("name"))
        job_url = norm_text(j.get("absolute_url"))
        desc = norm_text(j.get("content"))
        # departments/offices are lists in content=true responses
        depts = j.get("departments") or []
        dept_names = ", ".join([d.get("name", "") for d in depts if isinstance(d, dict) and d.get("name")])[:500]
        offices = j.get("offices") or []
        office_names = ", ".join([o.get("name", "") for o in offices if isinstance(o, dict) and o.get("name")])[:500]
        updated = norm_text(j.get("updated_at"))

        out.append({
            "company": company,
            "source_system": "greenhouse",
            "title": title,
            "department": dept_names,
            "team": "",
            "location": loc,
            "office": office_names,
            "remote_type": "",
            "country_code": "",
            "apply_url": job_url,
            "job_url": job_url,
            "updated_at": updated,
            "description_snippet": desc[:1000],
        })
    return out


def fetch_lever(site: str, company: str) -> List[Dict[str, Any]]:
    base = f"https://api.lever.co/v0/postings/{site}"
    out = []
    skip = 0
    limit = 100

    while True:
        data = http_get_json(base, params={"mode": "json", "skip": skip, "limit": limit})
        if not isinstance(data, list) or len(data) == 0:
            break

        for j in data:
            title = norm_text(j.get("text"))
            cats = j.get("categories") or {}
            loc = norm_text(cats.get("location"))
            team = norm_text(cats.get("team"))
            dept = norm_text(cats.get("department"))
            country = norm_text(j.get("country"))
            job_url = norm_text(j.get("hostedUrl"))
            apply_url = norm_text(j.get("applyUrl"))
            remote_type = norm_text(j.get("workplaceType"))
            desc = norm_text(j.get("descriptionPlain"))

            out.append({
                "company": company,
                "source_system": "lever",
                "title": title,
                "department": dept,
                "team": team,
                "location": loc,
                "office": "",
                "remote_type": remote_type,
                "country_code": country,
                "apply_url": apply_url or job_url,
                "job_url": job_url,
                "updated_at": "",
                "description_snippet": desc[:1000],
            })

        if len(data) < limit:
            break
        skip += limit
        time.sleep(0.25)

    return out


def fetch_workable_widget(account: str, company: str) -> List[Dict[str, Any]]:
    url = f"https://apply.workable.com/api/v1/widget/accounts/{account}"
    data = http_get_json(url)
    jobs = data.get("jobs", []) if isinstance(data, dict) else []
    out = []

    for j in jobs:
        title = norm_text(j.get("title"))
        loc = norm_text(j.get("location"))
        dept = norm_text(j.get("department"))
        job_url = norm_text(j.get("url"))
        # Depending on Workable config, "application_url" may exist
        apply_url = norm_text(j.get("application_url")) or job_url
        remote_type = "remote" if "remote" in (loc.lower()) else ""
        desc = norm_text(j.get("short_description")) or norm_text(j.get("description"))

        out.append({
            "company": company,
            "source_system": "workable_widget",
            "title": title,
            "department": dept,
            "team": "",
            "location": loc,
            "office": "",
            "remote_type": remote_type,
            "country_code": "",
            "apply_url": apply_url,
            "job_url": job_url,
            "updated_at": "",
            "description_snippet": desc[:1000],
        })

    return out


# =========================
# MAIN EXECUTION
# =========================

def collect_all_jobs() -> Tuple[pd.DataFrame, pd.DataFrame]:
    all_rows: List[Dict[str, Any]] = []
    manual_rows: List[Dict[str, Any]] = []

    for src in SOURCES:
        if src.source_type == "manual":
            manual_rows.append({
                "company": src.company,
                "career_portal_url": src.url,
                "notes": src.notes
            })
            continue

        try:
            if src.source_type == "greenhouse":
                all_rows.extend(fetch_greenhouse(src.token_or_site, src.company))
            elif src.source_type == "lever":
                all_rows.extend(fetch_lever(src.token_or_site, src.company))
            elif src.source_type == "workable":
                all_rows.extend(fetch_workable_widget(src.token_or_site, src.company))
            else:
                manual_rows.append({
                    "company": src.company,
                    "career_portal_url": src.url,
                    "notes": f"Unknown source_type={src.source_type}"
                })
        except Exception as e:
            manual_rows.append({
                "company": src.company,
                "career_portal_url": src.url,
                "notes": f"ERROR while fetching ({src.source_type}): {repr(e)}"
            })

    df_jobs = pd.DataFrame(all_rows)
    df_manual = pd.DataFrame(manual_rows)

    if df_jobs.empty:
        # Ensure columns exist for downstream steps
        df_jobs = pd.DataFrame(columns=[
            "company","source_system","title","department","team","location","office",
            "remote_type","country_code","apply_url","job_url","updated_at","description_snippet"
        ])

    return df_jobs, df_manual


def enrich_and_filter(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["location_blob"] = (df["location"].fillna("") + " " + df["office"].fillna("")).str.strip()
    df["is_germany"] = df["location_blob"].apply(classify_germany)
    df["region_bucket"] = df["location_blob"].apply(classify_region)
    df["is_bw"] = df["location_blob"].apply(classify_bw)
    df["is_tech"] = df.apply(lambda r: is_tech_role(
        str(r.get("title","")),
        str(r.get("department","")),
        str(r.get("team","")),
        str(r.get("description_snippet",""))
    ), axis=1)

    # Convenience columns
    df["collected_at_utc"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return df


def export_excel(df_all: pd.DataFrame, df_manual: pd.DataFrame) -> None:
    # Split sheets
    df_germany = df_all[df_all["is_germany"]].copy()
    df_germany_tech = df_all[(df_all["is_germany"]) & (df_all["is_tech"])].copy()
    df_bw_tech = df_all[(df_all["region_bucket"] == "Baden-Württemberg") & (df_all["is_tech"])].copy()
    df_non_germany = df_all[~df_all["is_germany"]].copy()

    # Summary
    summary = []
    summary.append(("Total collected roles", len(df_all)))
    summary.append(("Germany roles", len(df_germany)))
    summary.append(("Germany tech roles", len(df_germany_tech)))
    summary.append(("Baden-Württemberg tech roles", len(df_bw_tech)))
    summary.append(("Non-Germany roles (informational)", len(df_non_germany)))
    df_summary = pd.DataFrame(summary, columns=["metric", "value"])

    # Write Excel
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        df_bw_tech.to_excel(writer, index=False, sheet_name="BW_Tech_Roles")
        df_germany_tech.to_excel(writer, index=False, sheet_name="Germany_Tech_Roles")
        df_germany.to_excel(writer, index=False, sheet_name="Germany_All_Roles")
        df_all.to_excel(writer, index=False, sheet_name="All_Roles_Raw")
        df_manual.to_excel(writer, index=False, sheet_name="Manual_Portals")

    print(f"\n✅ Excel exported locally: {OUTPUT_XLSX}")


def main():
    print("Collecting jobs from public feeds (Greenhouse / Lever / Workable) ...")
    df_jobs, df_manual = collect_all_jobs()
    df_jobs = enrich_and_filter(df_jobs)

    print("\nTop rows (preview):")
    if not df_jobs.empty:
        print(df_jobs[["company","title","location","region_bucket","is_tech"]].head(10).to_string(index=False))
    else:
        print("No jobs returned from feeds. See 'Manual_Portals' sheet in Excel for portals to check manually.")

    export_excel(df_jobs, df_manual)


if __name__ == "__main__":
    main()