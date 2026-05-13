# paymaster_jobs_germany_it_export.py
# Collect IT/Technology/Software job openings in Germany (focus BW) and export to Excel.
#
# Sources (public):
# - Greenhouse Job Board API: https://boards-api.greenhouse.io/v1/boards/{token}/jobs?content=true
# - Lever Postings API:       https://api.lever.co/v0/postings/{site}?mode=json
# - Workable widget JSON:     https://apply.workable.com/api/v1/widget/accounts/{account}
#
# Requirements:
#   pip install requests pandas openpyxl

import re
import time
import unicodedata
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

import requests
import pandas as pd


# ----------------------------
# OUTPUT
# ----------------------------
OUTPUT_XLSX = "paymaster_IT_jobs_Germany_{ts}.xlsx".format(
    ts=datetime.now().strftime("%Y%m%d_%H%M")
)

# ----------------------------
# LOCATION & FILTER SETTINGS
# ----------------------------

GERMANY_ALIASES = [
    "germany", "deutschland", "remote - germany", "remote (germany)", "remote, germany", ", de"
]

# Baden-Wuerttemberg focus (ASCII only on purpose)
BW_KEYWORDS = [
    "baden-wuerttemberg", "baden wuerttemberg", "bw",
    "stuttgart", "karlsruhe", "mannheim", "heidelberg", "freiburg", "ulm",
    "heilbronn", "reutlingen", "tuebingen", "pforzheim", "konstanz",
    "offenburg", "esslingen", "boeblingen", "ludwigsburg", "ravensburg"
]

# Tech include/exclude (kept as simple one-line strings to avoid parser problems)
TECH_INCLUDE_PATTERN = (
    r"\b(software|engineer|engineering|developer|development|devops|sre|platform|"
    r"data|analytics|machine learning|ml|ai|artificial intelligence|"
    r"security|infosec|cyber|cloud|infrastructure|it\b|sysadmin|"
    r"qa\b|quality|test|automation|architect|site reliability|"
    r"backend|frontend|front-end|fullstack|full-stack|mobile|ios|android|"
    r"database|dba|etl|bi\b|observability|kubernetes)\b"
)

# Exclusions (tuned to remove obvious non-tech functions)
TECH_EXCLUDE_PATTERN = (
    r"\b(sales|account executive|business development|marketing|"
    r"hr\b|people\b|recruit|talent acquisition|customer success|"
    r"payroll specialist\b|legal|finance|accounting)\b"
)

TECH_INCLUDE_REGEX = re.compile(TECH_INCLUDE_PATTERN, re.IGNORECASE)
TECH_EXCLUDE_REGEX = re.compile(TECH_EXCLUDE_PATTERN, re.IGNORECASE)

# Allowlist for tech roles that might contain excluded words
TECH_ALLOWLIST_REGEX = re.compile(
    r"\b(payroll engineer|software engineer|data engineer|security engineer|platform engineer)\b",
    re.IGNORECASE
)

# ----------------------------
# COMPANY SOURCES
# ----------------------------
# source_type: "greenhouse" | "lever" | "workable" | "manual"
SOURCES: List[Dict[str, str]] = [
    {
        "company": "Omnipresent",
        "source_type": "greenhouse",
        "token_or_site": "omnipresent",
        "portal_url": "https://boards.greenhouse.io/omnipresent"
    },
    {
        "company": "Remote",
        "source_type": "greenhouse",
        "token_or_site": "remotereferralboardinternaluseonly",
        "portal_url": "https://job-boards.greenhouse.io/remotereferralboardinternaluseonly"
    },
    {
        "company": "RemoFirst",
        "source_type": "lever",
        "token_or_site": "remofirst",
        "portal_url": "https://jobs.lever.co/remofirst"
    },
    {
        "company": "WorkMotion",
        "source_type": "workable",
        "token_or_site": "workmotion",
        "portal_url": "https://apply.workable.com/workmotion"
    },

    # Manual-only (kept in Excel for your tracking)
    {"company": "Deel", "source_type": "manual", "token_or_site": "", "portal_url": "https://www.deel.com/careers/"},
    {"company": "Safeguard Global", "source_type": "manual", "token_or_site": "", "portal_url": "https://safeguardglobal.wd3.myworkdayjobs.com/External_Careers/"},
    {"company": "CloudPay", "source_type": "manual", "token_or_site": "", "portal_url": "https://cloudpay.wd3.myworkdayjobs.com/CloudPay_External"},
    {"company": "TMF Group", "source_type": "manual", "token_or_site": "", "portal_url": "https://www.tmf-group.com/en/careers/"},
    {"company": "SD Worx", "source_type": "manual", "token_or_site": "", "portal_url": "https://careers.sdworx.com/jobs"},
    {"company": "Multiplier", "source_type": "manual", "token_or_site": "", "portal_url": "https://www.usemultiplier.com/careers"},
    {"company": "Rippling", "source_type": "manual", "token_or_site": "", "portal_url": "https://www.rippling.com/careers"},
    {"company": "Velocity Global / Pebl", "source_type": "manual", "token_or_site": "", "portal_url": "https://hellopebl.com/company/careers/"},
    {"company": "Lano", "source_type": "manual", "token_or_site": "", "portal_url": "https://www.lano.io/careers"},
]


# ----------------------------
# HTTP
# ----------------------------
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; PaymasterJobCollector/1.0)"
})

def http_get_json(url: str, params: Optional[Dict[str, Any]] = None, timeout: int = 30) -> Any:
    r = SESSION.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()


# ----------------------------
# NORMALIZATION & CLASSIFIERS
# ----------------------------
def normalize_text(s: str) -> str:
    if s is None:
        s = ""
    # Normalize unicode (converts fancy hyphens to standard where possible)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u2011", "-").replace("\u2013", "-").replace("\u2014", "-")
    return s.strip()

def is_germany(location_blob: str) -> bool:
    t = normalize_text(location_blob).lower()
    for alias in GERMANY_ALIASES:
        if alias in t:
            return True
    # additional simple heuristic
    if " germany" in t:
        return True
    return False

def is_bw(location_blob: str) -> bool:
    t = normalize_text(location_blob).lower()
    for k in BW_KEYWORDS:
        if k in t:
            return True
    return False

def region_bucket(location_blob: str) -> str:
    if not is_germany(location_blob):
        return "Non-Germany"
    if is_bw(location_blob):
        return "Baden-Wuerttemberg"
    if "remote" in normalize_text(location_blob).lower():
        return "Germany (Remote/Unspecified)"
    return "Germany (Other State/Unspecified)"

def is_tech_role(title: str, department: str, team: str, description: str) -> bool:
    blob = normalize_text(" ".join([title or "", department or "", team or "", description or ""]))
    if TECH_INCLUDE_REGEX.search(blob) is None:
        return False
    if TECH_ALLOWLIST_REGEX.search(blob) is not None:
        return True
    if TECH_EXCLUDE_REGEX.search(blob) is not None:
        return False
    return True


# ----------------------------
# FETCHERS
# ----------------------------
def fetch_greenhouse(board_token: str, company: str) -> List[Dict[str, Any]]:
    url = "https://boards-api.greenhouse.io/v1/boards/{tok}/jobs".format(tok=board_token)
    data = http_get_json(url, params={"content": "true"})
    jobs = data.get("jobs", []) if isinstance(data, dict) else []
    rows: List[Dict[str, Any]] = []

    for j in jobs:
        title = normalize_text(j.get("title", ""))
        loc = normalize_text((j.get("location") or {}).get("name", ""))
        job_url = normalize_text(j.get("absolute_url", ""))
        desc = normalize_text(j.get("content", ""))

        departments = j.get("departments") or []
        dept_names = ", ".join([d.get("name", "") for d in departments if isinstance(d, dict) and d.get("name")])
        offices = j.get("offices") or []
        office_names = ", ".join([o.get("name", "") for o in offices if isinstance(o, dict) and o.get("name")])

        updated = normalize_text(j.get("updated_at", ""))

        rows.append({
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
    return rows


def fetch_lever(site: str, company: str) -> List[Dict[str, Any]]:
    base = "https://api.lever.co/v0/postings/{site}".format(site=site)
    rows: List[Dict[str, Any]] = []
    skip = 0
    limit = 100

    while True:
        data = http_get_json(base, params={"mode": "json", "skip": skip, "limit": limit})
        if not isinstance(data, list) or len(data) == 0:
            break

        for j in data:
            title = normalize_text(j.get("text", ""))
            cats = j.get("categories") or {}
            loc = normalize_text(cats.get("location", ""))
            team = normalize_text(cats.get("team", ""))
            dept = normalize_text(cats.get("department", ""))
            country = normalize_text(j.get("country", ""))
            job_url = normalize_text(j.get("hostedUrl", ""))
            apply_url = normalize_text(j.get("applyUrl", "")) or job_url
            remote_type = normalize_text(j.get("workplaceType", ""))
            desc = normalize_text(j.get("descriptionPlain", ""))

            rows.append({
                "company": company,
                "source_system": "lever",
                "title": title,
                "department": dept,
                "team": team,
                "location": loc,
                "office": "",
                "remote_type": remote_type,
                "country_code": country,
                "apply_url": apply_url,
                "job_url": job_url,
                "updated_at": "",
                "description_snippet": desc[:1000],
            })

        if len(data) < limit:
            break
        skip += limit
        time.sleep(0.25)

    return rows


def fetch_workable_widget(account: str, company: str) -> List[Dict[str, Any]]:
    url = "https://apply.workable.com/api/v1/widget/accounts/{acct}".format(acct=account)
    data = http_get_json(url)
    jobs = data.get("jobs", []) if isinstance(data, dict) else []
    rows: List[Dict[str, Any]] = []

    for j in jobs:
        title = normalize_text(j.get("title", ""))
        loc = normalize_text(j.get("location", ""))
        dept = normalize_text(j.get("department", ""))
        job_url = normalize_text(j.get("url", ""))
        apply_url = normalize_text(j.get("application_url", "")) or job_url
        desc = normalize_text(j.get("short_description", "")) or normalize_text(j.get("description", ""))

        rows.append({
            "company": company,
            "source_system": "workable_widget",
            "title": title,
            "department": dept,
            "team": "",
            "location": loc,
            "office": "",
            "remote_type": "remote" if "remote" in loc.lower() else "",
            "country_code": "",
            "apply_url": apply_url,
            "job_url": job_url,
            "updated_at": "",
            "description_snippet": desc[:1000],
        })
    return rows


# ----------------------------
# PIPELINE
# ----------------------------
def collect_jobs() -> Tuple[pd.DataFrame, pd.DataFrame]:
    all_rows: List[Dict[str, Any]] = []
    manual_rows: List[Dict[str, Any]] = []

    for src in SOURCES:
        company = src["company"]
        stype = src["source_type"]
        token = src["token_or_site"]
        portal = src["portal_url"]

        if stype == "manual":
            manual_rows.append({"company": company, "career_portal_url": portal, "notes": "Manual portal"})
            continue

        try:
            if stype == "greenhouse":
                all_rows.extend(fetch_greenhouse(token, company))
            elif stype == "lever":
                all_rows.extend(fetch_lever(token, company))
            elif stype == "workable":
                all_rows.extend(fetch_workable_widget(token, company))
            else:
                manual_rows.append({"company": company, "career_portal_url": portal, "notes": "Unknown source_type"})
        except Exception as e:
            manual_rows.append({"company": company, "career_portal_url": portal, "notes": "Fetch error: {0}".format(repr(e))})

    df_all = pd.DataFrame(all_rows)
    df_manual = pd.DataFrame(manual_rows)

    if df_all.empty:
        df_all = pd.DataFrame(columns=[
            "company","source_system","title","department","team","location","office","remote_type",
            "country_code","apply_url","job_url","updated_at","description_snippet"
        ])

    return df_all, df_manual


def enrich_filter(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["location_blob"] = (df["location"].fillna("") + " " + df["office"].fillna("")).str.strip()
    df["is_germany"] = df["location_blob"].apply(is_germany)
    df["is_bw"] = df["location_blob"].apply(is_bw)
    df["region_bucket"] = df["location_blob"].apply(region_bucket)

    df["is_tech"] = df.apply(lambda r: is_tech_role(
        str(r.get("title","")),
        str(r.get("department","")),
        str(r.get("team","")),
        str(r.get("description_snippet",""))
    ), axis=1)

    df["collected_at_utc"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return df


def export_excel(df_all: pd.DataFrame, df_manual: pd.DataFrame) -> None:
    df_germany = df_all[df_all["is_germany"]].copy()
    df_germany_tech = df_all[(df_all["is_germany"]) & (df_all["is_tech"])].copy()
    df_bw_tech = df_all[(df_all["region_bucket"] == "Baden-Wuerttemberg") & (df_all["is_tech"])].copy()

    df_summary = pd.DataFrame([
        {"metric": "Total collected roles", "value": len(df_all)},
        {"metric": "Germany roles", "value": len(df_germany)},
        {"metric": "Germany tech roles", "value": len(df_germany_tech)},
        {"metric": "Baden-Wuerttemberg tech roles", "value": len(df_bw_tech)},
        {"metric": "Manual portals listed", "value": len(df_manual)},
    ])

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        df_bw_tech.to_excel(writer, index=False, sheet_name="BW_Tech_Roles")
        df_germany_tech.to_excel(writer, index=False, sheet_name="Germany_Tech_Roles")
        df_germany.to_excel(writer, index=False, sheet_name="Germany_All_Roles")
        df_all.to_excel(writer, index=False, sheet_name="All_Roles_Raw")
        df_manual.to_excel(writer, index=False, sheet_name="Manual_Portals")

    print("OK - Excel exported locally: {0}".format(OUTPUT_XLSX))


def main():
    print("Collecting jobs from public feeds...")
    df_all, df_manual = collect_jobs()
    df_all = enrich_filter(df_all)

    if len(df_all) == 0:
        print("No jobs returned from public feeds. Check Manual_Portals in Excel output.")
    else:
        print("Preview:")
        print(df_all[["company","title","location","region_bucket","is_tech"]].head(12).to_string(index=False))

    export_excel(df_all, df_manual)


if __name__ == "__main__":
    main()