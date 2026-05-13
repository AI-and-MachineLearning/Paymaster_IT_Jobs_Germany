# Paymaster IT Jobs Germany

A Python scraper that collects IT/Technology job openings from Employer of Record (EOR) / Paymaster vendors operating in Germany, with a focus on Baden-Württemberg. Results are exported to a structured Excel workbook for easy analysis.

## What It Does

- Fetches public job listings from Greenhouse, Lever, and Workable APIs
- Filters jobs by location (Germany-wide, with a dedicated Baden-Württemberg view)
- Filters jobs by role type (tech roles only, excludes sales, HR, finance, etc.)
- Exports results to a timestamped Excel file with multiple categorized sheets

## Prerequisites

- Python 3.9 or higher
- Internet connection (all APIs are public — no authentication required)

## Setup

### 1. Clone the Repository

```bash
git clone <repository-url>
cd Paymaster_IT_Jobs_Germany
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
```

### 3. Activate the Virtual Environment

**Windows (Command Prompt):**
```cmd
.venv\Scripts\activate
```

**Windows (PowerShell):**
```powershell
.venv\Scripts\Activate.ps1
```

**macOS / Linux:**
```bash
source .venv/bin/activate
```

### 4. Install Dependencies

```bash
pip install requests pandas openpyxl
```

## Running the Script

### Standard Version

```bash
python paymaster-it-jobs-in-germany.py
```

### Enhanced Version (recommended — better Unicode/special character handling)

```bash
python paymaster-it-jobs-in-germany-enhanced.py
```

The script prints progress to the console and writes a timestamped Excel file to the current directory when done, e.g.:

```
paymaster_IT_jobs_Germany_20260513_1025.xlsx
```

Runtime is typically 5–30 seconds depending on API response times.

## Output

The generated Excel workbook contains six sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Job counts by region and tech-role classification |
| **BW_Tech_Roles** | Tech roles in Baden-Württemberg (primary view) |
| **Germany_Tech_Roles** | All tech roles across Germany |
| **Germany_All_Roles** | All German jobs, including non-tech roles |
| **All_Roles_Raw** | Unfiltered dataset from all sources |
| **Manual_Portals** | Companies that require manual searching (e.g., Workday portals) |

## Supported Job Platforms

| Platform | How It's Used |
|---|---|
| **Greenhouse** | Public Job Board API (`boards.greenhouse.io/v1/boards/{token}/jobs`) |
| **Lever** | Public Postings API (`jobs.lever.co/v0/postings/{site}`) with pagination |
| **Workable** | Public Widget API (`apply.workable.com/api/v3/accounts/{account}/jobs`) |

To add a new company, add an entry to the `SOURCES` list inside the script and set the appropriate `type` (`greenhouse`, `lever`, or `workable`).

## Project Structure

```
Paymaster_IT_Jobs_Germany/
├── paymaster-it-jobs-in-germany.py          # Original scraper script
├── paymaster-it-jobs-in-germany-enhanced.py # Enhanced version (Unicode normalization)
├── .venv/                                   # Python virtual environment
└── *.xlsx                                   # Generated output files (timestamped)
```

## Configuration

All configuration is defined as constants at the top of each script. No environment variables or external config files are needed.

| Setting | Description |
|---|---|
| `OUTPUT_XLSX` | Output filename pattern (auto-timestamped) |
| `GERMANY_ALIASES` | Keywords used to detect Germany-based locations |
| `BW_KEYWORDS` | City and region names used to detect Baden-Württemberg |
| `TECH_INCLUDE_REGEX` | Regex to identify tech roles (software, engineer, developer, etc.) |
| `TECH_EXCLUDE_REGEX` | Regex to exclude non-tech roles (sales, HR, finance, etc.) |
| `SOURCES` | List of companies and their job board API endpoints |

## No Authentication Required

All job board APIs used by this project are publicly accessible — no API keys, tokens, or login credentials are needed.

## Troubleshooting

**SSL errors on corporate networks**

Some corporate proxies interrupt SSL verification. If you see SSL errors, try:
```bash
pip install --upgrade certifi
```

**Script runs but Excel file is empty or has few results**

- Check your internet connection.
- One or more APIs may be temporarily unavailable — the script logs skipped sources to the console.
- The company's job board may have migrated to a different platform; update the `SOURCES` list accordingly.

**`ModuleNotFoundError`**

Make sure the virtual environment is activated and dependencies are installed:
```bash
pip install requests pandas openpyxl
```

**Output file won't open (locked)**

Close any previously opened Excel output file before running the script again, as Excel locks the file while it is open.
