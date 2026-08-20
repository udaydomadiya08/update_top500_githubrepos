import os
import time
import requests
import openpyxl
from datetime import datetime, timezone

# ============================================================
# GITHUB TOP 500 REPOSITORIES → EXCEL
# ============================================================

API_URL = "https://api.github.com/search/repositories"
EXCEL_FILE = "github_top_repos_by_interval.xlsx"

# ============================================================
# GITHUB TOKEN
# ============================================================
# GitHub Actions takes your repository secret GROW_HUB
# and passes it to this Python program as GITHUB_TOKEN.
#
# Workflow:
# GROW_HUB secret
#       ↓
# GITHUB_TOKEN environment variable
#       ↓
# Python
# ============================================================

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

if not GITHUB_TOKEN:
    raise RuntimeError(
        "GITHUB_TOKEN environment variable not found. "
        "Check the GitHub Actions env configuration."
    )

HEADERS = {
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "Authorization": f"Bearer {GITHUB_TOKEN}"
}

# ============================================================
# FETCH TOP 500 REPOSITORIES
# ============================================================

print("=" * 70)
print("🚀 GITHUB TOP 500 REPOSITORY UPDATE")
print("=" * 70)

all_repos = []

for page in range(1, 6):

    print(f"📡 Fetching page {page}/5...")

    params = {
        "q": "stars:>0",
        "sort": "stars",
        "order": "desc",
        "per_page": 100,
        "page": page
    }

    response = requests.get(
        API_URL,
        params=params,
        headers=HEADERS,
        timeout=60
    )

    # Show useful error if GitHub rejects request
    if response.status_code != 200:
        print("❌ GitHub API request failed")
        print(f"HTTP Status: {response.status_code}")
        print(f"Response: {response.text[:1000]}")
        response.raise_for_status()

    data = response.json()

    repos = data.get("items", [])

    all_repos.extend(repos)

    print(
        f"✅ Page {page}/5 fetched — "
        f"{len(repos)} repositories"
    )

    time.sleep(1)

# ============================================================
# REMOVE DUPLICATES
# ============================================================

unique_repos = {
    repo["id"]: repo
    for repo in all_repos
}

all_repos = list(unique_repos.values())

# ============================================================
# SORT BY STARS
# ============================================================

all_repos.sort(
    key=lambda x: x["stargazers_count"],
    reverse=True
)

all_repos = all_repos[:500]

print()
print(f"✅ Retrieved {len(all_repos)} unique repositories")

if len(all_repos) == 0:
    raise RuntimeError("GitHub API returned zero repositories.")

# ============================================================
# CURRENT SNAPSHOT DATE
# ============================================================

today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

print(f"📅 Snapshot date: {today}")

# ============================================================
# BUILD ROWS
# ============================================================

rows = []

for rank, repo in enumerate(all_repos, start=1):

    rows.append([
        today,
        rank,
        repo.get("full_name", ""),
        repo.get("stargazers_count", 0),
        repo.get("forks_count", 0),
        repo.get("watchers_count", 0),
        repo.get("open_issues_count", 0),
        repo.get("language") or "Unknown",
        repo.get("description") or "",
        repo.get("created_at", "")[:10],
        repo.get("updated_at", "")[:10],
        repo.get("pushed_at", "")[:10],
        repo.get("html_url", "")
    ])

# ============================================================
# OPEN EXISTING EXCEL WORKBOOK
# ============================================================

print()
print(f"📂 Opening: {EXCEL_FILE}")

if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(
        f"Excel file not found: {EXCEL_FILE}"
    )

wb = openpyxl.load_workbook(EXCEL_FILE)

print(
    "📑 Existing sheets:",
    ", ".join(wb.sheetnames)
)

# ============================================================
# DAILY SHEET
# ============================================================

if "Daily" in wb.sheetnames:

    ws = wb["Daily"]

    print("✅ Using existing 'Daily' sheet")

else:

    ws = wb.create_sheet("Daily")

    print("🆕 Created 'Daily' sheet")

# ============================================================
# HEADERS
# ============================================================

headers = [
    "Date",
    "Rank",
    "Repository",
    "Stars",
    "Forks",
    "Watchers",
    "Open Issues",
    "Language",
    "Description",
    "Created",
    "Last Updated",
    "Last Push",
    "URL"
]

# Check whether the sheet already has headers
first_row = [
    ws.cell(row=1, column=i).value
    for i in range(1, len(headers) + 1)
]

if first_row != headers:

    print("📝 Setting/updating Daily sheet headers")

    for col, header in enumerate(headers, start=1):
        ws.cell(
            row=1,
            column=col,
            value=header
        )

# ============================================================
# CHECK FOR DUPLICATE DATE
# ============================================================

print()
print(f"🔍 Checking whether {today} already exists...")

existing_dates = set()

if ws.max_row >= 2:

    for row in ws.iter_rows(
        min_row=2,
        min_col=1,
        max_col=1,
        values_only=True
    ):

        value = row[0]

        if value is not None:

            if hasattr(value, "strftime"):
                date_value = value.strftime("%Y-%m-%d")
            else:
                date_value = str(value)[:10]

            existing_dates.add(date_value)

# ============================================================
# APPEND TODAY'S DATA
# ============================================================

if today in existing_dates:

    print(
        f"⚠️ A snapshot for {today} already exists."
    )

    print(
        "⚠️ No duplicate rows will be added."
    )

else:

    print(
        f"📊 Adding {len(rows)} repositories "
        f"for {today}..."
    )

    for row_data in rows:
        ws.append(row_data)

    print(
        f"✅ Added {len(rows)} repositories "
        f"for {today}"
    )

# ============================================================
# EXCEL FORMATTING
# ============================================================

ws.freeze_panes = "A2"

# Auto-filter
ws.auto_filter.ref = ws.dimensions

# Column widths
widths = {
    "A": 14,
    "B": 8,
    "C": 45,
    "D": 14,
    "E": 14,
    "F": 14,
    "G": 16,
    "H": 18,
    "I": 60,
    "J": 14,
    "K": 16,
    "L": 14,
    "M": 60
}

for column, width in widths.items():
    ws.column_dimensions[column].width = width

# ============================================================
# SAVE SAME EXCEL FILE
# ============================================================

print()
print(f"💾 Saving: {EXCEL_FILE}")

wb.save(EXCEL_FILE)

# ============================================================
# FINAL STATUS
# ============================================================

print()
print("=" * 70)
print("🎉 COMPLETE")
print("=" * 70)
print(f"📅 Date: {today}")
print(f"📊 Current Top 500: {len(rows)} repositories")
print(f"📁 Excel: {EXCEL_FILE}")
print(f"📑 Sheet: Daily")
print(f"📈 Total Daily rows: {ws.max_row - 1}")
print("=" * 70)
