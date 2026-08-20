import os
import time
import requests
import openpyxl
from datetime import datetime, timezone

# ============================================================
# CONFIGURATION
# ============================================================

API_URL = "https://api.github.com/search/repositories"

EXCEL_FILE = "github_top_repos_by_interval.xlsx"

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

if not GITHUB_TOKEN:
    raise RuntimeError("GITHUB_TOKEN environment variable not found.")

HEADERS = {
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "Authorization": f"Bearer {GITHUB_TOKEN}"
}

# ============================================================
# FETCH TOP 500 REPOSITORIES
# ============================================================

print("🚀 Fetching GitHub Top 500 repositories...")

all_repos = []

for page in range(1, 6):

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

    response.raise_for_status()

    data = response.json()

    repos = data["items"]

    all_repos.extend(repos)

    print(
        f"✅ Page {page}/5 fetched — "
        f"{len(repos)} repositories"
    )

    time.sleep(1)

# Remove duplicates
unique_repos = {
    repo["id"]: repo
    for repo in all_repos
}

all_repos = list(unique_repos.values())

# Sort by stars
all_repos.sort(
    key=lambda x: x["stargazers_count"],
    reverse=True
)

all_repos = all_repos[:500]

print(f"✅ Retrieved {len(all_repos)} repositories")

# ============================================================
# CURRENT SNAPSHOT
# ============================================================

today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

rows = []

for rank, repo in enumerate(all_repos, start=1):

    rows.append([
        today,
        rank,
        repo["full_name"],
        repo["stargazers_count"],
        repo["forks_count"],
        repo["watchers_count"],
        repo["open_issues_count"],
        repo["language"] or "Unknown",
        repo["description"] or "",
        repo["created_at"][:10],
        repo["updated_at"][:10],
        repo["pushed_at"][:10],
        repo["html_url"]
    ])

# ============================================================
# OPEN EXISTING EXCEL
# ============================================================

print(f"📂 Opening {EXCEL_FILE}")

wb = openpyxl.load_workbook(EXCEL_FILE)

# ============================================================
# DAILY SHEET
# ============================================================

if "Daily" in wb.sheetnames:
    ws = wb["Daily"]
else:
    ws = wb.create_sheet("Daily")

# ============================================================
# CREATE HEADER IF EMPTY
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

if ws.max_row == 1 and ws.cell(1, 1).value is None:

    for col, header in enumerate(headers, start=1):
        ws.cell(
            row=1,
            column=col,
            value=header
        )

# ============================================================
# APPEND TODAY'S DATA
# ============================================================

print(f"📊 Adding snapshot for {today}")

# Prevent duplicate snapshot if workflow accidentally runs twice
existing_dates = set()

for row in ws.iter_rows(
    min_row=2,
    min_col=1,
    max_col=1,
    values_only=True
):
    if row[0]:
        existing_dates.add(str(row[0])[:10])

if today in existing_dates:

    print(
        f"⚠️ Data for {today} already exists. "
        "Skipping duplicate."
    )

else:

    for row_data in rows:

        ws.append(row_data)

    print(
        f"✅ Added {len(rows)} repositories "
        f"for {today}"
    )

# ============================================================
# FORMAT
# ============================================================

ws.freeze_panes = "A2"

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

wb.save(EXCEL_FILE)

print("=" * 70)
print("🎉 COMPLETE")
print("=" * 70)
print(f"📅 Date: {today}")
print(f"📊 Repositories added: {len(rows)}")
print(f"📁 Updated file: {EXCEL_FILE}")
print("=" * 70)
