# ============================================================
#         INVOICE DATA DOWNLOADER - GITHUB ACTIONS VERSION
# ============================================================
# Changes from local version:
#   • Headless browser (no display)
#   • Dynamic paths using environment/repo structure
#   • Dynamic date range based on fiscal year logic
#   • Start = 1 Apr of (previous fiscal year)
#   • End   = Last complete month before run date
#   • Saves files to reports/YYYY/MonthName/downloads/
# ============================================================

from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import calendar
import os
import time
import glob

# ============================================================
#                      CONFIGURATION
# ============================================================

LOGIN_URL        = "https://wisx.aaveg.co.in"
REPORT_URL       = "https://wisx.aaveg.co.in/Report/INV_702_InvoiceMIS?rptid_c=10"
CHANGE_FISCAL_URL = "https://wisx.aaveg.co.in/Home/ChangeFiscal"
LOGOUT_URL       = "https://wisx.aaveg.co.in/Account/Logout"

# Credentials from GitHub Secrets (set as environment variables)
USER_ID  = os.environ.get("WISX_USER_ID", "")
PASSWORD = os.environ.get("WISX_PASSWORD", "")

# ============================================================
#                    DATE RANGE LOGIC
# ============================================================

def get_fiscal_start_date(run_date):
    """
    Returns 1 April of the PREVIOUS fiscal year relative to run_date.

    Logic:
        If run_date is in Apr-Mar fiscal year 2026-27
        → Start = 1 Apr 2025

        If run_date is in Apr-Mar fiscal year 2027-28
        → Start = 1 Apr 2026

    Fiscal year starts in April.
    If run_date month >= 4 (Apr onwards) → current fiscal year start = Apr of same year
    If run_date month < 4 (Jan-Mar)      → current fiscal year start = Apr of previous year
    Then we go one more year back for the PREVIOUS fiscal year.
    """
    if run_date.month >= 4:
        current_fy_start_year = run_date.year
    else:
        current_fy_start_year = run_date.year - 1

    # Previous fiscal year start
    prev_fy_start_year = current_fy_start_year - 1

    return datetime(prev_fy_start_year, 4, 1)


def get_last_complete_month_end(run_date):
    """
    Returns the last day of the previous complete month.

    Example:
        Run on 10 Aug 2026 → 31 Jul 2026
        Run on 02 Aug 2027 → 31 Jul 2027
        Run on 10 Jan 2027 → 31 Dec 2026
    """
    first_of_current_month = run_date.replace(day=1)
    last_of_prev_month     = first_of_current_month - timedelta(days=1)
    return last_of_prev_month


def get_month_ranges(start_date, end_date):
    """
    Build list of monthly date ranges from start_date to end_date.
    Each entry = full calendar month within the range.

    Returns list of dicts:
        {
            'from_date'  : datetime,
            'to_date'    : datetime,
            'month_name' : 'Apr_2025'
        }
    """
    month_ranges = []
    current = start_date.replace(day=1)

    while current <= end_date:
        last_day_num = calendar.monthrange(current.year, current.month)[1]
        month_end    = current.replace(day=last_day_num)

        # Cap at end_date if needed
        actual_end = min(month_end, end_date)

        month_ranges.append({
            'from_date'  : current,
            'to_date'    : actual_end,
            'month_name' : current.strftime("%b_%Y")
        })

        # Move to next month
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1, day=1)
        else:
            current = current.replace(month=current.month + 1, day=1)

    return month_ranges


import glob  # Add this at top if not present

def get_download_directory(run_date):
    """
    Build download directory path and clear existing
    files to prevent duplicate month data.
    """
    repo_root    = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )
    year_str     = run_date.strftime("%Y")
    month_str    = run_date.strftime("%B")
    download_dir = os.path.join(
        repo_root, "reports",
        year_str, month_str, "downloads"
    )
    os.makedirs(download_dir, exist_ok=True)

    # Clear existing files to prevent duplicates
    existing = glob.glob(
        os.path.join(download_dir, "*.xlsx")
    )
    if existing:
        print(
            f"\n  🗑️  Clearing {len(existing)} "
            f"existing file(s) from downloads/..."
        )
        for f in existing:
            os.remove(f)
            print(
                f"     Removed: {os.path.basename(f)}"
            )
        print("  ✅ Downloads folder cleared.")

    return download_dir


# ============================================================
#                     HELPER FUNCTIONS
# ============================================================

def subtract_months(source_date, months):
    """Subtract N months from a date handling month-end edge cases."""
    month = source_date.month - months
    year  = source_date.year

    while month <= 0:
        month += 12
        year  -= 1

    day = min(source_date.day, calendar.monthrange(year, month)[1])
    return source_date.replace(year=year, month=month, day=day)


def select_date(page, selector, target_date):
    """
    Select a date from the datepicker widget.
    Navigates to the correct month and clicks the target day.
    """
    page.click(selector)
    page.wait_for_selector(".datepicker-days")

    target_month_year = target_date.strftime("%B %Y")
    target_day        = str(target_date.day)

    max_iterations = 192
    iteration      = 0

    while iteration < max_iterations:
        current_month_year = page.locator(
            ".datepicker-days .datepicker-switch"
        ).inner_text()

        if current_month_year == target_month_year:
            break
        elif datetime.strptime(current_month_year, "%B %Y") > target_date:
            page.locator(".datepicker-days th.prev").click()
        else:
            page.locator(".datepicker-days th.next").click()

        iteration += 1
        time.sleep(0.15)

    final_month_year = page.locator(
        ".datepicker-days .datepicker-switch"
    ).inner_text()

    if final_month_year != target_month_year:
        raise RuntimeError(
            f"Could not navigate to {target_month_year}. "
            f"Stuck at {final_month_year}."
        )

    page.locator(
        f".datepicker-days td.day:not(.old):not(.new):text-is('{target_day}')"
    ).click()
    page.mouse.click(10, 10)


def get_fiscal_year_for_date(target_date):
    """Convert a date into the fiscal year string used by the portal."""
    if target_date.month >= 4:
        return f"{target_date.year}-{target_date.year + 1}"
    return f"{target_date.year - 1}-{target_date.year}"


def find_fiscal_year_selector(page):
    """Locate the fiscal year dropdown on the Change Fiscal page."""
    candidate_selectors = [
        "#FiscalYear", "select#FiscalYear", "select[name='FiscalYear']",
        "#FiscalYearId", "select#FiscalYearId", "select[name='FiscalYearId']",
        "#ddlFiscalYear", "select#ddlFiscalYear", "select[name='ddlFiscalYear']"
    ]
    for selector in candidate_selectors:
        if page.query_selector(selector):
            return selector

    selects = page.query_selector_all("select")
    if len(selects) == 1:
        return "select"

    raise RuntimeError(
        "Unable to locate fiscal year dropdown on the Change Fiscal page."
    )


def get_current_fiscal_year(page):
    """Read the currently selected fiscal year from the portal."""
    page.goto(CHANGE_FISCAL_URL)
    try:
        page.wait_for_selector("text=Change Fiscal Year", timeout=10000)
    except Exception:
        page.wait_for_selector("text=Fiscal Year", timeout=10000)

    selector = find_fiscal_year_selector(page)
    current_fiscal = page.eval_on_selector(
        selector,
        "el => el.options[el.selectedIndex]?.text?.trim()"
    )
    return current_fiscal


def switch_fiscal_year(page, target_fiscal_year):
    """Change the active fiscal year in the portal."""
    print(f"    🔄 Switching fiscal year to {target_fiscal_year}...")
    page.goto(CHANGE_FISCAL_URL)
    try:
        page.wait_for_selector("text=Change Fiscal Year", timeout=10000)
    except Exception:
        page.wait_for_selector("text=Fiscal Year", timeout=10000)

    selector = find_fiscal_year_selector(page)
    page.locator(selector).select_option(label=target_fiscal_year)

    switch_button = page.locator(
        "button:has-text('Switch to Fiscal'), input[value='Switch to Fiscal']"
    )
    if switch_button.count() == 0:
        raise RuntimeError(
            "Could not find the 'Switch to Fiscal' button."
        )

    switch_button.click()
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    current = page.eval_on_selector(
        selector,
        "el => el.options[el.selectedIndex]?.text?.trim()"
    )
    if current != target_fiscal_year:
        raise RuntimeError(
            f"Fiscal year did not switch to {target_fiscal_year}. "
            f"Current: {current}."
        )

    print(f"    ✅ Fiscal year set to {current}.")
    return current


def download_single_month_report(page, from_date, to_date,
                                  save_directory, timestamp_str):
    """
    Download report for a specific month's date range.
    Returns the saved file path.
    """
    month_name = from_date.strftime("%b_%Y")

    page.goto(REPORT_URL)
    page.wait_for_selector("#DownloadType")

    page.select_option("#DownloadType", value="2")

    page.eval_on_selector("#FromDate", "el => el.removeAttribute('readonly')")
    page.eval_on_selector("#ToDate",   "el => el.removeAttribute('readonly')")

    select_date(page, "#FromDate", from_date)
    select_date(page, "#ToDate",   to_date)

    wisx_file_name = (
        f"InvoiceReport_702_{month_name}_{timestamp_str}.xlsx"
    )
    file_path = os.path.join(save_directory, wisx_file_name)

    print(
        f"    📥 Fetching: "
        f"{from_date.strftime('%d-%m-%Y')} to {to_date.strftime('%d-%m-%Y')}"
    )

    with page.expect_download(timeout=1800000) as download_info:
        page.locator("input[value='Export']").click(no_wait_after=True)

    download = download_info.value
    download.save_as(file_path)

    print(f"    💾 Saved: {wisx_file_name}")
    return file_path


# ============================================================
#                       MAIN FUNCTION
# ============================================================

def main():
    run_date = datetime.now()

    # ── Read Test Mode Flag ──────────────────────────────────
    # TEST_MODE env variable is set by GitHub Actions workflow
    # When true: download only last 1 complete month
    # When false: download full fiscal year range (production)
    test_mode_raw = os.environ.get("TEST_MODE", "false").strip().lower()
    test_mode     = test_mode_raw == "true"

    print("\n" + "=" * 65)
    print("         🗂️  WISX INVOICE REPORT DOWNLOADER (CI/CD)")
    print("=" * 65)
    print(f"\n  📅 Run Date   : {run_date.strftime('%d-%b-%Y')}")
    print(
        f"  🔧 Mode       : "
        f"{'🧪 TEST (last 1 month)' if test_mode else '🚀 PRODUCTION (full range)'}"
    )

    # ── Calculate Date Range ─────────────────────────────────
    if test_mode:
        # TEST MODE:
        # Download only last 1 complete month
        # Example: Run on 10 Aug 2026 → Download Jul 2026 only
        end_date   = get_last_complete_month_end(run_date)
        start_date = end_date.replace(day=1)

        print(f"\n  ⚠️  TEST MODE ACTIVE")
        print(f"  📅 Downloading only last complete month:")
        print(f"     {start_date.strftime('%d-%b-%Y')} → "
              f"{end_date.strftime('%d-%b-%Y')}")

    else:
        # PRODUCTION MODE:
        # Download from 1 Apr of previous fiscal year
        # to last complete month
        start_date = get_fiscal_start_date(run_date)
        end_date   = get_last_complete_month_end(run_date)

        print(f"\n  📅 Start Date : {start_date.strftime('%d-%b-%Y')}")
        print(f"  📅 End Date   : {end_date.strftime('%d-%b-%Y')}")

    # ── Build Month Ranges ───────────────────────────────────
    month_ranges = get_month_ranges(start_date, end_date)

    if not month_ranges:
        print("\n❌ No valid date ranges to download. Exiting.")
        raise SystemExit(1)

    print(f"\n  📦 Total months to download: {len(month_ranges)}")
    for i, mr in enumerate(month_ranges, 1):
        print(
            f"     {i:2}. {mr['month_name']:10}  "
            f"{mr['from_date'].strftime('%d-%m-%Y')} → "
            f"{mr['to_date'].strftime('%d-%m-%Y')}"
        )

    # ── Setup Save Directory ─────────────────────────────────
    download_dir  = get_download_directory(run_date)
    timestamp_str = run_date.strftime("%d%m%Y_%H%M")

    print(f"\n  📁 Save directory: {download_dir}")

    # ── Track Results ────────────────────────────────────────
    successful_downloads = 0
    failed_downloads     = []
    downloaded_files     = []

    print("\n" + "=" * 65)
    print("               🚀 STARTING DOWNLOAD PROCESS")
    print("=" * 65)

    with sync_playwright() as p:
        # Launch headless browser for GitHub Actions
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # ── Login ────────────────────────────────────────────
        print("\n🔐 [LOGIN] Connecting to WISX portal...")
        page.goto(LOGIN_URL)
        page.fill("#Email",    USER_ID)
        page.click("#Password")
        page.fill("#Password", PASSWORD)
        page.locator(
            "input[type='submit'][value='Log in']"
        ).click()
        page.wait_for_load_state("networkidle")
        print("✅ [LOGIN] Login successful!")

        current_fiscal_year = get_current_fiscal_year(page)
        print(
            f"✅ [FISCAL] Current fiscal year: "
            f"{current_fiscal_year}"
        )

        # ── Download Each Month ──────────────────────────────
        print("\n" + "-" * 65)
        print("                    DOWNLOADING REPORTS")
        print("-" * 65)

        for i, month_range in enumerate(month_ranges, 1):
            target_fiscal_year = get_fiscal_year_for_date(
                month_range['from_date']
            )

            if current_fiscal_year != target_fiscal_year:
                current_fiscal_year = switch_fiscal_year(
                    page, target_fiscal_year
                )

            print(
                f"\n📦 [{i}/{len(month_ranges)}] "
                f"Processing: {month_range['month_name']}"
            )

            try:
                file_path = download_single_month_report(
                    page,
                    month_range['from_date'],
                    month_range['to_date'],
                    download_dir,
                    timestamp_str
                )

                successful_downloads += 1
                downloaded_files.append(file_path)
                print(f"    ✅ Status: SUCCESS")

                progress = (i / len(month_ranges)) * 100
                print(f"    📊 Progress: {progress:.0f}% complete")

                if i < len(month_ranges):
                    time.sleep(2)

            except Exception as e:
                print(f"    ❌ Status: FAILED")
                print(f"    ⚠️  Error: {str(e)}")
                failed_downloads.append(month_range['month_name'])

        # ── Logout ───────────────────────────────────────────
        print("\n" + "-" * 65)
        print("🔓 [LOGOUT] Logging out...")
        page.goto(LOGOUT_URL)
        page.wait_for_load_state("networkidle")
        print("✅ [LOGOUT] Logged out successfully!")
        browser.close()

    # ── Final Summary ────────────────────────────────────────
    print("\n" + "=" * 65)
    print("                    📊 DOWNLOAD SUMMARY")
    print("=" * 65)
    print(
        f"\n  🔧 Mode       : "
        f"{'🧪 TEST' if test_mode else '🚀 PRODUCTION'}"
    )
    print(
        f"  ✅ Successful : "
        f"{successful_downloads}/{len(month_ranges)}"
    )

    if failed_downloads:
        print(f"  ❌ Failed     : {failed_downloads}")
    else:
        print("  ❌ Failed     : None")

    print(f"\n  📁 Files saved to: {download_dir}")
    print("=" * 65)

    # ── Write output path for next step ─────────────────────
    # Save RELATIVE path so followup workflow on a
    # different runner can reconstruct the absolute path
    repo_root  = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )
    state_file = os.path.join(
        repo_root, "reports", ".current_run_dir.txt"
    )

    # Convert to relative path from repo root
    relative_download_dir = os.path.relpath(
        download_dir, repo_root
    )

    with open(state_file, "w") as f:
        f.write(relative_download_dir)

    print(
        f"\n  📝 Run directory saved (relative): "
        f"{relative_download_dir}"
    )
    print(f"  📝 State file: {state_file}")

    print(f"\n  📝 Run directory saved to: {state_file}")

    if failed_downloads:
        raise SystemExit(
            f"⚠️ {len(failed_downloads)} month(s) failed."
        )


if __name__ == "__main__":
    main()
