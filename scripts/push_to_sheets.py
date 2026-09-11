# ============================================================
#     GOOGLE SHEETS PUSHER - GITHUB ACTIONS VERSION (FIXED)
# ============================================================
# Changes vs previous version:
#   • Pushes ONLY CRD vertical rows (was: all verticals)
#   • sheets_pushed = True ONLY when every year succeeds
#   • Exits non-zero (fails the workflow step) if any
#     year fails or has no sheet config  → no false "success",
#     no dead README link, next trigger will retry
#   • Reads directly from parquet, batch writing (unchanged)
# ============================================================

import os
import sys
import json
import glob
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================================
#                    CONFIGURATION
# ============================================================

SHEET_CONFIG = {
    2025: {
        "spreadsheet_id": "1CAOhOmdW1BXgaa8IDc5os8m3tXjHQ0pblPdoPnAGxDs",
        "sheet_name"    : "INVOICE 2025"
    },
    2026: {
        "spreadsheet_id": "1cOEVewBvnmaxt58oa4ZSD55r-r6dlzoqBq8Pp_nuU9g",
        "sheet_name"    : "INVOICE 2026"
    },
    # 2027: {
    #     "spreadsheet_id": "YOUR_2027_SHEET_ID",
    #     "sheet_name"    : "INVOICE 2027"
    # },
}

SHEETS_COLUMNS = [
    'Rental', 'ReportingTime', 'Month',
    'BookingNo', 'CorporateID', 'Corporate',
    'Group Entity', 'Vertical',
    'New GST', 'Revenue',
    'Hub', 'Group Hub'
]

DRIVE_ROOT_FOLDER = "CRD Report Automation"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


# ============================================================
#                    PATH HELPERS
# ============================================================

def get_repo_root():
    return os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )


def resolve_path(path):
    """Resolve relative or old-runner absolute path."""
    if not path:
        return path
    repo_root = get_repo_root()
    if os.path.isabs(path):
        if os.path.exists(path):
            return path
        for marker in [
            'reports' + os.sep, 'reports/',
            'data' + os.sep,    'data/'
        ]:
            idx = path.find(marker)
            if idx != -1:
                return os.path.join(
                    repo_root, path[idx:]
                )
        return path
    return os.path.join(repo_root, path)


def read_state_file():
    """Read and resolve all paths in state file."""
    repo_root  = get_repo_root()
    state_path = os.path.join(
        repo_root, "reports", ".run_state.json"
    )
    if not os.path.exists(state_path):
        raise FileNotFoundError(
            f"❌ State file not found: {state_path}"
        )
    with open(state_path, "r") as f:
        state = json.load(f)

    for key in [
        "output_dir", "combined_file_path",
        "parquet_file_path",
        "not_found_summary_path",
        "not_found_json_path",
        "sample_invoice_path",
        "final_report_path",
        "final_sample_invoice_path"
    ]:
        if state.get(key):
            state[key] = resolve_path(state[key])

    return state, state_path


def update_state_file(state_path, updates):
    """Update specific keys in state file."""
    repo_root = get_repo_root()
    with open(state_path, "r") as f:
        state = json.load(f)

    def to_rel(path):
        if not path:
            return path
        try:
            return os.path.relpath(path, repo_root)
        except ValueError:
            return path

    for key, val in updates.items():
        if isinstance(val, str) and (
            key.endswith('_path') or
            key == 'output_dir'
        ):
            state[key] = to_rel(val) if val else val
        else:
            state[key] = val

    with open(state_path, "w") as f:
        json.dump(state, f, indent=4)
    print(f"      📝 State file updated.")


# ============================================================
#                    GOOGLE AUTH
# ============================================================

def get_google_credentials():
    """Build credentials from GitHub Secret."""
    sa_json = os.environ.get(
        "GOOGLE_SERVICE_ACCOUNT_JSON", ""
    ).strip()
    if not sa_json:
        raise ValueError(
            "❌ GOOGLE_SERVICE_ACCOUNT_JSON not set."
        )
    sa_info     = json.loads(sa_json)
    credentials = (
        service_account.Credentials
        .from_service_account_info(
            sa_info, scopes=SCOPES
        )
    )
    print(
        f"      ✅ Auth: "
        f"{sa_info.get('client_email', '?')}"
    )
    return credentials


# ============================================================
#                    GOOGLE DRIVE
# ============================================================

def get_or_create_folder(drive_service,
                          folder_name,
                          parent_id=None):
    """Get or create Drive folder."""
    query = (
        f"name='{folder_name}' and "
        f"mimeType='application/vnd.google-apps.folder'"
        f" and trashed=false"
    )
    if parent_id:
        query += f" and '{parent_id}' in parents"

    results = drive_service.files().list(
        q=query, fields="files(id, name)"
    ).execute()
    files = results.get("files", [])

    if files:
        print(f"      📁 Exists: {folder_name}")
        return files[0]["id"]

    metadata = {
        "name"    : folder_name,
        "mimeType": "application/vnd.google-apps.folder"
    }
    if parent_id:
        metadata["parents"] = [parent_id]

    folder = drive_service.files().create(
        body=metadata, fields="id"
    ).execute()
    print(f"      📁 Created: {folder_name}")
    return folder.get("id")


def get_drive_folder(drive_service, run_date):
    """Get/create: CRD Report Automation/YYYY/Month/"""
    root_id  = get_or_create_folder(
        drive_service, DRIVE_ROOT_FOLDER
    )
    year_id  = get_or_create_folder(
        drive_service,
        run_date.strftime("%Y"),
        root_id
    )
    month_id = get_or_create_folder(
        drive_service,
        run_date.strftime("%B"),
        year_id
    )
    return month_id


def upload_to_drive(drive_service, file_path,
                    file_name, folder_id,
                    mime_type='application/octet-stream'):
    """Upload or update file in Drive folder."""
    query = (
        f"name='{file_name}' and "
        f"'{folder_id}' in parents and trashed=false"
    )
    results  = drive_service.files().list(
        q=query, fields="files(id)"
    ).execute()
    existing = results.get("files", [])

    media = MediaFileUpload(
        file_path,
        mimetype=mime_type,
        resumable=True
    )

    if existing:
        file_id = existing[0]["id"]
        drive_service.files().update(
            fileId=file_id, media_body=media
        ).execute()
        print(f"      ✅ Updated: {file_name}")
    else:
        result = drive_service.files().create(
            body={
                "name": file_name,
                "parents": [folder_id]
            },
            media_body=media,
            fields="id"
        ).execute()
        file_id = result.get("id")
        print(f"      ✅ Uploaded: {file_name}")

    return file_id


# ============================================================
#                    GOOGLE SHEETS
# ============================================================

def expand_sheet_rows(sheets_service,
                      spreadsheet_id,
                      sheet_name,
                      required_rows):
    """Expand sheet rows if grid is too small."""
    try:
        ss = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()

        sheet_id = None
        curr_rows = 0
        for s in ss.get('sheets', []):
            p = s.get('properties', {})
            if p.get('title') == sheet_name:
                sheet_id  = p.get('sheetId')
                curr_rows = p.get(
                    'gridProperties', {}
                ).get('rowCount', 1000)
                break

        if sheet_id is None:
            print(f"      ⚠️  Sheet not found: {sheet_name}")
            return

        needed = required_rows + 500
        if curr_rows < needed:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "rowCount": needed
                            }
                        },
                        "fields": "gridProperties.rowCount"
                    }
                }]}
            ).execute()
            print(
                f"      📐 Expanded: "
                f"{curr_rows:,} → {needed:,} rows"
            )
        else:
            print(
                f"      ✅ Rows OK: {curr_rows:,}"
            )
    except Exception as e:
        print(f"      ⚠️  Expand failed: {e}")


def push_to_sheet(sheets_service, spreadsheet_id,
                  sheet_name, df):
    """
    Clear and push DataFrame to Google Sheet.
    Uses batch writing of 5000 rows per call.
    """
    BATCH = 5000
    headers = df.columns.tolist()

    # Build clean rows
    data_rows = []
    for row in df.values.tolist():
        clean = []
        for val in row:
            try:
                is_nan = pd.isna(val)
            except Exception:
                is_nan = False

            if is_nan:
                clean.append('')
            elif isinstance(val, float):
                clean.append(round(val, 4))
            elif isinstance(val, int):
                clean.append(val)
            else:
                clean.append(str(val))
        data_rows.append(clean)

    total = len(data_rows)

    # Expand sheet first
    expand_sheet_rows(
        sheets_service, spreadsheet_id,
        sheet_name, total + 1
    )

    # Clear
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:ZZ"
    ).execute()
    print(f"      🗑️  Cleared: {sheet_name}")

    # Write header
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": [headers]}
    ).execute()
    print(f"      ✅ Header: {len(headers)} cols")

    # Write data in batches
    written = 0
    batch_n = 0
    while written < total:
        batch_n += 1
        batch     = data_rows[written:written + BATCH]
        start_row = written + 2

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A{start_row}",
            valueInputOption="RAW",
            body={"values": batch}
        ).execute()

        written += len(batch)
        pct = (written / total) * 100
        print(
            f"      📝 Batch {batch_n}: "
            f"{written:,}/{total:,} ({pct:.0f}%)"
        )

    print(f"      ✅ Done: {total:,} rows → {sheet_name}")


# ============================================================
#                 DATA FILTERING (NEW — CRD ONLY)
# ============================================================

def prepare_clean_df(df):
    """
    Filter the processed dataset down to the rows that
    belong in the Google Sheets:

        1. Vertical == 'CRD'          (case/space safe)
        2. Drop unmapped rows         (Not Found / blank)

    Returns (clean_df, total, crd_count, clean_count).
    """
    total = len(df)

    vertical = df['Vertical'].astype(str).str.strip().str.upper()
    crd_df   = df[vertical == 'CRD'].copy()
    crd_count = len(crd_df)

    clean_mask = (
        crd_df['Group Entity'].notna() &
        crd_df['Vertical'].notna() &
        (~crd_df['Group Entity'].astype(str).isin(
            ['Not Found', '', 'nan', 'None']
        )) &
        (~crd_df['Vertical'].astype(str).isin(
            ['Not Found', '', 'nan', 'None']
        ))
    )
    clean_df = crd_df[clean_mask].copy().reset_index(drop=True)

    return clean_df, total, crd_count, len(clean_df)


# ============================================================
#                       MAIN FUNCTION
# ============================================================

def main():
    print("\n" + "=" * 65)
    print("     📊 GOOGLE SHEETS PUSHER (CI/CD)")
    print("=" * 65)

    run_date      = datetime.now()
    timestamp_str = run_date.strftime("%d%m%Y_%H%M")

    # ── Step 1: Read State ────────────────────────────────────
    print("\n[1/6] Reading state file...")
    state, state_path = read_state_file()

    has_not_found     = state.get("has_not_found", False)
    output_dir        = state.get("output_dir", "")
    parquet_file_path = state.get("parquet_file_path", "")
    combined_path     = state.get("combined_file_path", "")

    print(f"      Has Not Found : {has_not_found}")
    print(f"      Output dir    : {output_dir}")
    print(f"      Parquet file  : {parquet_file_path}")

    if has_not_found:
        print(
            "\n      ⚠️  Not Found entities exist. "
            "Skipping push."
        )
        return

    # ── Find data file ────────────────────────────────────────
    if parquet_file_path and \
       os.path.exists(parquet_file_path):
        data_source = parquet_file_path
        print(f"      Data source   : parquet ✅")
    elif combined_path and os.path.exists(combined_path):
        data_source = combined_path
        print(f"      Data source   : excel ✅")
    else:
        # Search for any parquet in output_dir
        found = glob.glob(
            os.path.join(output_dir, "combined_data_*.parquet")
        )
        if found:
            # Use most recent
            data_source = max(found, key=os.path.getmtime)
            print(
                f"      Data source   : "
                f"{os.path.basename(data_source)} ✅"
            )
        else:
            print(
                "\n      ❌ No data file found. "
                "Cannot push to Sheets."
            )
            sys.exit(1)

    # ── Step 2: Load Data ────────────────────────────────────
    print(f"\n[2/6] Loading data from: {data_source}")

    if data_source.endswith('.parquet'):
        df = pd.read_parquet(
            data_source, engine='pyarrow'
        )
    else:
        df = pd.read_excel(
            data_source, engine='openpyxl'
        )

    clean_df, total, crd_count, clean_count = \
        prepare_clean_df(df)

    print(f"      Total rows     : {total:,}")
    print(f"      CRD rows       : {crd_count:,}")
    print(f"      Pushable rows  : {clean_count:,} (CRD, mapped)")

    if clean_count == 0:
        print(
            "\n      ❌ No CRD rows to push. Aborting."
        )
        sys.exit(1)

    # ── Step 3: Authenticate ─────────────────────────────────
    print("\n[3/6] Authenticating Google APIs...")
    credentials    = get_google_credentials()
    sheets_service = build(
        'sheets', 'v4', credentials=credentials
    )
    drive_service  = build(
        'drive', 'v3', credentials=credentials
    )
    print("      ✅ Sheets + Drive APIs ready.")

    # ── Step 4: Drive Upload Skipped ─────────────────────────
    print("\n[4/6] Skipping Drive upload.")
    print(
        "      Files available via "
        "GitHub Actions Artifacts."
    )

    # ── Step 5: Push to Google Sheets (CRD only) ─────────────
    print("\n[5/6] Pushing CRD data to Google Sheets...")

    # Split by calendar year
    clean_df['_report_date'] = pd.to_datetime(
        clean_df['ReportingTime'],
        format='%d-%b-%Y',
        errors='coerce'
    )
    clean_df['_year'] = clean_df['_report_date'].dt.year

    push_results = {}

    for year in sorted(
        clean_df['_year'].dropna().unique().tolist()
    ):
        year     = int(year)
        year_df  = clean_df[
            clean_df['_year'] == year
        ].copy()

        print(f"\n      ── Year {year} ──")
        print(f"      Rows: {len(year_df):,}")

        if year not in SHEET_CONFIG:
            print(
                f"      ❌ No sheet config for {year}. "
                f"Add it to SHEET_CONFIG in push_to_sheets.py."
            )
            push_results[year] = "NO_CONFIG"
            continue

        cfg = SHEET_CONFIG[year]

        # Only required 12 columns
        cols_available = [
            c for c in SHEETS_COLUMNS
            if c in year_df.columns
        ]
        sheets_df = year_df[cols_available].copy()

        try:
            push_to_sheet(
                sheets_service = sheets_service,
                spreadsheet_id = cfg["spreadsheet_id"],
                sheet_name     = cfg["sheet_name"],
                df             = sheets_df
            )
            push_results[year] = "SUCCESS"
        except Exception as e:
            print(f"      ❌ Failed: {e}")
            push_results[year] = f"FAILED: {e}"

    # ── Step 6: Evaluate results ─────────────────────────────
    print("\n[6/6] Evaluating results...")

    failed_years = {
        yr: res
        for yr, res in push_results.items()
        if res != "SUCCESS"
    }

    print("\n" + "=" * 65)
    print("        📊 SHEETS PUSH SUMMARY")
    print("=" * 65)
    for yr, result in push_results.items():
        icon = "✅" if result == "SUCCESS" else "❌"
        print(f"    {icon} {yr} : {result}")
    print("=" * 65)

    # ── HARD FAIL if anything did not push ───────────────────
    # sheets_pushed stays False, state is NOT committed,
    # README link is NOT written, workflow shows red,
    # and the next trigger retries the push.
    if not push_results or failed_years:
        print(
            "\n❌ Sheets push INCOMPLETE — "
            f"{len(failed_years)} year(s) failed "
            "or missing config."
        )
        print(
            "   sheets_pushed NOT set. "
            "Fix the issue and re-run."
        )
        if any(
            res == "NO_CONFIG"
            for res in failed_years.values()
        ):
            print(
                "   ⚠️  A calendar year has no target "
                "spreadsheet in SHEET_CONFIG — add the "
                "new year's spreadsheet ID."
            )
        print(
            "   💡 If this is a 403 permission error: share "
            "the spreadsheet(s) with the service account "
            "email (Editor role)."
        )
        sys.exit(1)

    # ── Success: cleanup + state update ──────────────────────
    # Cleanup old intermediate files
    for pattern in [
        "Combined_Invoice_*.xlsx",
        "Sample_Invoice_*.xlsx"
    ]:
        for f in glob.glob(
            os.path.join(output_dir, pattern)
        ):
            os.remove(f)
            print(
                f"      🗑️  Deleted: {os.path.basename(f)}"
            )

    update_state_file(state_path, {
        "sheets_push_results": push_results,
        "sheets_pushed"      : True,
        "sheets_push_time"   : run_date.strftime(
            "%d-%b-%Y %H:%M"
        )
    })

    print("\n" + "=" * 65)
    print("        ✅ SHEETS PUSH COMPLETE — ALL YEARS OK")
    print("=" * 65)


if __name__ == "__main__":
    main()
