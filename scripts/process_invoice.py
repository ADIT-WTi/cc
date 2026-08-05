# ============================================================
#         INVOICE DATA PROCESSOR - GITHUB ACTIONS VERSION
# ============================================================
# What this script does:
#   1. Reads downloaded Excel files from downloads/ folder
#   2. Combines all monthly files into one combined file
#   3. Applies all transformations (Month, Group Hub, Revenue etc.)
#   4. Assigns Group Entity & Vertical from master file
#   5. Identifies Not Found rows
#   6. Generates Not Found summary Excel if needed
#   7. Saves processed files to reports/YYYY/MonthName/
#   8. Writes state flags for next script (generate_report.py)
# ============================================================

import pandas as pd
import numpy as np
import glob
import os
import json
from datetime import datetime
import calendar

# ============================================================
#                    PATH CONFIGURATION
# ============================================================

def get_repo_root():
    """Returns the root directory of the repository."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_current_run_dir():
    """
    Reads the download directory path saved by
    download_invoice.py.

    Handles both relative and absolute paths.
    Relative paths are resolved from repo root
    so they work across different runners.
    """
    repo_root  = get_repo_root()

    # ── Try .current_run_dir.txt first ────────────────────
    state_file = os.path.join(
        repo_root, "reports", ".current_run_dir.txt"
    )

    if os.path.exists(state_file):
        with open(state_file, "r") as f:
            saved_path = f.read().strip()

        # If relative path → resolve from repo root
        if not os.path.isabs(saved_path):
            download_dir = os.path.join(
                repo_root, saved_path
            )
        else:
            # Absolute path from old run — extract
            # relative part and rebuild with current root
            # Example:
            # /home/runner/work/.../reports/2026/July/downloads
            # → reports/2026/July/downloads
            try:
                # Find 'reports/' in the path
                idx = saved_path.find('reports' + os.sep)
                if idx == -1:
                    idx = saved_path.find('reports/')
                if idx != -1:
                    relative_part = saved_path[idx:]
                    download_dir  = os.path.join(
                        repo_root, relative_part
                    )
                else:
                    download_dir = saved_path
            except Exception:
                download_dir = saved_path

        if os.path.exists(download_dir):
            print(
                f"      Using download dir from "
                f".current_run_dir.txt"
            )
            print(
                f"      Resolved path: {download_dir}"
            )
            return download_dir
        else:
            print(
                f"      ⚠️  Path from txt not found: "
                f"{download_dir}"
            )
            print("      Trying .run_state.json...")

    # ── Fallback: read from .run_state.json ───────────────
    run_state_path = os.path.join(
        repo_root, "reports", ".run_state.json"
    )

    if os.path.exists(run_state_path):
        with open(run_state_path, "r") as f:
            state = json.load(f)

        output_dir = state.get("output_dir", "")

        if output_dir:
            # Handle absolute path from different runner
            if not os.path.isabs(output_dir):
                abs_output = os.path.join(
                    repo_root, output_dir
                )
            else:
                # Extract relative part
                idx = output_dir.find('reports' + os.sep)
                if idx == -1:
                    idx = output_dir.find('reports/')
                if idx != -1:
                    rel_part   = output_dir[idx:]
                    abs_output = os.path.join(
                        repo_root, rel_part
                    )
                else:
                    abs_output = output_dir

            download_dir = os.path.join(
                abs_output, "downloads"
            )

            if os.path.exists(download_dir):
                print(
                    f"      Using download dir from "
                    f".run_state.json"
                )
                print(
                    f"      Resolved path: {download_dir}"
                )
                return download_dir
            else:
                print(
                    f"      ⚠️  downloads/ not found: "
                    f"{download_dir}"
                )

    raise FileNotFoundError(
        "❌ Could not find download directory.\n"
        "   Checked: .current_run_dir.txt and "
        ".run_state.json\n"
        "   The downloads/ folder may not exist in "
        "this runner.\n"
        "   Note: Downloaded files are not committed "
        "to git (excluded by .gitignore)."
    )


def get_run_output_dir(download_dir):
    """
    Returns the parent folder of downloads/ as the output directory.

    Example:
        download_dir = reports/2026/August/downloads/
        output_dir   = reports/2026/August/
    """
    return os.path.dirname(download_dir)


def get_master_file_path():
    """Returns path to the Group Entity & Vertical master file."""
    repo_root = get_repo_root()
    return os.path.join(
        repo_root, "data", "Group_Entity_Vertical_Master.xlsx"
    )


# ============================================================
#                    HELPER FUNCTIONS
# ============================================================

def extract_date_from_filename(file):
    """
    Extract month-year from filename like:
    InvoiceReport_702_Jan_2026_13072026_1231.xlsx
    Returns datetime for sorting.
    """
    base  = os.path.basename(file)
    parts = base.split("_")
    try:
        month_str = parts[2]   # Jan
        year_str  = parts[3]   # 2026
        return datetime.strptime(f"{month_str} {year_str}", "%b %Y")
    except Exception:
        return datetime.min


def add_month_column(df):
    """
    Add 'Month' column right after 'ReportingTime'.
    Format: APR-2026 (uppercase)
    """
    reporting_dates = pd.to_datetime(
        df['ReportingTime'], format='%d-%b-%Y', errors='coerce'
    )
    month_values = (
        reporting_dates
        .dt.strftime('%b-%Y')
        .str.upper()
    )

    insert_loc = df.columns.get_loc('ReportingTime') + 1
    df.insert(insert_loc, 'Month', month_values)
    return df


def add_group_hub_column(df):
    """
    Add 'Group Hub' column right after 'Hub'.
    Maps individual hubs to group hub names.
    """
    group_hub_map = {
        'Noida'          : 'Delhi NCR',
        'Noida Branch'   : 'Delhi NCR',
        'Noida ETS'      : 'Delhi NCR',
        'Manesar Hub'    : 'Delhi NCR',
        'ETS'            : 'Delhi NCR',
        'Gurugram'       : 'Delhi NCR',
        'New Delhi'      : 'Delhi NCR',
        'Kolkata Hub'    : 'Kolkata',
        'Chennai Hub'    : 'Chennai',
        'BLR Hub'        : 'Bengaluru',
        'BLR-Accenture'  : 'Bengaluru',
        'BLR-Domlur'     : 'Bengaluru',
        'BLR-Wipro'      : 'Bengaluru',
        'Hyderabad Hub'  : 'Hyderabad',
        'Pune Hub 1'     : 'Pune',
        'Pune Hub 2'     : 'Pune',
        'Mumbai Hub'     : 'Mumbai'
    }

    insert_loc = df.columns.get_loc('Hub') + 1
    df.insert(
        insert_loc,
        'Group Hub',
        df['Hub'].map(group_hub_map).fillna('Tier 2/3')
    )
    return df


def apply_exclusions(df):
    """
    Remove rows with excluded CorporateIDs.
    Remove rows where KoAmount < 1.
    Excluded rows are simply dropped — not saved anywhere.
    """
    exclude_corporate_ids = {
        '-1', '0', '1', '4109', '4297', '4299', '1590', '3170',
        '4135', '1847', '3384', '4298', '3816', '195', '4227',
        '4893', '4392', '4700', '3500', '4315', '3278', '4748',
        '3097', '3979', '774', '4489', '4462', '4818', '4831',
        '4817', '4776', '4490', '4961', '4995'
    }

    before = len(df)

    # Remove excluded CorporateIDs
    df = df[
        ~df['CorporateID'].astype(str).str.strip().isin(
            exclude_corporate_ids
        )
    ]

    after_corp = len(df)

    # Remove KoAmount < 1 (unchanged as requested)
    df = df[
        pd.to_numeric(df['KoAmount'], errors='coerce').fillna(0) >= 1
    ]

    after_ko = len(df)

    # Reset index after filtering
    df = df.reset_index(drop=True)

    print(
        f"      Rows removed by CorporateID filter : "
        f"{before - after_corp:,}"
    )
    print(
        f"      Rows removed by KoAmount < 1 filter: "
        f"{after_corp - after_ko:,}"
    )
    print(
        f"      Total rows removed                 : "
        f"{before - after_ko:,}"
    )

    return df


def calculate_revenue(df):
    """
    Calculate New GST and Revenue columns.

    New GST = TotalTax / TotalAmount
              → NaN if TotalAmount = 0 or NaN (DIV/0 scenario)
              → inf/-inf replaced with NaN

    Revenue Logic (SINGLE BLOCK — no overwrite):
    ─────────────────────────────────────────────
    Rule 1 (mask_div0)   : TotalAmount = 0 or NaN
                           → Revenue = KoAmount
                           [ETS / Reimbursement rows]

    Rule 2 (mask_formula): TotalAmount > 0
                           AND CrAmount ≠ 0
                           AND New GST is not NaN
                           AND New GST is not 0
                           → Revenue = KoAmount / (1 + New GST)
                           [Credit note exists, strip GST]

    Rule 3 (default)     : Everything else
                           → Revenue = TotalAmount
                           [Normal invoice, no credit note]
    """
    # ── Numeric conversions (done once) ───────────────────
    total_tax_raw    = pd.to_numeric(
        df['TotalTax'],    errors='coerce'
    )
    total_amount_raw = pd.to_numeric(
        df['TotalAmount'], errors='coerce'
    )
    ko_amount        = pd.to_numeric(
        df['KoAmount'],    errors='coerce'
    ).fillna(0)
    cr_amount        = pd.to_numeric(
        df['CrAmount'],    errors='coerce'
    ).fillna(0)

    # ── DIV/0 mask ────────────────────────────────────────
    # Rows where TotalAmount = 0 or NaN
    # (ETS / Reimbursement rows)
    mask_div0 = (
        total_amount_raw.eq(0) | total_amount_raw.isna()
    )

    # ── New GST calculation ───────────────────────────────
    # NaN where TotalAmount = 0 or NaN
    # inf/-inf replaced with NaN for safety
    new_gst = (
        total_tax_raw / total_amount_raw
    ).replace([np.inf, -np.inf], np.nan)

    # Insert New GST right after CrAmount
    # Use .values to avoid index alignment issues
    if 'New GST' in df.columns:
        df['New GST'] = new_gst.values
    else:
        insert_loc = df.columns.get_loc('CrAmount') + 1
        df.insert(insert_loc, 'New GST', new_gst.values)

    # ── Revenue calculation (SINGLE BLOCK) ───────────────
    total_amount = total_amount_raw.fillna(0)

    # Step A: Default = TotalAmount (covers Rule 3)
    revenue = total_amount.copy()

    # Step B: Rule 1 override
    # TotalAmount = 0/NaN → Revenue = KoAmount
    # Handles ETS / Reimbursement rows correctly
    revenue.loc[mask_div0] = ko_amount.loc[mask_div0]

    # Step C: Rule 2 mask (fully guarded)
    # All four conditions must be True
    mask_formula = (
        (~mask_div0)      &   # Not a DIV/0 row
        (cr_amount != 0)  &   # Credit note exists
        new_gst.notna()   &   # GST is calculable (NaN guard)
        (new_gst != 0)        # GST is not zero (zero guard)
    )

    # Step D: Rule 2 override
    # CrAmount ≠ 0 and valid GST → strip GST from KoAmount
    revenue.loc[mask_formula] = (
        ko_amount.loc[mask_formula] /
        (1 + new_gst.loc[mask_formula])
    ).round(1)

    # Final rounding
    revenue = revenue.round(1)

    # Insert Revenue right after New GST
    # Use .values to avoid index alignment issues
    if 'Revenue' in df.columns:
        df['Revenue'] = revenue.values
    else:
        insert_loc = df.columns.get_loc('New GST') + 1
        df.insert(insert_loc, 'Revenue', revenue.values)

    # ── Revenue sanity print ──────────────────────────────
    print(
        f"      Rule 1 rows (DIV/0 → KoAmount)    : "
        f"{mask_div0.sum():,}"
    )
    print(
        f"      Rule 2 rows (formula → KoAmt/GST) : "
        f"{mask_formula.sum():,}"
    )
    print(
        f"      Rule 3 rows (default → TotalAmt)  : "
        f"{(~mask_div0 & ~mask_formula).sum():,}"
    )
    print(
        f"      Revenue NaN count                  : "
        f"{pd.Series(revenue.values).isna().sum():,}"
    )
    print(
        f"      Total Revenue                      : "
        f"₹{revenue.sum():,.1f}"
    )

    return df


def assign_group_entity_vertical(df, master_file_path):
    """
    Assigns 'Group Entity' and 'Vertical' columns from master file.

    Lookup key = Corporate Name + || + Hub

    Returns:
        df           : DataFrame with Group Entity & Vertical assigned
        not_found_df : Rows where Group Entity or Vertical = 'Not Found'
        found_df     : Rows where both are found
    """
    print(f"\n      Loading master file: {master_file_path}")

    # Determine engine based on file extension
    master_ext    = os.path.splitext(master_file_path)[1].lower()
    master_engine = 'openpyxl' if master_ext == '.xlsx' else 'xlrd'

    client_master = pd.read_excel(
        master_file_path,
        sheet_name='Sheet1',
        engine=master_engine
    )

    # Build lookup dictionaries
    master_key = (
        client_master['Corporate Name']
        .fillna('').astype(str).str.strip()
        + '||' +
        client_master['Hub']
        .fillna('').astype(str).str.strip()
    )
    group_entity_lookup = dict(
        zip(master_key, client_master['GroupEntity'])
    )
    vertical_lookup = dict(
        zip(master_key, client_master['Vertical'])
    )

    # Build key for main dataframe
    df_key = (
        df['Corporate'].fillna('').astype(str).str.strip()
        + '||' +
        df['Hub'].fillna('').astype(str).str.strip()
    )

    group_entity_values = df_key.map(
        group_entity_lookup
    ).fillna('Not Found')

    vertical_values = df_key.map(
        vertical_lookup
    ).fillna('Not Found')

    # Drop old columns if exist
    for col in ['Group Entity', 'Vertical']:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Insert after Corporate column
    insert_loc = df.columns.get_loc('Corporate') + 1
    df.insert(insert_loc,     'Group Entity', group_entity_values)
    df.insert(insert_loc + 1, 'Vertical',     vertical_values)

    # Separate not found rows
    not_found_mask = (
        (df['Group Entity'] == 'Not Found') |
        (df['Vertical']     == 'Not Found')
    )

    not_found_df = df[not_found_mask].copy()
    found_df     = df[~not_found_mask].copy()

    print(f"      Total rows       : {len(df):,}")
    print(f"      Mapped rows      : {len(found_df):,}")
    print(f"      Not Found rows   : {len(not_found_df):,}")

    return df, not_found_df, found_df


def generate_not_found_summary(not_found_df):
    """
    Aggregate Not Found rows into summary format.

    Output columns:
        Month | CorporateID | Corporate | Hub | Group Hub |
        Group Entity | Vertical | Bookings Count | Revenue

    Group Entity and Vertical are LEFT BLANK for user to fill.
    """
    if not_found_df.empty:
        return pd.DataFrame()

    not_found_df = not_found_df.copy()

    not_found_df['Revenue'] = pd.to_numeric(
        not_found_df['Revenue'], errors='coerce'
    ).fillna(0)

    not_found_df['BookingNo'] = (
        not_found_df['BookingNo'].astype(str)
    )

    # Aggregate
    summary = (
        not_found_df
        .groupby(
            [
                'Month', 'CorporateID', 'Corporate',
                'Hub', 'Group Hub'
            ],
            as_index=False,
            sort=False
        )
        .agg(
            **{
                'Bookings Count': ('BookingNo', 'nunique'),
                'Revenue'       : ('Revenue',   'sum')
            }
        )
    )

    # Add blank columns for user to fill
    summary['Group Entity'] = ''
    summary['Vertical']     = ''

    # Sort by Month then Revenue descending
    summary['_sort_key'] = pd.to_datetime(
        summary['Month'],
        format='%b-%Y',
        errors='coerce'
    )
    summary = summary.sort_values(
        ['_sort_key', 'Revenue'],
        ascending=[True, False]
    ).drop(columns=['_sort_key'])

    summary['Revenue'] = (
        summary['Revenue'].round(0).astype(int)
    )

    # Final column order
    summary = summary[[
        'Month', 'CorporateID', 'Corporate',
        'Hub', 'Group Hub',
        'Group Entity', 'Vertical',
        'Bookings Count', 'Revenue'
    ]].reset_index(drop=True)

    return summary


def save_not_found_summary(summary_df, output_dir):
    """
    Save Not Found summary to Excel file AND
    also save as JSON for reliable reading by
    generate_report.py.

    Returns the Excel file path.
    """
    import json

    excel_path = os.path.join(
        output_dir, "Not_Found_Summary.xlsx"
    )
    json_path  = os.path.join(
        output_dir, "not_found_data.json"
    )

    # ── Save as JSON (primary data store) ─────────────────
    records = []
    for _, row in summary_df.iterrows():
        record = {}
        for col in summary_df.columns:
            val = row[col]
            if hasattr(val, 'item'):
                val = val.item()
            record[col] = str(val) if val != '' else ''
        records.append(record)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print(
        f"\n      💾 Not Found JSON saved: {json_path}"
        f" ({len(records)} records)"
    )

    # ── Save as formatted Excel ────────────────────────────
    summary_df.to_excel(
        excel_path,
        sheet_name='Not Found Summary',
        index=False,
        engine='openpyxl'
    )

    print(
        f"      💾 Not Found Excel saved: {excel_path}"
    )

    return excel_path, json_path


def save_combined_file(df, output_dir, timestamp_str):
    """
    Save combined processed invoice data.
    Saves as PARQUET (compressed, small size) for
    cross-workflow use AND as Excel for reference.
    """
    df_parquet = df.copy()

    for col in df_parquet.columns:
        col_dtype = df_parquet[col].dtype

        if col_dtype == object:
            df_parquet[col] = (
                df_parquet[col]
                .astype(str)
                .replace('nan', '')
                .replace('<NA>', '')
            )

        elif str(col_dtype) in [
            'Int8', 'Int16', 'Int32', 'Int64',
            'UInt8', 'UInt16', 'UInt32', 'UInt64'
        ]:
            df_parquet[col] = (
                df_parquet[col]
                .fillna(0)
                .astype('int64')
            )

        elif str(col_dtype) in ['Float32', 'Float64']:
            df_parquet[col] = (
                df_parquet[col]
                .fillna(0.0)
                .astype('float64')
            )

    parquet_path = os.path.join(
        output_dir,
        f"combined_data_{timestamp_str}.parquet"
    )
    df_parquet.to_parquet(
        parquet_path,
        index=False,
        engine='pyarrow',
        compression='gzip'
    )

    size_mb = os.path.getsize(parquet_path) / 1024 / 1024
    print(f"\n      💾 Parquet saved: {parquet_path}")
    print(f"      📦 Size: {size_mb:.1f} MB")

    excel_path = os.path.join(
        output_dir,
        f"Combined_Invoice_{timestamp_str}.xlsx"
    )
    df.to_excel(excel_path, index=False)
    print(f"      💾 Excel saved : {excel_path}")

    return parquet_path, excel_path


def save_sample_invoice(df, output_dir, timestamp_str):
    """
    Save the sample invoice (processed data with highlights
    on Group Entity and Vertical columns).
    """
    file_path = os.path.join(
        output_dir,
        f"Sample_Invoice_{timestamp_str}.xlsx"
    )

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        highlight_fmt = workbook.add_format(
            {'bg_color': '#FFF2CC'}
        )

        group_entity_col_idx = df.columns.get_loc('Group Entity')
        vertical_col_idx     = df.columns.get_loc('Vertical')

        worksheet.set_column(
            group_entity_col_idx, group_entity_col_idx,
            None, highlight_fmt
        )
        worksheet.set_column(
            vertical_col_idx, vertical_col_idx,
            None, highlight_fmt
        )

    print(f"      💾 Sample Invoice saved: {file_path}")
    return file_path


def write_state_file(
    output_dir,
    has_not_found,
    not_found_summary_path,
    not_found_json_path,
    combined_file_path,
    parquet_file_path,
    sample_invoice_path,
    run_date_str,
    start_month_str,
    end_month_str
):
    """
    Write state JSON file with relative paths.
    Preserves existing keys set by update_master.py.
    """
    repo_root  = get_repo_root()
    state_path = os.path.join(
        repo_root, "reports", ".run_state.json"
    )

    def to_relative(path):
        if not path:
            return path
        try:
            return os.path.relpath(path, repo_root)
        except ValueError:
            return path

    # Read existing state to preserve flags
    existing_state = {}
    if os.path.exists(state_path):
        try:
            with open(state_path, "r") as f:
                existing_state = json.load(f)
        except Exception:
            existing_state = {}

    state = {
        "output_dir"             : to_relative(output_dir),
        "has_not_found"          : has_not_found,
        "not_found_summary_path" : to_relative(
            not_found_summary_path
        ),
        "not_found_json_path"    : to_relative(
            not_found_json_path
        ),
        "combined_file_path"     : to_relative(
            combined_file_path
        ),
        "parquet_file_path"      : to_relative(
            parquet_file_path
        ),
        "sample_invoice_path"    : to_relative(
            sample_invoice_path
        ),
        "run_date"               : run_date_str,
        "start_month"            : start_month_str,
        "end_month"              : end_month_str,
        "final_report_path"      : "",
        "email2_sent"            : False,
        "master_updated"         : existing_state.get(
            "master_updated", False
        ),
        "master_update_time"     : existing_state.get(
            "master_update_time", ""
        ),
        "rows_added_to_master"   : existing_state.get(
            "rows_added_to_master", 0
        ),
        "email1_sent"            : existing_state.get(
            "email1_sent", False
        ),
    }

    with open(state_path, "w") as f:
        json.dump(state, f, indent=4)

    print(f"\n      📝 State file saved: {state_path}")
    print(
        f"      master_updated preserved: "
        f"{state['master_updated']}"
    )
    return state_path


# ============================================================
#                       MAIN FUNCTION
# ============================================================

def main():
    run_date      = datetime.now()
    timestamp_str = run_date.strftime("%d%m%Y_%H%M")
    run_date_str  = run_date.strftime("%d-%b-%Y")

    print("\n" + "=" * 65)
    print("         📊 INVOICE DATA PROCESSOR (CI/CD)")
    print("=" * 65)

    # ── Step 1: Get Paths ────────────────────────────────────
    print("\n[1/8] Resolving paths...")
    download_dir = get_current_run_dir()
    output_dir   = get_run_output_dir(download_dir)
    master_path  = get_master_file_path()

    print(f"      Download dir : {download_dir}")
    print(f"      Output dir   : {output_dir}")
    print(f"      Master file  : {master_path}")

    # ── Step 2: Load All Downloaded Files ───────────────────
    print("\n[2/8] Loading downloaded files...")
    file_paths = glob.glob(
        os.path.join(download_dir, "*.xlsx")
    )

    if not file_paths:
        raise FileNotFoundError(
            f"❌ No Excel files found in: {download_dir}"
        )

    file_paths_sorted = sorted(
        file_paths, key=extract_date_from_filename
    )
    print(f"      Found {len(file_paths_sorted)} file(s):")
    for f in file_paths_sorted:
        print(f"        • {os.path.basename(f)}")

    # ── Step 3: Combine Files ────────────────────────────────
    print("\n[3/8] Combining files...")
    dfs = []
    for file in file_paths_sorted:
        file_ext    = os.path.splitext(file)[1].lower()
        file_engine = 'openpyxl' if file_ext == '.xlsx' else 'xlrd'

        df_temp = pd.read_excel(file, engine=file_engine)
        df_temp.dropna(how='all', inplace=True)
        dfs.append(df_temp)

    combined_df = pd.concat(dfs, ignore_index=True)
    print(
        f"      Total rows after combining: "
        f"{len(combined_df):,}"
    )

    # ── Step 4: Apply Transformations ───────────────────────
    print("\n[4/8] Applying transformations...")
    combined_df = add_month_column(combined_df)
    combined_df = add_group_hub_column(combined_df)
    combined_df = apply_exclusions(combined_df)

    print("\n      Calculating Revenue...")
    combined_df = calculate_revenue(combined_df)

    print(
        f"\n      Rows after transformation: "
        f"{len(combined_df):,}"
    )

    # ── Step 5: Assign Group Entity & Vertical ───────────────
    print("\n[5/8] Assigning Group Entity & Vertical...")
    combined_df, not_found_df, found_df = \
        assign_group_entity_vertical(combined_df, master_path)

    # ── Step 5b: Filter CRD Vertical Only ───────────────────
    print("\n[5b/8] Filtering CRD vertical only...")

    crd_mask = (
        combined_df['Vertical']
        .str.upper()
        .str.strip() == 'CRD'
    )

    not_found_mask = (
        (combined_df['Group Entity'] == 'Not Found') |
        (combined_df['Vertical']     == 'Not Found')
    )

    crd_found_df = combined_df[
        crd_mask & ~not_found_mask
    ].copy()

    not_found_df = combined_df[not_found_mask].copy()

    print(f"      CRD mapped rows    : {len(crd_found_df):,}")
    print(f"      Not Found rows     : {len(not_found_df):,}")
    print(
        f"      Other vertical rows: "
        f"{len(combined_df) - len(crd_found_df) - len(not_found_df):,}"
        f" (excluded)"
    )

    # ── Step 6: Save Files ───────────────────────────────────
    print("\n[6/8] Saving processed files...")

    parquet_path, combined_file_path = save_combined_file(
        combined_df, output_dir, timestamp_str
    )

    sample_invoice_path = save_sample_invoice(
        combined_df, output_dir, timestamp_str
    )

    # ── Step 7: Handle Not Found ─────────────────────────────
    print("\n[7/8] Handling Not Found entities...")
    has_not_found          = not not_found_df.empty
    not_found_summary_path = ""
    not_found_json_path    = ""

    if has_not_found:
        print(
            f"      ⚠️  {len(not_found_df):,} rows with "
            f"Not Found Group Entity / Vertical"
        )
        summary_df = generate_not_found_summary(not_found_df)

        print(
            f"\n      Not Found Summary preview "
            f"({len(summary_df)} rows):"
        )
        print(f"      Columns: {list(summary_df.columns)}")
        for i, row in summary_df.head(3).iterrows():
            print(
                f"      Row {i}: "
                f"{row['Corporate']} | "
                f"{row['Hub']} | "
                f"Bookings={row['Bookings Count']} | "
                f"Revenue={row['Revenue']}"
            )

        not_found_summary_path, not_found_json_path = \
            save_not_found_summary(summary_df, output_dir)

    else:
        print(
            "      ✅ All rows mapped. "
            "No Not Found entities."
        )

    # ── Step 8: Write State File ─────────────────────────────
    print("\n[8/8] Writing state file...")

    months_in_data = (
        combined_df['Month'].dropna().unique().tolist()
    )

    def month_sort_key(m):
        try:
            return pd.Period(m, freq='M')
        except Exception:
            return pd.Period('1900-01', freq='M')

    months_sorted = sorted(
        [
            m for m in months_in_data
            if m not in ('', 'nan')
        ],
        key=month_sort_key
    )
    start_month_str = months_sorted[0]  if months_sorted else ""
    end_month_str   = months_sorted[-1] if months_sorted else ""

    state_path = write_state_file(
        output_dir             = output_dir,
        has_not_found          = has_not_found,
        not_found_summary_path = not_found_summary_path,
        not_found_json_path    = not_found_json_path,
        combined_file_path     = combined_file_path,
        parquet_file_path      = parquet_path,
        sample_invoice_path    = sample_invoice_path,
        run_date_str           = run_date_str,
        start_month_str        = start_month_str,
        end_month_str          = end_month_str
    )

    # ── Final Summary ────────────────────────────────────────
    print("\n" + "=" * 65)
    print("                  📊 PROCESSING SUMMARY")
    print("=" * 65)
    print(f"  Total rows processed : {len(combined_df):,}")
    print(f"  Mapped rows          : {len(found_df):,}")
    print(f"  Not Found rows       : {len(not_found_df):,}")
    print(
        f"  Date range           : "
        f"{start_month_str} → {end_month_str}"
    )
    print(
        f"  Has Not Found        : "
        f"{'YES ⚠️' if has_not_found else 'NO ✅'}"
    )
    print(f"  Output directory     : {output_dir}")
    print("=" * 65)


if __name__ == "__main__":
    main()
