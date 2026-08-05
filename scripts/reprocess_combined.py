# ============================================================
#     REPROCESS COMBINED INVOICE - FOLLOWUP WORKFLOW
# ============================================================
# Reads parquet file (small, committed to git)
# Re-assigns Group Entity + Vertical from updated master
# Does NOT need downloads/ folder
# ============================================================

import pandas as pd
import numpy as np
import os
import json
from datetime import datetime


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
        "final_report_path"
    ]:
        if state.get(key):
            state[key] = resolve_path(state[key])

    return state, state_path


def update_state_file(state_path, updates):
    """Update state file preserving relative paths."""
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


def assign_group_entity_vertical(df, master_path):
    """Re-assign Group Entity and Vertical from master."""
    print(f"\n      Loading master: {master_path}")

    client_master = pd.read_excel(
        master_path,
        sheet_name='Sheet1',
        engine='openpyxl'
    )

    master_key = (
        client_master['Corporate Name']
        .fillna('').astype(str).str.strip()
        + '||' +
        client_master['Hub']
        .fillna('').astype(str).str.strip()
    )
    ge_lookup = dict(
        zip(master_key, client_master['GroupEntity'])
    )
    vt_lookup = dict(
        zip(master_key, client_master['Vertical'])
    )

    df_key = (
        df['Corporate'].fillna('').astype(str).str.strip()
        + '||' +
        df['Hub'].fillna('').astype(str).str.strip()
    )

    # Update only Not Found rows
    not_found_mask = (
        (df['Group Entity'] == 'Not Found') |
        (df['Vertical']     == 'Not Found')
    )

    df.loc[not_found_mask, 'Group Entity'] = (
        df_key[not_found_mask]
        .map(ge_lookup).fillna('Not Found')
    )
    df.loc[not_found_mask, 'Vertical'] = (
        df_key[not_found_mask]
        .map(vt_lookup).fillna('Not Found')
    )

    still_nf = (
        (df['Group Entity'] == 'Not Found') |
        (df['Vertical']     == 'Not Found')
    )

    print(f"      Total rows      : {len(df):,}")
    print(f"      Now mapped      : {(~still_nf).sum():,}")
    print(f"      Still Not Found : {still_nf.sum():,}")

    return df, df[still_nf].copy(), df[~still_nf].copy()


def generate_not_found_summary(not_found_df):
    """Generate Not Found summary if rows still exist."""
    if not_found_df.empty:
        return pd.DataFrame()

    nf = not_found_df.copy()
    nf['Revenue']   = pd.to_numeric(
        nf['Revenue'], errors='coerce'
    ).fillna(0)
    nf['BookingNo'] = nf['BookingNo'].astype(str)

    summary = (
        nf.groupby(
            ['Month', 'CorporateID', 'Corporate',
             'Hub', 'Group Hub'],
            as_index=False, sort=False
        )
        .agg(**{
            'Bookings Count': ('BookingNo', 'nunique'),
            'Revenue'       : ('Revenue',   'sum')
        })
    )
    summary['Group Entity'] = ''
    summary['Vertical']     = ''
    summary['_sort']        = pd.to_datetime(
        summary['Month'], format='%b-%Y', errors='coerce'
    )
    summary = summary.sort_values(
        ['_sort', 'Revenue'], ascending=[True, False]
    ).drop(columns=['_sort'])
    summary['Revenue'] = (
        summary['Revenue'].round(0).astype(int)
    )
    return summary[[
        'Month', 'CorporateID', 'Corporate',
        'Hub', 'Group Hub', 'Group Entity', 'Vertical',
        'Bookings Count', 'Revenue'
    ]].reset_index(drop=True)


def main():
    print("\n" + "=" * 65)
    print("     🔄 REPROCESS COMBINED INVOICE (CI/CD)")
    print("=" * 65)

    repo_root     = get_repo_root()
    timestamp_str = datetime.now().strftime("%d%m%Y_%H%M")

    # ── Step 1: Read State ───────────────────────────────────
    print("\n[1/6] Reading state file...")
    state, state_path = read_state_file()

    parquet_file_path = state.get("parquet_file_path", "")
    output_dir        = state.get("output_dir", "")
    start_month       = state.get("start_month", "")
    end_month         = state.get("end_month", "")

    print(f"      Parquet file : {parquet_file_path}")
    print(f"      Output dir   : {output_dir}")

    # ── Step 2: Validate Parquet File ───────────────────────
    print("\n[2/6] Validating parquet file...")

    if not parquet_file_path or \
       not os.path.exists(parquet_file_path):
        raise FileNotFoundError(
            f"❌ Parquet file not found:\n"
            f"   {parquet_file_path}\n"
            f"   Ensure combined_data_*.parquet is "
            f"committed to git from Workflow 1.\n"
            f"   Check reports/.gitignore allows "
            f"*.parquet files."
        )

    size_mb = (
        os.path.getsize(parquet_file_path)
        / 1024 / 1024
    )
    print(
        f"      ✅ Found: {parquet_file_path} "
        f"({size_mb:.1f} MB)"
    )

    # ── Step 3: Load Parquet File ────────────────────────────
    print("\n[3/6] Loading parquet data...")
    df = pd.read_parquet(
        parquet_file_path,
        engine='pyarrow'
    )
    print(f"      Total rows: {len(df):,}")
    print(f"      Columns   : {list(df.columns)}")

    # ── Step 4: Re-assign Group Entity + Vertical ────────────
    print("\n[4/6] Re-assigning Group Entity & Vertical...")
    master_path = os.path.join(
        repo_root, "data",
        "Group_Entity_Vertical_Master.xlsx"
    )
    df, still_nf_df, found_df = assign_group_entity_vertical(
        df, master_path
    )
    has_not_found = not still_nf_df.empty

    # ── Step 5: Save Updated Parquet ─────────────────────────
    print("\n[5/6] Saving updated parquet file...")

    new_parquet_path = os.path.join(
        output_dir,
        f"combined_data_{timestamp_str}.parquet"
    )
    df.to_parquet(
        new_parquet_path,
        index=False,
        engine='pyarrow',
        compression='gzip'
    )
    new_size = (
        os.path.getsize(new_parquet_path)
        / 1024 / 1024
    )
    print(
        f"      💾 Saved: {new_parquet_path} "
        f"({new_size:.1f} MB)"
    )

    # Handle Not Found Summary if still any
    not_found_summary_path = ""
    not_found_json_path    = ""

    if has_not_found:
        print(
            f"\n      ⚠️  {len(still_nf_df):,} rows "
            f"still Not Found after update."
        )
        summary_df = generate_not_found_summary(still_nf_df)

        not_found_summary_path = os.path.join(
            output_dir, "Not_Found_Summary.xlsx"
        )
        summary_df.to_excel(
            not_found_summary_path,
            sheet_name='Not Found Summary',
            index=False,
            engine='openpyxl'
        )

        not_found_json_path = os.path.join(
            output_dir, "not_found_data.json"
        )
        records = []
        for _, row in summary_df.iterrows():
            record = {}
            for col in summary_df.columns:
                val = row[col]
                if hasattr(val, 'item'):
                    val = val.item()
                record[col] = (
                    str(val) if val != '' else ''
                )
            records.append(record)

        with open(not_found_json_path, 'w') as f:
            json.dump(records, f, indent=2)

        print(
            f"      💾 Not Found Summary: "
            f"{not_found_summary_path}"
        )
    else:
        print(
            "\n      ✅ All rows mapped. "
            "No Not Found entities."
        )

    # ── Step 6: Update State File ────────────────────────────
    print("\n[6/6] Updating state file...")

    # For generate_report.py — it needs combined_file_path
    # We pass the parquet as both so it knows what to use
    update_state_file(state_path, {
        "parquet_file_path"      : new_parquet_path,
        "combined_file_path"     : new_parquet_path,
        "has_not_found"          : has_not_found,
        "not_found_summary_path" : not_found_summary_path,
        "not_found_json_path"    : not_found_json_path,
        "start_month"            : start_month,
        "end_month"              : end_month
    })

    print("\n" + "=" * 65)
    print("        ✅ REPROCESS COMBINED COMPLETE")
    print("=" * 65)
    print(f"  Total rows      : {len(df):,}")
    print(f"  Mapped rows     : {len(found_df):,}")
    print(f"  Still Not Found : {len(still_nf_df):,}")
    print(
        f"  Has Not Found   : "
        f"{'YES ⚠️' if has_not_found else 'NO ✅'}"
    )
    print(f"  Parquet file    : {new_parquet_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
