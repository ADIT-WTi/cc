# ============================================================
#         MASTER FILE UPDATER - GITHUB ACTIONS VERSION
# ============================================================
# What this script does:
#   1. Reads uploaded Not_Found_Summary.xlsx from data/ folder
#   2. Extracts filled Group Entity + Vertical mappings
#   3. Appends only truly new rows to master file
#   4. Resets state flags for clean report generation
#   5. Deletes Not_Found_Summary.xlsx from data/ folder
# ============================================================

import pandas as pd
import os
import json
from datetime import datetime


# ============================================================
#                    PATH CONFIGURATION
# ============================================================

def get_repo_root():
    """Returns the root directory of the repository."""
    return os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )


def read_state_file():
    """Reads state file and resolves all paths."""
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

    # Resolve all path fields to absolute paths
    path_keys = [
        "output_dir",
        "not_found_summary_path",
        "not_found_json_path",
        "combined_file_path",
        "sample_invoice_path",
        "final_report_path",
        "final_sample_invoice_path"
    ]

    for key in path_keys:
        if key in state and state[key]:
            state[key] = resolve_path(state[key])

    return state, state_path


def update_state_file(state_path, updates):
    """Updates specific keys in the state JSON file."""
    with open(state_path, "r") as f:
        state = json.load(f)

    state.update(updates)

    with open(state_path, "w") as f:
        json.dump(state, f, indent=4)

    print(f"      📝 State file updated.")
    return state

def resolve_path(relative_or_absolute_path):
    """
    Resolve a path that may be relative or absolute
    from a different runner.
    Always returns absolute path on current runner.
    """
    if not relative_or_absolute_path:
        return relative_or_absolute_path

    repo_root = get_repo_root()

    # If already absolute and exists → use it
    if os.path.isabs(relative_or_absolute_path):
        if os.path.exists(relative_or_absolute_path):
            return relative_or_absolute_path

        # Extract relative part from old absolute path
        for marker in ['reports/', 'reports' + os.sep,
                       'data/', 'data' + os.sep]:
            idx = relative_or_absolute_path.find(marker)
            if idx != -1:
                rel_part = relative_or_absolute_path[idx:]
                resolved = os.path.join(repo_root, rel_part)
                return resolved

    # Relative path → join with repo root
    return os.path.join(
        repo_root, relative_or_absolute_path
    )


# ============================================================
#                    MASTER FILE UPDATER
# ============================================================

def update_master_file(master_path, new_rows_df):
    """
    Append only truly new rows to master file.
    New rows are appended at END of existing data.
    Complete duplicate rows are skipped.

    Returns rows_before, rows_after, rows_added.
    """
    master_df   = pd.read_excel(
        master_path,
        sheet_name='Sheet1',
        engine='openpyxl'
    )
    rows_before = len(master_df)

    # Normalize strings
    for col in [
        'Corporate Name', 'Hub',
        'GroupEntity', 'Vertical'
    ]:
        if col in master_df.columns:
            master_df[col] = (
                master_df[col]
                .fillna('').astype(str).str.strip()
            )
        if col in new_rows_df.columns:
            new_rows_df[col] = (
                new_rows_df[col]
                .fillna('').astype(str).str.strip()
            )

    # Find rows NOT already in master
    # (all 4 columns must match to be considered duplicate)
    merged = new_rows_df.merge(
        master_df,
        on=[
            'Corporate Name', 'Hub',
            'GroupEntity', 'Vertical'
        ],
        how='left',
        indicator=True
    )

    truly_new = new_rows_df[
        merged['_merge'] == 'left_only'
    ].copy().reset_index(drop=True)

    print(
        f"      New unique rows to add: "
        f"{len(truly_new):,}"
    )

    if truly_new.empty:
        print(
            "      ℹ️  No new rows to add. "
            "Master unchanged."
        )
        return rows_before, rows_before, 0

    # Append at END
    updated_master = pd.concat(
        [master_df, truly_new],
        ignore_index=True
    )

    rows_after = len(updated_master)
    rows_added = rows_after - rows_before

    updated_master.to_excel(
        master_path,
        sheet_name='Sheet1',
        index=False,
        engine='openpyxl'
    )

    return rows_before, rows_after, rows_added


# ============================================================
#                       MAIN FUNCTION
# ============================================================

def main():
    print("\n" + "=" * 65)
    print("     🔄 MASTER FILE UPDATER (CI/CD)")
    print("=" * 65)

    repo_root = get_repo_root()

    # ── Step 1: Read State File ──────────────────────────────
    print("\n[1/5] Reading state file...")
    state, state_path = read_state_file()

    has_not_found = state.get("has_not_found", False)
    print(f"      Has Not Found : {has_not_found}")

    if not has_not_found:
        print(
            "      ✅ No Not Found entities. "
            "Master update not required."
        )
        return

    # ── Step 2: Load Not Found Summary uploaded by user ──────
    print("\n[2/5] Loading uploaded Not Found Summary...")

    not_found_upload_path = os.path.join(
        repo_root, "data", "Not_Found_Summary.xlsx"
    )

    if not os.path.exists(not_found_upload_path):
        raise FileNotFoundError(
            f"❌ Not Found Summary not found: "
            f"{not_found_upload_path}\n"
            f"   Please upload filled "
            f"Not_Found_Summary.xlsx to data/ folder."
        )

    nf_df = pd.read_excel(
        not_found_upload_path,
        engine='openpyxl'
    )

    print(f"      Total rows in uploaded file : {len(nf_df):,}")
    print(f"      Columns : {list(nf_df.columns)}")

    # Validate required columns
    required_cols = [
        'Corporate', 'Hub', 'Group Entity', 'Vertical'
    ]
    missing = [
        c for c in required_cols
        if c not in nf_df.columns
    ]
    if missing:
        raise ValueError(
            f"❌ Missing columns in uploaded file: {missing}"
        )

    # Filter only rows where BOTH Group Entity
    # AND Vertical are filled by user
    filled_mask = (
        nf_df['Group Entity'].notna() &
        nf_df['Vertical'].notna() &
        (
            nf_df['Group Entity']
            .astype(str).str.strip() != ''
        ) &
        (
            nf_df['Vertical']
            .astype(str).str.strip() != ''
        )
    )

    filled_df   = nf_df[filled_mask].copy()
    unfilled_df = nf_df[~filled_mask].copy()

    print(f"      Filled rows   : {len(filled_df):,}")
    print(f"      Unfilled rows : {len(unfilled_df):,}")

    if filled_df.empty:
        print(
            "      ⚠️  No filled rows found. "
            "Nothing to update."
        )
        return

    # ── Step 3: Update Master File ───────────────────────────
    print("\n[3/5] Updating master file...")

    master_path = os.path.join(
        repo_root, "data",
        "Group_Entity_Vertical_Master.xlsx"
    )

    if not os.path.exists(master_path):
        raise FileNotFoundError(
            f"❌ Master file not found: {master_path}"
        )

    # Build new rows in master file format
    new_rows_df = (
        filled_df[[
            'Corporate', 'Hub',
            'Group Entity', 'Vertical'
        ]]
        .drop_duplicates()
        .rename(columns={
            'Corporate'    : 'Corporate Name',
            'Hub'          : 'Hub',
            'Group Entity' : 'GroupEntity',
            'Vertical'     : 'Vertical'
        })
        .reset_index(drop=True)
    )

    rows_before, rows_after, rows_added = \
        update_master_file(master_path, new_rows_df)

    print(f"\n      Master File Update Summary:")
    print(f"      Rows before  : {rows_before:,}")
    print(f"      Rows after   : {rows_after:,}")
    print(f"      Rows added   : {rows_added:,}")

    # ── Step 4: Reset State for Clean Report Generation ──────
    print("\n[4/5] Resetting state file flags...")

    # Read current state first
    with open(state_path, "r") as f:
        current_state = json.load(f)

    # Update only specific keys
    current_state["has_not_found"]           = False
    current_state["not_found_summary_path"]  = ""
    current_state["not_found_json_path"]     = ""
    current_state["master_updated"]          = True
    current_state["master_update_time"]      = datetime.now().strftime(
        "%d-%b-%Y %H:%M"
    )
    current_state["rows_added_to_master"]    = rows_added

    with open(state_path, "w") as f:
        json.dump(current_state, f, indent=4)

    print("      ✅ State updated:")
    print("         has_not_found  → False")
    print("         master_updated → True")

    # ── Step 5: Delete Not Found Summary from data/ ──────────
    print("\n[5/5] Cleaning up uploaded Not Found Summary...")

    if os.path.exists(not_found_upload_path):
        os.remove(not_found_upload_path)
        print(
            "      ✅ Not_Found_Summary.xlsx deleted "
            "from data/ folder."
        )
    else:
        print(
            "      ℹ️  File already removed."
        )

    print("\n" + "=" * 65)
    print("      ✅ MASTER UPDATE COMPLETE")
    print("=" * 65)
    print(f"  Rows added to master : {rows_added:,}")
    print(f"  Master file          : {master_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
