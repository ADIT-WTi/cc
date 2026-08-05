# ============================================================
#         CRD BUSINESS REPORT GENERATOR - GITHUB ACTIONS VERSION
# ============================================================
# What this script does:
#   1. Reads state file from process_invoice.py
#   2. Loads combined invoice data
#   3. Filters CRD vertical (excludes Not Found rows)
#   4. Generates professional Excel business report
#   5. Saves report to reports/YYYY/MonthName/
#   6. Updates state file with report path
# ============================================================

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
import os
import json
from datetime import datetime

# ============================================================
#                    PATH CONFIGURATION
# ============================================================

def get_repo_root():
    """Returns the root directory of the repository."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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
    """
    Updates specific keys in the state JSON file.
    """
    with open(state_path, "r") as f:
        state = json.load(f)

    state.update(updates)

    with open(state_path, "w") as f:
        json.dump(state, f, indent=4)

    return state


# ============================================================
#                    STYLE HELPERS
# ============================================================

# ── Color Palette ─────────────────────────────────────────
DARK_NAVY    = "0D1B2A"
NAVY_3       = "1B3A5C"
HEADER_BG    = "2E5090"
LIGHT_HEADER = "3A6BC5"
SUB_BOOK_BG  = "4472C4"
SUB_REV_BG   = "2E75B6"
SUMMARY_BG   = "1F4E79"
GT_COL_BG    = "FFF2CC"
ACCENT_GOLD  = "FFD700"
ACCENT_AMBER = "FFC000"
ALT_ROW_1    = "F7FAFF"
ALT_ROW_2    = "E8F0FE"
WHITE        = "FFFFFF"
BLACK        = "1A1A1A"
GREY_DARK    = "404040"
GREY_MED     = "808080"
BORDER_BLUE  = "B4C6E7"
BORDER_DARK  = "4472C4"
DATA_BOOK_BG1 = "F7FAFF"
DATA_BOOK_BG2 = "EBF1FB"
DATA_REV_BG1  = "EFF6FF"
DATA_REV_BG2  = "E0ECFA"
GT_FONT_DARK  = "1B2A4A"
SUBTITLE_BG   = "2E5090"


def make_fill(c):
    return PatternFill(start_color=c, end_color=c, fill_type='solid')


def make_border(c, s='thin'):
    return Border(
        left   = Side(style=s, color=c),
        right  = Side(style=s, color=c),
        top    = Side(style=s, color=c),
        bottom = Side(style=s, color=c)
    )


def make_font(n='Segoe UI', s=10, b=False,
              i=False, c=BLACK, u=None):
    return Font(
        name=n, size=s, bold=b,
        italic=i, color=c, underline=u
    )


def make_align(h='center', v='center', w=False, i=0):
    return Alignment(
        horizontal=h, vertical=v,
        wrap_text=w, indent=i
    )


thin_border   = make_border(BORDER_BLUE, 'thin')
header_border = make_border(BORDER_DARK, 'thin')
medium_border = make_border(DARK_NAVY,   'medium')


def fill_row(ws, r, c, end_col):
    """Fill entire row with a color."""
    for col in range(1, end_col + 1):
        ws.cell(row=r, column=col).fill = make_fill(c)


# ============================================================
#                    PIVOT BUILDER
# ============================================================

def build_pivots(crd_df, unique_months):
    """
    Build booking count and revenue pivot tables.

    Returns:
        booking_pivot : pivot of unique BookingNo counts
        revenue_pivot : pivot of Revenue in Lakhs INR
    """
    booking_pivot = crd_df.pivot_table(
        index   = ['Group Hub', 'Group Entity'],
        columns = 'Month',
        values  = 'BookingNo',
        aggfunc = 'nunique',
        fill_value = 0
    ).reindex(columns=unique_months, fill_value=0)

    revenue_pivot = (
        crd_df.pivot_table(
            index   = ['Group Hub', 'Group Entity'],
            columns = 'Month',
            values  = 'Revenue',
            aggfunc = 'sum',
            fill_value = 0
        ).reindex(columns=unique_months, fill_value=0)
        / 100_000.0
    )

    num_months = len(unique_months)

    # Summary columns
    booking_pivot['_TotalBk']   = booking_pivot[unique_months].sum(axis=1)
    revenue_pivot['_TotalRev']  = revenue_pivot[unique_months].sum(axis=1)
    booking_pivot['_AvgBk']     = (
        booking_pivot['_TotalBk'] / num_months
    ).round(1)
    revenue_pivot['_AvgRev']    = (
        revenue_pivot['_TotalRev'] / num_months
    ).round(2)
    booking_pivot['_AvgTicket'] = np.where(
        booking_pivot['_TotalBk'] > 0,
        (
            revenue_pivot['_TotalRev'] * 100_000 /
            booking_pivot['_TotalBk']
        ).round(0),
        0
    )

    # Sort by Total Revenue descending
    sort_idx      = revenue_pivot['_TotalRev'].sort_values(
        ascending=False
    ).index
    booking_pivot = booking_pivot.reindex(sort_idx)
    revenue_pivot = revenue_pivot.reindex(sort_idx)

    return booking_pivot, revenue_pivot


# ============================================================
#                    HEADER BUILDER
# ============================================================

def build_headers(ws, unique_months, num_months,
                  data_col_start, gt_book_col,
                  gt_rev_col, avg_bk_col,
                  avg_rev_col, avg_tkt_col,
                  total_cols, num_entities,
                  num_hubs, start_month, end_month):
    """
    Build all header rows:
        Row 1-2 : Title
        Row 3   : Subtitle
        Row 4   : Note
        Row 5   : Spacer
        Row 6   : Month group headers
        Row 7   : Column sub-headers
    """

    # ── Title Rows (1-2) ──────────────────────────────────
    ws.merge_cells(
        start_row=1, start_column=1,
        end_row=2,   end_column=total_cols
    )
    tc = ws.cell(
        1, 1,
        "CRD VERTICAL — MONTHLY BOOKING & REVENUE PERFORMANCE"
    )
    tc.font      = make_font(s=22, b=True, c=WHITE)
    tc.fill      = make_fill(DARK_NAVY)
    tc.alignment = make_align()
    fill_row(ws, 1, DARK_NAVY, total_cols)
    fill_row(ws, 2, DARK_NAVY, total_cols)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 26

    # ── Subtitle Row (3) ──────────────────────────────────
    ws.merge_cells(
        start_row=3, start_column=1,
        end_row=3,   end_column=total_cols
    )
    sc_val = (
        f"Reporting Period: {start_month} to {end_month}  |  "
        f"Generated: {pd.Timestamp.now().strftime('%d-%b-%Y %H:%M')}  |  "
        f"Entities: {num_entities:,}  |  "
        f"Hubs: {num_hubs}"
    )
    sc = ws.cell(3, 1, sc_val)
    sc.font      = make_font(s=11, i=True, c=WHITE)
    sc.fill      = make_fill(SUBTITLE_BG)
    sc.alignment = make_align()
    ws.row_dimensions[3].height = 24

    # ── Note Row (4) ──────────────────────────────────────
    ws.merge_cells(
        start_row=4, start_column=1,
        end_row=4,   end_column=total_cols
    )
    nc_val = (
        "⚑  Revenue figures are in LAKHS INR  |  "
        "Use the dropdown on 'Group Hub' to filter — "
        "Grand Total row auto-updates with filter"
    )
    nc = ws.cell(4, 1, nc_val)
    nc.font      = make_font(s=10.5, b=True, i=True, c=ACCENT_GOLD)
    nc.fill      = make_fill(NAVY_3)
    nc.alignment = make_align(w=True)
    ws.row_dimensions[4].height = 28

    # ── Spacer Row (5) ────────────────────────────────────
    fill_row(ws, 5, DARK_NAVY, total_cols)
    ws.row_dimensions[5].height = 3

    # ── Month Group Headers (Row 6) ───────────────────────
    month_grp_row = 6
    ws.row_dimensions[month_grp_row].height = 28

    for m_idx, month in enumerate(unique_months):
        cb = data_col_start + m_idx * 2
        cr = data_col_start + m_idx * 2 + 1

        ws.merge_cells(
            start_row=month_grp_row, start_column=cb,
            end_row=month_grp_row,   end_column=cr
        )
        mc = ws.cell(month_grp_row, cb, month.upper())
        mc.font      = make_font(s=12, b=True, c=WHITE)
        mc.fill      = make_fill(LIGHT_HEADER)
        mc.alignment = make_align()
        mc.border    = header_border

        ws.cell(month_grp_row, cr).fill   = make_fill(LIGHT_HEADER)
        ws.cell(month_grp_row, cr).border = header_border

    # Summary Metrics header
    ws.merge_cells(
        start_row=month_grp_row, start_column=gt_book_col,
        end_row=month_grp_row,   end_column=avg_tkt_col
    )
    sm = ws.cell(month_grp_row, gt_book_col, "SUMMARY METRICS")
    sm.font      = make_font(s=12, b=True, c=ACCENT_AMBER)
    sm.fill      = make_fill(SUMMARY_BG)
    sm.alignment = make_align()
    sm.border    = header_border

    # Fixed column headers for row 6
    for c in range(1, 4):
        ws.cell(month_grp_row, c).fill   = make_fill(HEADER_BG)
        ws.cell(month_grp_row, c).border = header_border

    # ── Column Sub-Headers (Row 7) ────────────────────────
    header_row = 7
    ws.row_dimensions[header_row].height = 36

    fixed_headers = [
        ('S.No',       8),
        ('GROUP HUB',  11),
        ('GROUP ENTITY', 11)
    ]

    for i, (label, size) in enumerate(fixed_headers, 1):
        c = ws.cell(header_row, i, label)
        c.font      = make_font(s=size, b=True, c=WHITE)
        c.fill      = make_fill(HEADER_BG)
        c.alignment = make_align(w=True)
        c.border    = header_border

    # Month sub-headers (Bookings / Revenue per month)
    for m_idx in range(num_months):
        for c_off, label, bg in [
            (0, 'Bookings', SUB_BOOK_BG),
            (1, 'Revenue',  SUB_REV_BG)
        ]:
            c = ws.cell(
                header_row,
                data_col_start + m_idx * 2 + c_off,
                label
            )
            c.font      = make_font(s=9.5, b=True, c=WHITE)
            c.fill      = make_fill(bg)
            c.alignment = make_align()
            c.border    = header_border

    # Summary column headers
    summary_headers = [
        (gt_book_col, 'Total\nBookings',     11, True,  WHITE),
        (gt_rev_col,  'Total\nRevenue',      11, True,  ACCENT_AMBER),
        (avg_bk_col,  'Avg Bookings\n/ Month', 9, True, WHITE),
        (avg_rev_col, 'Avg Revenue\n/ Month',  9, True, WHITE),
        (avg_tkt_col, 'Avg Ticket\nSize (₹)',  9, True, WHITE)
    ]

    for col, label, size, bold, fc in summary_headers:
        c = ws.cell(header_row, col, label)
        c.font      = make_font(s=size, b=bold, c=fc)
        c.fill      = make_fill(SUMMARY_BG)
        c.alignment = make_align(w=True)
        c.border    = header_border

    # Apply borders to all header cells
    for c in range(1, total_cols + 1):
        ws.cell(month_grp_row, c).border = header_border
        ws.cell(header_row,    c).border = header_border


# ============================================================
#                    DATA ROW WRITER
# ============================================================

def write_data_rows(ws, entities_sorted, booking_pivot,
                    revenue_pivot, unique_months,
                    data_col_start, gt_book_col,
                    data_start_row):
    """
    Write all entity data rows to the worksheet.
    Returns the last data row number.
    """
    for sno, (hub, entity) in enumerate(entities_sorted, 1):
        row     = data_start_row + sno - 1
        is_even = (sno % 2 == 0)

        row_bg  = ALT_ROW_2    if is_even else ALT_ROW_1
        bk_bg   = DATA_BOOK_BG2 if is_even else DATA_BOOK_BG1
        rev_bg  = DATA_REV_BG2  if is_even else DATA_REV_BG1

        ws.row_dimensions[row].height = 19

        # S.No
        c = ws.cell(row, 1, sno)
        c.font      = make_font(s=9, c=GREY_MED)
        c.fill      = make_fill(row_bg)
        c.alignment = make_align()
        c.border    = thin_border

        # Group Hub
        c = ws.cell(row, 2, hub)
        c.font      = make_font(s=10, b=True, c=DARK_NAVY)
        c.fill      = make_fill(row_bg)
        c.alignment = make_align(h='left', i=1)
        c.border    = thin_border

        # Group Entity
        c = ws.cell(row, 3, entity)
        c.font      = make_font(s=10, c=BLACK)
        c.fill      = make_fill(row_bg)
        c.alignment = make_align(h='left', i=1)
        c.border    = thin_border

        # Monthly Data
        for m_idx, month in enumerate(unique_months):
            bval = booking_pivot.loc[(hub, entity), month]
            rval = revenue_pivot.loc[(hub, entity), month]

            # Bookings cell
            b = ws.cell(row, data_col_start + m_idx * 2)
            b.value         = int(bval) if bval > 0 else None
            b.number_format = '#,##0'
            b.font          = make_font(
                s=10, b=(bval > 0),
                c=DARK_NAVY if bval > 0 else GREY_MED
            )
            b.fill          = make_fill(bk_bg)
            b.alignment     = make_align()
            b.border        = thin_border

            # Revenue cell
            r = ws.cell(row, data_col_start + m_idx * 2 + 1)
            r.value         = round(float(rval), 2) if rval > 0 else None
            r.number_format = '#,##0.00'
            r.font          = make_font(
                s=10,
                c=BLACK if rval > 0 else GREY_MED
            )
            r.fill          = make_fill(rev_bg)
            r.alignment     = make_align()
            r.border        = thin_border

        # Summary Columns
        summary_vals = [
            (
                booking_pivot.loc[(hub, entity), '_TotalBk'],
                '#,##0',    12, True,  GT_FONT_DARK
            ),
            (
                revenue_pivot.loc[(hub, entity), '_TotalRev'],
                '#,##0.00', 11, True,  GT_FONT_DARK
            ),
            (
                booking_pivot.loc[(hub, entity), '_AvgBk'],
                '#,##0.0',  10, False, GREY_DARK
            ),
            (
                revenue_pivot.loc[(hub, entity), '_AvgRev'],
                '#,##0.00', 10, False, GREY_DARK
            ),
            (
                booking_pivot.loc[(hub, entity), '_AvgTicket'],
                '#,##0',    10, False, GREY_DARK
            )
        ]

        for i, (val, fmt, size, bold, fc) in enumerate(summary_vals):
            c = ws.cell(row, gt_book_col + i)
            c.value         = float(val) if val > 0 else None
            c.number_format = fmt
            c.font          = make_font(s=size, b=bold, c=fc)
            c.fill          = make_fill(GT_COL_BG)
            c.alignment     = make_align()
            c.border        = thin_border

    data_end_row = data_start_row + len(entities_sorted) - 1
    return data_end_row


# ============================================================
#                    GRAND TOTAL ROW
# ============================================================

def write_grand_total_row(ws, gt_row, data_start_row,
                          data_end_row, num_months,
                          data_col_start, gt_book_col,
                          total_cols):
    """
    Write the Grand Total row with SUBTOTAL formulas.
    SUBTOTAL(9,...) automatically respects active filters.
    """
    ws.row_dimensions[gt_row].height = 32

    # Fixed labels
    fixed_labels = [
        ('GRAND',               13, ACCENT_GOLD, 'center'),
        ('TOTAL',               13, ACCENT_GOLD, 'left'),
        ('(Updates with Filter)', 9, GREY_MED,   'center')
    ]
    for i, (val, size, fc, al) in enumerate(fixed_labels, 1):
        c = ws.cell(gt_row, i, val)
        c.font      = make_font(
            s=size, b=True, i=(size == 9), c=fc
        )
        c.fill      = make_fill(DARK_NAVY)
        c.alignment = make_align(
            h=al, i=1 if al == 'left' else 0
        )
        c.border    = medium_border

    # Monthly SUBTOTAL formulas
    for m_idx in range(num_months):
        for c_off, fmt in [(0, '#,##0'), (1, '#,##0.00')]:
            col = data_col_start + m_idx * 2 + c_off
            cL  = get_column_letter(col)
            c   = ws.cell(
                gt_row, col,
                f"=SUBTOTAL(9,{cL}{data_start_row}:{cL}{data_end_row})"
            )
            c.number_format = fmt
            c.font          = make_font(s=11, b=True, c=WHITE)
            c.fill          = make_fill(DARK_NAVY)
            c.alignment     = make_align()
            c.border        = medium_border

    # Summary SUBTOTAL formulas
    summary_fmts = [
        ('#,##0',    13, ACCENT_GOLD),
        ('#,##0.00', 13, ACCENT_GOLD),
        ('#,##0.0',  10, WHITE),
        ('#,##0.00', 10, WHITE),
        ('#,##0',    10, WHITE)
    ]
    for i, (fmt, size, fc) in enumerate(summary_fmts):
        col = gt_book_col + i
        cL  = get_column_letter(col)
        c   = ws.cell(
            gt_row, col,
            f"=SUBTOTAL(9,{cL}{data_start_row}:{cL}{data_end_row})"
        )
        c.number_format = fmt
        c.font          = make_font(s=size, b=True, c=fc)
        c.fill          = make_fill(DARK_NAVY)
        c.alignment     = make_align()
        c.border        = medium_border

    # Fill remaining cells
    for col in range(1, total_cols + 1):
        cell = ws.cell(gt_row, col)
        if not cell.value:
            cell.fill   = make_fill(DARK_NAVY)
            cell.border = medium_border


# ============================================================
#                    COLUMN WIDTHS
# ============================================================

def set_column_widths(ws, num_months, data_col_start,
                      gt_book_col):
    """Set optimal column widths for readability."""
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 40

    for m_idx in range(num_months):
        # Bookings column
        ws.column_dimensions[
            get_column_letter(data_col_start + m_idx * 2)
        ].width = 11
        # Revenue column
        ws.column_dimensions[
            get_column_letter(data_col_start + m_idx * 2 + 1)
        ].width = 12

    # Summary columns
    summary_widths = [13, 13, 13, 13, 14]
    for i, width in enumerate(summary_widths):
        ws.column_dimensions[
            get_column_letter(gt_book_col + i)
        ].width = width

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
#                       MAIN FUNCTION
# ============================================================

def main():
    print("\n" + "=" * 65)
    print("     📊 CRD BUSINESS REPORT GENERATOR (CI/CD)")
    print("=" * 65)

    # ── Step 1: Read State File ──────────────────────────────
    print("\n[1/7] Reading state file...")
    state, state_path = read_state_file()

    combined_file_path = state['combined_file_path']
    output_dir         = state['output_dir']
    start_month        = state['start_month']
    end_month          = state['end_month']
    run_date_str       = state['run_date']

    print(f"      Combined file : {combined_file_path}")
    print(f"      Output dir    : {output_dir}")
    print(f"      Date range    : {start_month} → {end_month}")

    # ── Step 2: Load & Filter CRD Data ──────────────────────
    print("\n[2/7] Loading and filtering CRD data...")

    # Support both parquet and excel
    combined_file_path = state['combined_file_path']
    parquet_file_path  = state.get('parquet_file_path', '')

    # Prefer parquet if available and exists
    if parquet_file_path and \
       os.path.exists(parquet_file_path) and \
       parquet_file_path.endswith('.parquet'):
        print(f"      Reading parquet: {parquet_file_path}")
        df = pd.read_parquet(
            parquet_file_path,
            engine='pyarrow'
        )
    elif combined_file_path and \
         os.path.exists(combined_file_path):
        print(
            f"      Reading excel: {combined_file_path}"
        )
        df = pd.read_excel(
            combined_file_path,
            sheet_name='Sheet1',
            engine='openpyxl'
        )
    else:
        raise FileNotFoundError(
            f"❌ No data file found.\n"
            f"   Parquet: {parquet_file_path}\n"
            f"   Excel  : {combined_file_path}"
        )

    # Filter CRD vertical only — exclude Not Found
    crd_df = df[
        (df['Vertical'] == 'CRD') &
        (df['Group Entity'] != 'Not Found') &
        (df['Vertical']     != 'Not Found')
    ].copy()

    # ── Step 3: Build Sorted Month List ─────────────────────
    print("\n[3/7] Sorting months chronologically...")

    def month_sort_key(m):
        try:
            return pd.Period(m, freq='M')
        except Exception:
            return pd.Period('1900-01', freq='M')

    unique_months = sorted(
        [
            m for m in crd_df['Month'].dropna().unique()
            if m not in ('', 'nan', '(blank)')
        ],
        key=month_sort_key
    )

    num_months = len(unique_months)
    print(f"      Months : {unique_months[0]} → {unique_months[-1]}")
    print(f"      Count  : {num_months}")

    # ── Step 4: Build Pivots ─────────────────────────────────
    print("\n[4/7] Building pivot tables...")

    booking_pivot, revenue_pivot = build_pivots(crd_df, unique_months)
    entities_sorted = list(revenue_pivot.index)
    num_entities    = len(entities_sorted)
    num_hubs        = crd_df['Group Hub'].nunique()

    print(f"      Entities : {num_entities:,}")
    print(f"      Hubs     : {num_hubs}")

    # ── Step 5: Calculate Layout Positions ──────────────────
    print("\n[5/7] Building Excel workbook...")

    FIXED_COLS     = 3
    data_col_start = FIXED_COLS + 1
    gt_book_col    = FIXED_COLS + num_months * 2 + 1
    gt_rev_col     = gt_book_col + 1
    avg_bk_col     = gt_rev_col  + 1
    avg_rev_col    = avg_bk_col  + 1
    avg_tkt_col    = avg_rev_col + 1
    total_cols     = avg_tkt_col
    data_start_row = 8   # Row 1-5 headers, Row 6 month groups, Row 7 sub-headers

    # ── Step 6: Create Workbook ──────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "CRD Report"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation   = 'landscape'
    ws.page_setup.fitToWidth    = 1
    ws.page_setup.fitToHeight   = 0

    # Build headers
    build_headers(
        ws           = ws,
        unique_months = unique_months,
        num_months   = num_months,
        data_col_start = data_col_start,
        gt_book_col  = gt_book_col,
        gt_rev_col   = gt_rev_col,
        avg_bk_col   = avg_bk_col,
        avg_rev_col  = avg_rev_col,
        avg_tkt_col  = avg_tkt_col,
        total_cols   = total_cols,
        num_entities = num_entities,
        num_hubs     = num_hubs,
        start_month  = start_month,
        end_month    = end_month
    )

    # Write data rows
    print(f"      Writing {num_entities:,} entity rows...")
    data_end_row = write_data_rows(
        ws             = ws,
        entities_sorted = entities_sorted,
        booking_pivot  = booking_pivot,
        revenue_pivot  = revenue_pivot,
        unique_months  = unique_months,
        data_col_start = data_col_start,
        gt_book_col    = gt_book_col,
        data_start_row = data_start_row
    )

    # Write Grand Total row
    gt_row = data_end_row + 1
    write_grand_total_row(
        ws             = ws,
        gt_row         = gt_row,
        data_start_row = data_start_row,
        data_end_row   = data_end_row,
        num_months     = num_months,
        data_col_start = data_col_start,
        gt_book_col    = gt_book_col,
        total_cols     = total_cols
    )

    # AutoFilter
    ws.auto_filter.ref = (
        f"A{data_start_row - 1}:"
        f"{get_column_letter(total_cols)}{data_end_row}"
    )

    # Column widths
    set_column_widths(ws, num_months, data_col_start, gt_book_col)

    # Freeze panes (freeze after fixed cols and header rows)
    ws.freeze_panes = ws.cell(
        row    = data_start_row,
        column = data_col_start
    )

    # ── Step 7: Save Report ──────────────────────────────────
    print("\n[6/7] Saving report...")

    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = os.path.join(
        output_dir,
        f"CRD_Business_Report_{timestamp}.xlsx"
    )

    # ── Read Not Found data from JSON ─────────────────────────
    # JSON is more reliable than reading Excel file
    has_not_found       = state.get("has_not_found", False)
    not_found_json_path = state.get("not_found_json_path", "")

    nf_df = pd.DataFrame()

    if has_not_found and not_found_json_path:
        if os.path.exists(not_found_json_path):
            try:
                with open(
                    not_found_json_path, 'r',
                    encoding='utf-8'
                ) as f:
                    nf_records = json.load(f)

                nf_df = pd.DataFrame(nf_records)

                # Ensure correct column order
                expected_cols = [
                    'Month', 'CorporateID', 'Corporate',
                    'Hub', 'Group Hub',
                    'Group Entity', 'Vertical',
                    'Bookings Count', 'Revenue'
                ]
                # Only keep columns that exist
                nf_df = nf_df[
                    [
                        c for c in expected_cols
                        if c in nf_df.columns
                    ]
                ]

                print(
                    f"      ✅ Not Found data loaded "
                    f"from JSON: {len(nf_df)} rows"
                )
                print(
                    f"      Columns: {list(nf_df.columns)}"
                )

                # Preview first 3 rows
                for i, row in nf_df.head(3).iterrows():
                    print(
                        f"      Row {i}: "
                        f"{row.get('Corporate','?')} | "
                        f"{row.get('Hub','?')}"
                    )

            except Exception as e:
                print(
                    f"      ❌ Could not load JSON: {e}"
                )
                nf_df = pd.DataFrame()
        else:
            print(
                f"      ⚠️  JSON not found at: "
                f"{not_found_json_path}"
            )

    # ── Save Sheet 1: CRD Report ──────────────────────────────
    wb.save(report_file)
    print(f"      ✅ Sheet 1 (CRD Report) saved.")

    # ── Add Sheet 2: Not Found Summary ───────────────────────
    if has_not_found and not nf_df.empty:
        print(
            f"      📋 Writing Not Found Summary "
            f"Sheet 2 ({len(nf_df)} rows)..."
        )

        from openpyxl import load_workbook as lw
        from openpyxl.styles import (
            Font       as OFont,
            PatternFill as OFill,
            Alignment  as OAlign,
            Border     as OBorder,
            Side       as OSide
        )
        from openpyxl.utils import (
            get_column_letter as gcl
        )

        wb2   = lw(report_file)
        nf_ws = wb2.create_sheet(title="Not Found Summary")
        nf_ws.sheet_view.showGridLines = False

        # ── Colors (Navy Blue theme) ──────────────────────
        C_TITLE  = "0D1B2A"
        C_HEADER = "2E5090"
        C_WHITE  = "FFFFFF"
        C_ALT1   = "F7FAFF"
        C_ALT2   = "E8F0FE"
        C_GOLD   = "FFF2CC"
        C_GOLDF  = "7F6000"
        C_BORDER = "B4C6E7"
        C_BLACK  = "1A1A1A"

        def xfill(c):
            return OFill(
                start_color=c, end_color=c,
                fill_type='solid'
            )

        def xborder():
            s = OSide(style='thin', color=C_BORDER)
            return OBorder(
                left=s, right=s, top=s, bottom=s
            )

        def xfont(sz=10, bold=False,
                  color=C_BLACK, italic=False):
            return OFont(
                name='Segoe UI', size=sz,
                bold=bold, color=color, italic=italic
            )

        def xalign(h='center', v='center', wrap=False):
            return OAlign(
                horizontal=h, vertical=v,
                wrap_text=wrap
            )

        # ── Row 1: Title ──────────────────────────────────
        nf_ws.merge_cells(
            start_row=1, start_column=1,
            end_row=1, end_column=9
        )
        t = nf_ws.cell(
            1, 1,
            "NOT FOUND ENTITIES — "
            "GROUP ENTITY & VERTICAL MISSING"
        )
        t.font      = xfont(sz=14, bold=True, color=C_WHITE)
        t.fill      = xfill(C_TITLE)
        t.alignment = xalign()
        nf_ws.row_dimensions[1].height = 28

        # ── Row 2: Subtitle ───────────────────────────────
        nf_ws.merge_cells(
            start_row=2, start_column=1,
            end_row=2, end_column=9
        )
        s = nf_ws.cell(
            2, 1,
            f"Total Records: {len(nf_df):,}  |  "
            f"Please fill Group Entity & Vertical "
            f"columns and upload to data/ folder."
        )
        s.font      = xfont(
            sz=10, italic=True, color=C_WHITE
        )
        s.fill      = xfill(C_HEADER)
        s.alignment = xalign()
        nf_ws.row_dimensions[2].height = 20

        # ── Row 3: Headers ────────────────────────────────
        headers = [
            'Month', 'CorporateID', 'Corporate',
            'Hub', 'Group Hub',
            'Group Entity', 'Vertical',
            'Bookings Count', 'Revenue'
        ]
        widths  = [
            12, 13, 40, 22, 15,
            20, 15, 15, 14
        ]

        for ci, (hdr, wd) in enumerate(
            zip(headers, widths), 1
        ):
            c = nf_ws.cell(3, ci, hdr)
            c.font      = xfont(
                sz=11, bold=True, color=C_WHITE
            )
            c.fill      = xfill(C_HEADER)
            c.alignment = xalign(wrap=True)
            c.border    = xborder()
            nf_ws.column_dimensions[gcl(ci)].width = wd

        nf_ws.row_dimensions[3].height = 24

        # ── Data Rows ─────────────────────────────────────
        HIGHLIGHT = [6, 7]  # Group Entity, Vertical

        for ri, (_, row) in enumerate(
            nf_df.iterrows(), 4
        ):
            bg = C_ALT2 if (ri % 2 == 0) else C_ALT1

            for ci, col in enumerate(headers, 1):
                raw = row.get(col, '')
                # Clean value
                if pd.isna(raw) or str(raw) == 'nan':
                    val = ''
                else:
                    val = raw

                # Try numeric conversion for
                # Bookings Count and Revenue
                if col in ['Bookings Count', 'Revenue']:
                    try:
                        val = int(float(str(val)))
                    except Exception:
                        val = 0

                cell = nf_ws.cell(ri, ci, val)

                if ci in HIGHLIGHT:
                    cell.fill = xfill(C_GOLD)
                    cell.font = xfont(
                        sz=10, bold=True, color=C_GOLDF
                    )
                else:
                    cell.fill = xfill(bg)
                    cell.font = xfont(sz=10)

                if ci in [1, 2, 5, 8, 9]:
                    cell.alignment = xalign(h='center')
                elif ci in [3, 4]:
                    cell.alignment = xalign(h='left')
                else:
                    cell.alignment = xalign(h='center')

                cell.border = xborder()

            nf_ws.row_dimensions[ri].height = 18

        # ── Freeze + Filter ───────────────────────────────
        nf_ws.freeze_panes = nf_ws.cell(row=4, column=1)
        nf_ws.auto_filter.ref = (
            f"A3:{gcl(9)}{3 + len(nf_df)}"
        )

        # Rename Sheet 1 to CRD Report
        for sname in wb2.sheetnames:
            if sname != "Not Found Summary":
                wb2[sname].title = "CRD Report"
                break

        wb2.save(report_file)

        print(
            f"      ✅ Final report saved with 2 sheets:"
            f"\n         → Sheet 1 : CRD Report"
            f"\n         → Sheet 2 : Not Found Summary"
            f" ({len(nf_df)} rows)"
        )

    else:
        print(
            "      ✅ Report saved: "
            "CRD Report only (no Not Found data)"
        )

    print(f"\n      📁 Saved at: {report_file}")
    # ── Update State File ────────────────────────────────────
    print("\n[7/7] Updating state file...")
    update_state_file(state_path, {
        "final_report_path"  : report_file,
        "total_entities"     : num_entities,
        "total_bookings"     : int(booking_pivot['_TotalBk'].sum()),
        "total_revenue_lakhs": round(
            float(revenue_pivot['_TotalRev'].sum()), 2
        )
    })


if __name__ == "__main__":
    main()
