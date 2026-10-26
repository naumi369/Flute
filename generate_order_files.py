#!/usr/bin/env python3
"""
Generate individual Order Excel files from daily sample data.
Supports both command-line and a simple GUI.

Layout (reverted to stable version):
- Header style matching original samples (yellow labels)
- Min. T.Weight under PO Num on the left
- Stone columns continuous from C; Total Carat Weight at J by default,
  or right after last stone when there are more than 6 items
- Blank row after METAL
- Item sequence numbers with invisible (white) font
- PICTURE / STYLE INFORMATION / BOM with plain outer borders
- Gridlines off
- Remembers last input/output folders
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from collections import defaultdict
import os
import sys
import threading
import json
from pathlib import Path

# ---------- Styles (modern multi-color palette) ----------
# Header labels (soft slate blue)
header_label_fill = PatternFill(start_color="D6E3F0", end_color="D6E3F0", fill_type="solid")
# Header values / key inputs (light sky blue)
header_value_fill = PatternFill(start_color="DEEBF7", end_color="DEEBF7", fill_type="solid")
# BOM structure rows: Sub-Bag / Qlty / Size / Stones (soft lavender)
bom_struct_fill = PatternFill(start_color="E2D5F1", end_color="E2D5F1", fill_type="solid")
# Pre-calculated carat reference row (warm amber)
carat_ref_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
# Flute data-entry rows (soft mint – user types here)
flute_input_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
# Total row (stronger green)
total_row_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
# Section title bars
section_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
# Legacy aliases used in a few places
yellow_fill = header_label_fill
pale_green_fill = total_row_fill
light_gray_fill = section_fill
no_fill = PatternFill(fill_type=None)
thin_side = Side(style='thin')
dotted_side = Side(style='dotted')
thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
dotted = Border(left=dotted_side, right=dotted_side, top=dotted_side, bottom=dotted_side)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center')
arial_8_bold = Font(name='Arial', size=8, bold=True)
arial_9_bold = Font(name='Arial', size=9, bold=True)
arial_10_bold = Font(name='Arial', size=10, bold=True)
calibri_11_bold = Font(name='Calibri', size=11, bold=True)
white_font = Font(name='Arial', size=10, bold=True, color="FFFFFF")
dark_font = Font(name='Arial', size=8, bold=True, color="1F4E79")

FIRST_STONE_COL = 3  # Column C



def apply_dotted_outer_border(ws, start_row, start_col, end_row, end_col):
    """Outer dotted border around a range (keeps inner cell borders intact)."""
    no = Side(style=None)
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=dotted_side if c == start_col else (existing.left if existing else no),
                right=dotted_side if c == end_col else (existing.right if existing else no),
                top=dotted_side if r == start_row else (existing.top if existing else no),
                bottom=dotted_side if r == end_row else (existing.bottom if existing else no),
            )

def apply_outer_border(ws, start_row, start_col, end_row, end_col):
    no = Side(style=None)
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=thin_side if c == start_col else no,
                right=thin_side if c == end_col else no,
                top=thin_side if r == start_row else no,
                bottom=thin_side if r == end_row else no,
            )


def excel_serial_to_date(serial):
    if isinstance(serial, datetime):
        return serial
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except Exception:
        return None


def create_order_file(order_no, rows, output_dir="output"):
    if not rows:
        return None

    first = rows[0]
    style = first.get("Style#") or first.get("Style #") or ""
    ship_to = first.get("Ship To")
    units = int(first.get("Units") or 1)
    due_serial = first.get("Due_date") or first.get("Due Date")
    due_date = excel_serial_to_date(due_serial)
    po_num = first.get("PO Num") or ""

    stone_cols = []
    for r in rows:
        stones_total = float(r.get("Stones#") or 0)
        size = float(r.get("Size") or 0)
        stones_per = stones_total / units if units else 0
        if abs(stones_per - round(stones_per)) < 1e-9:
            stones_per = int(round(stones_per))
        carat_per = size * stones_per
        stone_cols.append({
            "sub_bag": r.get("Sub Bag#"),
            "qlty": r.get("Qlty") or "",
            "size": size,
            "stones": stones_per,
            "carat_per": carat_per,
        })

    n_stones = len(stone_cols)
    min_t_weight = sum(s["carat_per"] for s in stone_cols)

    # Total Carat Weight: at J when ≤6 stones; after last stone when more
    if n_stones <= 6:
        last_stone_col = 8   # H
        total_col = 9        # I — immediately after last stone, no blank gap
        display_slots = 6
    else:
        last_stone_col = FIRST_STONE_COL + n_stones - 1
        total_col = last_stone_col + 1
        display_slots = n_stones

    last_col = total_col

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.sheet_view.showGridLines = False

    # Column widths
    ws.column_dimensions['A'].width = 5.5
    ws.column_dimensions['B'].width = 9.0
    for col in range(FIRST_STONE_COL, last_stone_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9.0
    ws.column_dimensions[get_column_letter(total_col)].width = 16.0

    max_col_used = max(last_col, 10)

    # ========== HEADER (exact preferred layout from user sample) ==========
    # Row 1: STYLE INFORMATION (A1:H1 gray) | PICTURE (I1:J1)
    # STYLE INFORMATION bar: A1 through column before Total
    style_end = total_col - 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=style_end)
    ws['A1'] = "STYLE INFORMATION"
    ws['A1'].font = arial_9_bold
    ws['A1'].alignment = center
    for c in range(1, style_end + 1):
        ws.cell(row=1, column=c).fill = section_fill
        ws.cell(row=1, column=c).border = thin

    # PICTURE — single cell in the Total Carat Weight column only
    pic = ws.cell(row=1, column=total_col, value="PICTURE")
    pic.font = arial_10_bold
    pic.alignment = center
    pic.border = thin
    pic.fill = no_fill

    # Picture body area: same single column, rows 2–7
    for r in range(2, 8):
        cell = ws.cell(row=r, column=total_col, value="")
        cell.border = thin
        cell.fill = no_fill
    apply_outer_border(ws, 2, total_col, 7, total_col)

    def lbl(row, col, text_val):
        cell = ws.cell(row=row, column=col, value=text_val)
        cell.font = dark_font
        cell.alignment = left_align
        cell.border = thin
        cell.fill = no_fill
        return cell

    def val_cd(row, value, fmt=None):
        """Value in merged C:D with light blue fill."""
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=3, value=value)
        cell.font = arial_8_bold
        cell.alignment = center
        cell.fill = header_value_fill
        cell.border = thin
        ws.cell(row=row, column=4).fill = header_value_fill
        ws.cell(row=row, column=4).border = thin
        if fmt:
            cell.number_format = fmt
        return cell

    def val_g(row, value, fmt=None):
        """Value in column G with light blue fill."""
        cell = ws.cell(row=row, column=7, value=value)
        cell.font = arial_8_bold
        cell.alignment = center
        cell.fill = header_value_fill
        cell.border = thin
        if fmt:
            cell.number_format = fmt
        return cell

    # Row 2: STYLE # | value | Order No | value
    lbl(2, 2, "STYLE #")
    val_cd(2, style)
    lbl(2, 6, "Order No")
    val_g(2, order_no)

    # Row 3: DESCRIPTION | (empty value area) | SHIP TO | value
    lbl(3, 2, "DESCRIPTION")
    val_cd(3, "")
    lbl(3, 6, "SHIP TO")
    val_g(3, ship_to)

    # Row 4: TOTAL CARAT WT. | 0 | PO Num |
    lbl(4, 2, "TOTAL CARAT WT.")
    val_cd(4, 0)
    lbl(4, 6, "PO Num")
    val_g(4, po_num)

    # Row 5: Due Date | date | Units | units
    lbl(5, 2, "Due Date")
    val_cd(5, due_date, fmt='YYYY-MM-DD' if due_date else None)
    lbl(5, 6, "Units")
    val_g(5, units)

    # Row 6: Min. T.Weight | value | Vendor # |
    lbl(6, 2, "Min. T.Weight")
    val_cd(6, round(min_t_weight, 3) if isinstance(min_t_weight, float) else min_t_weight, fmt='0.000')
    lbl(6, 6, "Vendor #")
    val_g(6, "")

    # Row 7: Vendor Name | Frederick Goldman | METAL |
    lbl(7, 2, "Vendor Name")
    val_cd(7, "Frederick Goldman")
    lbl(7, 6, "METAL")
    val_g(7, "")

    # Row heights
    for r in range(1, 8):
        ws.row_dimensions[r].height = 16

    # Column widths matching preferred form
    ws.column_dimensions['A'].width = 5.5
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 9
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 9
    ws.column_dimensions['I'].width = 8
    ws.column_dimensions['J'].width = 16

    # Row 8 blank
    # Row 9: sequence numbers (invisible) in stone columns C–H (or more)
    SEQ_ROW = 9
    for i in range(display_slots):
        col = FIRST_STONE_COL + i
        cell = ws.cell(row=SEQ_ROW, column=col, value=i + 1)
        cell.font = white_font
        cell.alignment = center
        cell.border = thin

    # Ensure stone / total column widths
    for col in range(FIRST_STONE_COL, last_stone_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9.0
    ws.column_dimensions[get_column_letter(total_col)].width = 16.0


    # ========== BOM SECTION ==========
    SUBBAG_ROW = 10
    QLTY_ROW = 11
    SIZE_ROW = 12
    STONES_ROW = 13
    CARAT_ROW = 14
    FLUTE_HDR_ROW = 15
    FIRST_FLUTE_ROW = 16

    def stone_col(i):
        return FIRST_STONE_COL + i

    if n_stones <= 6:
        stone_col_list = list(range(FIRST_STONE_COL, 9))  # C–H
    else:
        stone_col_list = list(range(FIRST_STONE_COL, FIRST_STONE_COL + n_stones))

    # Sub-Bag
    ws.cell(row=SUBBAG_ROW, column=2, value="Sub-Bag").font = arial_10_bold
    ws.cell(row=SUBBAG_ROW, column=2).alignment = center
    ws.cell(row=SUBBAG_ROW, column=2).border = thin
    ws.cell(row=SUBBAG_ROW, column=2).fill = bom_struct_fill

    for col in stone_col_list:
        cell = ws.cell(row=SUBBAG_ROW, column=col, value="")
        cell.fill = bom_struct_fill
        cell.border = thin
        cell.alignment = center
    for i, s in enumerate(stone_cols):
        col = stone_col(i)
        cell = ws.cell(row=SUBBAG_ROW, column=col, value=s["sub_bag"])
        cell.font = arial_10_bold
        cell.fill = bom_struct_fill
        cell.alignment = center
        cell.border = thin
    ws.cell(row=SUBBAG_ROW, column=total_col).fill = bom_struct_fill
    ws.cell(row=SUBBAG_ROW, column=total_col).border = thin

    # BOM label
    ws.merge_cells(start_row=QLTY_ROW, start_column=1, end_row=STONES_ROW, end_column=1)
    ws.cell(row=QLTY_ROW, column=1, value="BOM")
    ws.cell(row=QLTY_ROW, column=1).font = calibri_11_bold
    ws.cell(row=QLTY_ROW, column=1).alignment = center
    ws.cell(row=QLTY_ROW, column=1).fill = no_fill
    apply_outer_border(ws, QLTY_ROW, 1, STONES_ROW, 1)

    # Qlty / Size / Stones
    for r_idx, key, label in [
        (QLTY_ROW, "qlty", "Qlty"),
        (SIZE_ROW, "size", "Size"),
        (STONES_ROW, "stones", "Stones"),
    ]:
        cell = ws.cell(row=r_idx, column=2, value=label)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = bom_struct_fill

        for col in stone_col_list:
            c = ws.cell(row=r_idx, column=col, value="")
            c.fill = bom_struct_fill
            c.border = thin
            c.alignment = center

        for i, s in enumerate(stone_cols):
            col = stone_col(i)
            c = ws.cell(row=r_idx, column=col, value=s[key])
            c.font = arial_10_bold
            c.fill = bom_struct_fill
            c.alignment = center
            c.border = thin

        ws.cell(row=r_idx, column=total_col).fill = bom_struct_fill
        ws.cell(row=r_idx, column=total_col).border = thin

    # Total Carat Weight header
    cell = ws.cell(row=QLTY_ROW, column=total_col, value="Total Carat Weight")
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = bom_struct_fill

    # Pre-calculated carat row
    ws.cell(row=CARAT_ROW, column=2, value=1).font = arial_10_bold
    ws.cell(row=CARAT_ROW, column=2).alignment = center
    ws.cell(row=CARAT_ROW, column=2).border = thin
    ws.cell(row=CARAT_ROW, column=2).fill = carat_ref_fill

    for col in stone_col_list:
        c = ws.cell(row=CARAT_ROW, column=col, value=0)
        c.font = arial_10_bold
        c.alignment = center
        c.border = thin
        c.fill = carat_ref_fill

    for i, s in enumerate(stone_cols):
        col = stone_col(i)
        val = round(s["carat_per"], 6) if isinstance(s["carat_per"], float) else s["carat_per"]
        cell = ws.cell(row=CARAT_ROW, column=col, value=val)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = carat_ref_fill
        if isinstance(s["carat_per"], float):
            cell.number_format = '0.######'

    # Flute # header
    ws.cell(row=FLUTE_HDR_ROW, column=2, value="Flute #").font = arial_10_bold
    ws.cell(row=FLUTE_HDR_ROW, column=2).alignment = center
    ws.cell(row=FLUTE_HDR_ROW, column=2).border = thin
    ws.cell(row=FLUTE_HDR_ROW, column=2).fill = flute_input_fill

    for col in stone_col_list:
        c = ws.cell(row=FLUTE_HDR_ROW, column=col)
        c.border = thin
        c.fill = flute_input_fill
    ws.cell(row=FLUTE_HDR_ROW, column=total_col).border = thin
    ws.cell(row=FLUTE_HDR_ROW, column=total_col).fill = flute_input_fill

    first_stone_letter = get_column_letter(FIRST_STONE_COL)
    last_stone_letter = get_column_letter(last_stone_col)

    # Flute data rows
    for u in range(1, units + 1):
        r = FIRST_FLUTE_ROW + u - 1
        cell = ws.cell(row=r, column=2, value=u)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = flute_input_fill

        for col in stone_col_list:
            c = ws.cell(row=r, column=col, value="")
            c.border = thin
            c.alignment = center
            c.fill = flute_input_fill

        formula = f"=SUM({first_stone_letter}{r}:{last_stone_letter}{r})"
        cell = ws.cell(row=r, column=total_col, value=formula)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin

    last_flute_row = FIRST_FLUTE_ROW + units - 1
    total_row = last_flute_row + 1

    # Light gray background + dotted outer border for BOM entry block
    bom_start_row = SUBBAG_ROW
    bom_end_row = last_flute_row
    for r in range(bom_start_row, bom_end_row + 1):
        for c in stone_col_list:
            cell = ws.cell(row=r, column=c)
            # Keep yellow on header structure rows; light gray on empty flute input cells only
            if r >= FIRST_FLUTE_ROW and (cell.value is None or cell.value == ""):
                if cell.fill is None or cell.fill.fgColor is None or cell.fill.patternType is None:
                    cell.fill = light_gray_fill
        # Also light gray on empty total-col flute cells is not needed (has formula)
    # Dotted outer border around the whole BOM data block
    apply_dotted_outer_border(ws, bom_start_row, 2, last_flute_row, total_col)

    # Conditional formatting: Flute Total Carat Weight within 10% of Min. T.Weight (E9)
    # Green if within ±10%, red if outside
    total_letter = get_column_letter(total_col)
    cf_range = f"{total_letter}{FIRST_FLUTE_ROW}:{total_letter}{last_flute_row}"
    # Within 10%: >= Min*0.9 AND <= Min*1.1
    green_fill_cf = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill_cf = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    # Formula relative to top-left of range
    within_formula = f'AND({total_letter}{FIRST_FLUTE_ROW}>=$C$6*0.9,{total_letter}{FIRST_FLUTE_ROW}<=$C$6*1.1)'
    outside_formula = f'OR({total_letter}{FIRST_FLUTE_ROW}<$C$6*0.9,{total_letter}{FIRST_FLUTE_ROW}>$C$6*1.1)'
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[within_formula], fill=green_fill_cf))
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[outside_formula], fill=red_fill_cf))


    # Total row
    cell = ws.cell(row=total_row, column=2, value="Total")
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = total_row_fill

    for col in stone_col_list:
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}{FIRST_FLUTE_ROW}:{col_letter}{last_flute_row})"
        cell = ws.cell(row=total_row, column=col, value=formula)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = total_row_fill

    total_letter = get_column_letter(total_col)
    formula = f"=SUM({total_letter}{FIRST_FLUTE_ROW}:{total_letter}{last_flute_row})"
    cell = ws.cell(row=total_row, column=total_col, value=formula)
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = total_row_fill

    # ----- Print-ready setup -----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # as many pages tall as needed
    ws.page_setup.paperSize = 9  # A4
    ws.print_area = f"A1:{get_column_letter(total_col)}{total_row}"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ----- Print setup: A4, landscape, narrow margins, fit to width -----
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # as many pages tall as needed
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1
    # Print area: used block
    last_print_row = total_row
    last_print_col = get_column_letter(total_col)
    ws.print_area = f"A1:{last_print_col}{last_print_row}"
    ws.print_title_rows = "1:9"  # repeat header rows on each printed page

    # ----- Print setup: A4, portrait, narrow margins -----
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # as many pages tall as needed
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    # Narrow margins (inches)
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    # Print area = used range
    last_row = total_row
    last_print_col = get_column_letter(total_col)
    ws.print_area = f"A1:{last_print_col}{last_row}"
    ws.print_title_rows = "1:9"  # repeat header rows on each printed page

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{order_no}.xlsx")
    wb.save(out_path)
    return out_path, n_stones, units, min_t_weight



def merge_order_files(file_paths, output_path):
    """
    Merge multiple order Excel files onto one sheet, stacked top-to-bottom
    with 2 blank rows between each file.

    Print: A4, portrait, narrow margins, fit to width.
    No forced page breaks between files — short forms share a page when they fit.
    """
    from copy import copy
    from openpyxl.utils import get_column_letter

    if not file_paths:
        raise ValueError("No files to merge.")

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Merged Orders"
    out_ws.sheet_view.showGridLines = False

    current_row = 1
    max_col_used = 1
    col_widths = {}

    for idx, fpath in enumerate(file_paths):
        src_wb = openpyxl.load_workbook(fpath)
        src_ws = src_wb.active

        # Determine used range
        min_r, min_c = 1, 1
        max_r = src_ws.max_row or 1
        max_c = src_ws.max_column or 1
        # Tighten max_r/max_c to last non-empty
        for r in range(max_r, 0, -1):
            if any(src_ws.cell(r, c).value not in (None, "") for c in range(1, max_c + 1)):
                max_r = r
                break
        for c in range(max_c, 0, -1):
            if any(src_ws.cell(r, c).value not in (None, "") for r in range(1, max_r + 1)):
                max_c = c
                break

        max_col_used = max(max_col_used, max_c)

        # Copy column widths from first file (and expand if later files are wider)
        for c in range(1, max_c + 1):
            letter = get_column_letter(c)
            w = src_ws.column_dimensions[letter].width
            if w:
                col_widths[letter] = max(col_widths.get(letter, 0), w)

        # Copy merged cells (offset by current_row - 1)
        row_offset = current_row - 1
        merges_to_add = []
        for merged in src_ws.merged_cells.ranges:
            merges_to_add.append((
                merged.min_row + row_offset,
                merged.min_col,
                merged.max_row + row_offset,
                merged.max_col,
            ))

        # Copy cells (value + style)
        for r in range(1, max_r + 1):
            src_height = src_ws.row_dimensions[r].height
            if src_height:
                out_ws.row_dimensions[current_row].height = src_height

            for c in range(1, max_c + 1):
                src_cell = src_ws.cell(r, c)
                dst_cell = out_ws.cell(current_row, c)

                dst_cell.value = src_cell.value

                if src_cell.has_style:
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format

            current_row += 1

        for mr in merges_to_add:
            try:
                out_ws.merge_cells(
                    start_row=mr[0], start_column=mr[1],
                    end_row=mr[2], end_column=mr[3],
                )
            except Exception:
                pass

        src_wb.close()

        # 2 blank rows between files (not after the last)
        if idx < len(file_paths) - 1:
            current_row += 2

    # Apply column widths
    for letter, w in col_widths.items():
        out_ws.column_dimensions[letter].width = w

    # Print setup: A4, portrait, narrow margins, fit width (auto-fit)
    out_ws.page_setup.paperSize = out_ws.PAPERSIZE_A4
    out_ws.page_setup.orientation = "portrait"
    out_ws.page_setup.fitToPage = True
    out_ws.page_setup.fitToWidth = 1
    out_ws.page_setup.fitToHeight = 0
    out_ws.sheet_properties.pageSetUpPr.fitToPage = True
    out_ws.page_margins.left = 0.25
    out_ws.page_margins.right = 0.25
    out_ws.page_margins.top = 0.5
    out_ws.page_margins.bottom = 0.5
    out_ws.page_margins.header = 0.2
    out_ws.page_margins.footer = 0.2

    last_col_letter = get_column_letter(max_col_used)
    out_ws.print_area = f"A1:{last_col_letter}{current_row - 1}"
    # No forced page breaks between orders — short forms share a page when they fit

    out_wb.save(output_path)
    return output_path, len(file_paths), current_row - 1

def process_file(input_file, output_dir, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    if not os.path.isfile(input_file):
        log(f"ERROR: Input file not found: {input_file}")
        return []

    try:
        wb = openpyxl.load_workbook(input_file, data_only=True)
        ws = wb.active
    except Exception as e:
        log(f"ERROR: Could not open Excel file: {e}")
        return []

    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    header_map = {}
    for i, h in enumerate(headers):
        if h is None:
            continue
        header_map[i] = str(h).strip()

    orders = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        rec = {}
        for i, val in enumerate(row):
            if i in header_map:
                rec[header_map[i]] = val
        order_no = rec.get("Order No") or rec.get("OrderNo")
        if order_no:
            orders[str(order_no).strip()].append(rec)

    log(f"Found {len(orders)} unique order(s) in the file.")
    created = []

    for order_no, rows in sorted(orders.items()):
        try:
            result = create_order_file(order_no, rows, output_dir=output_dir)
            if result:
                path, n_stones, units, min_tw = result
                log(f"  ✓ {order_no}.xlsx  (stones={n_stones}, units={units}, Min.T.Weight={min_tw:.4f})")
                created.append(path)
        except Exception as e:
            log(f"  ✗ Failed {order_no}: {e}")

    log(f"\nDone. {len(created)} file(s) written to:\n{output_dir}")
    return created


# ---------- Remember last folders ----------
def _config_path():
    try:
        base = Path(__file__).resolve().parent
    except NameError:
        base = Path.home()
    return base / ".order_generator_config.json"

def load_config():
    p = _config_path()
    if p.is_file():
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(data):
    p = _config_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ==================== GUI ====================

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    root = tk.Tk()
    root.title("Order File Generator")
    root.geometry("720x520")
    root.minsize(600, 420)

    cfg = load_config()
    input_var = tk.StringVar(value=cfg.get("last_input", ""))
    output_var = tk.StringVar(value=cfg.get("last_output", os.path.join(os.getcwd(), "output")))

    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    title = ttk.Label(main, text="Order File Generator", font=("Segoe UI", 16, "bold"))
    title.pack(pady=(0, 8))

    subtitle = ttk.Label(main, text="Select the daily sample Excel → choose output folder → Generate",
                         font=("Segoe UI", 9))
    subtitle.pack(pady=(0, 12))

    in_frame = ttk.LabelFrame(main, text="Input Excel file (daily sample data)", padding=8)
    in_frame.pack(fill=tk.X, pady=4)
    ttk.Entry(in_frame, textvariable=input_var, width=70).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    def browse_input():
        path = filedialog.askopenfilename(title="Select sample data Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            input_var.set(path)
    ttk.Button(in_frame, text="Browse…", command=browse_input).pack(side=tk.RIGHT)

    out_frame = ttk.LabelFrame(main, text="Output folder", padding=8)
    out_frame.pack(fill=tk.X, pady=4)
    ttk.Entry(out_frame, textvariable=output_var, width=70).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    def browse_output():
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            output_var.set(path)
    ttk.Button(out_frame, text="Browse…", command=browse_output).pack(side=tk.RIGHT)

    log_frame = ttk.LabelFrame(main, text="Log", padding=6)
    log_frame.pack(fill=tk.BOTH, expand=True, pady=8)
    log_text = scrolledtext.ScrolledText(log_frame, height=14, state="disabled",
                                         font=("Consolas", 9), wrap=tk.WORD)
    log_text.pack(fill=tk.BOTH, expand=True)

    def append_log(msg):
        log_text.configure(state="normal")
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        log_text.configure(state="disabled")
        root.update_idletasks()

    btn_frame = ttk.Frame(main)
    btn_frame.pack(fill=tk.X, pady=4)
    generate_btn = ttk.Button(btn_frame, text="Generate Files")
    generate_btn.pack(side=tk.LEFT, padx=(0, 8))
    open_btn = ttk.Button(btn_frame, text="Open Output Folder")
    open_btn.pack(side=tk.LEFT, padx=(0, 8))
    clear_btn = ttk.Button(btn_frame, text="Clear Log")
    clear_btn.pack(side=tk.LEFT)
    status_var = tk.StringVar(value="Ready")
    ttk.Label(btn_frame, textvariable=status_var, foreground="gray").pack(side=tk.RIGHT)

    def clear_log():
        log_text.configure(state="normal")
        log_text.delete(1.0, tk.END)
        log_text.configure(state="disabled")
    clear_btn.configure(command=clear_log)

    def open_output():
        folder = output_var.get().strip()
        if folder and os.path.isdir(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        else:
            messagebox.showinfo("Info", "Output folder does not exist yet.")
    open_btn.configure(command=open_output)

    def do_generate():
        input_path = input_var.get().strip()
        output_path = output_var.get().strip()
        if not input_path:
            messagebox.showwarning("Missing input", "Please select the daily sample Excel file.")
            return
        if not output_path:
            messagebox.showwarning("Missing output", "Please choose an output folder.")
            return
        generate_btn.configure(state="disabled")
        status_var.set("Working…")
        append_log("=" * 50)
        append_log(f"Input : {input_path}")
        append_log(f"Output: {output_path}")
        append_log("")

        def worker():
            try:
                created = process_file(input_path, output_path, log_callback=append_log)
                root.after(0, lambda: on_finished(True, len(created)))
            except Exception as e:
                append_log(f"\nERROR: {e}")
                root.after(0, lambda: on_finished(False, 0))

        def on_finished(success, count):
            generate_btn.configure(state="normal")
            if success:
                status_var.set(f"Finished – {count} file(s) created")
                save_config({"last_input": input_path, "last_output": output_path})
                if count:
                    messagebox.showinfo("Success", f"Created {count} order file(s).\n\nSaved in:\n{output_path}")
            else:
                status_var.set("Finished with errors")
                messagebox.showerror("Error", "Something went wrong. Check the log.")

        threading.Thread(target=worker, daemon=True).start()

    generate_btn.configure(command=do_generate)
    root.bind("<Control-g>", lambda e: do_generate())
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f"+{x}+{y}")
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] not in ("--gui", "-g"):
        process_file(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "output")
    else:
        run_gui()
