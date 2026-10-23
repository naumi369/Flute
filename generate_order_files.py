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

# ---------- Styles ----------
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
pale_green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Total row
light_gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # BOM section
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
        total_col = 10       # J
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
    if n_stones <= 6:
        ws.column_dimensions['I'].width = 3.0
    ws.column_dimensions[get_column_letter(total_col)].width = 16.0

    max_col_used = max(last_col, 10)

    # ========== HEADER (original-style) ==========
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=max_col_used)

    # PICTURE
    ws.merge_cells('B2:C2')
    ws['B2'] = "PICTURE"
    ws['B2'].font = arial_10_bold
    ws['B2'].alignment = center
    ws['B2'].fill = no_fill
    apply_outer_border(ws, 2, 2, 2, 3)

    # STYLE INFORMATION
    ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=max_col_used)
    ws['D2'] = "STYLE INFORMATION"
    ws['D2'].font = arial_9_bold
    ws['D2'].alignment = center
    ws['D2'].fill = no_fill
    apply_outer_border(ws, 2, 4, 2, max_col_used)

    # Empty area under PICTURE
    ws.merge_cells('B3:C10')
    apply_outer_border(ws, 3, 2, 10, 3)

    # Left labels (Min. T.Weight under PO Num)
    labels = [
        (3, "STYLE #", style),
        (4, "Order No", order_no),
        (5, "DESCRIPTION", ""),
        (6, "TOTAL CARAT WT.", 0),
        (7, "SHIP TO", ship_to),
        (8, "PO Num", po_num),
        (9, "Min. T.Weight", round(min_t_weight, 3) if isinstance(min_t_weight, float) else min_t_weight),
    ]
    for row, label, value in labels:
        cell_l = ws.cell(row=row, column=4, value=label)
        cell_l.font = arial_8_bold
        cell_l.fill = yellow_fill
        cell_l.border = thin
        cell_l.alignment = left_align

        cell_v = ws.cell(row=row, column=5, value=value)
        cell_v.font = arial_8_bold
        cell_v.border = thin

    # Right labels
    right_labels = [
        (3, "Item #", ""),
        (4, "Vendor Name", "Frederick Goldman"),
        (5, "Vendor #", ""),
        (6, "Due Date", due_date),
        (7, "Units", units),
        (8, "METAL", ""),
    ]
    for row, label, value in right_labels:
        cell_l = ws.cell(row=row, column=7, value=label)
        cell_l.font = arial_8_bold
        cell_l.fill = yellow_fill
        cell_l.border = thin
        cell_l.alignment = left_align

        cell_v = ws.cell(row=row, column=8, value=value)
        cell_v.font = arial_8_bold
        cell_v.border = thin
        cell_v.alignment = center
        if row == 6 and due_date:
            cell_v.number_format = 'YYYY-MM-DD'

    # Tighter header row heights + consistent alignment
    for r in range(3, 11):
        ws.row_dimensions[r].height = 15
    # Number format for Min. T.Weight (3 decimals)
    ws.cell(row=9, column=5).number_format = '0.000'


    # Blank row 11 after METAL
    # Sequence numbers row 12 (invisible)
    SEQ_ROW = 12
    for i in range(display_slots):
        col = FIRST_STONE_COL + i
        cell = ws.cell(row=SEQ_ROW, column=col, value=i + 1)
        cell.font = white_font
        cell.alignment = center
        cell.border = thin

    # ========== BOM SECTION ==========
    SUBBAG_ROW = 13
    QLTY_ROW = 14
    SIZE_ROW = 15
    STONES_ROW = 16
    CARAT_ROW = 17
    FLUTE_HDR_ROW = 18
    FIRST_FLUTE_ROW = 19

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
    ws.cell(row=SUBBAG_ROW, column=2).fill = yellow_fill

    for col in stone_col_list:
        cell = ws.cell(row=SUBBAG_ROW, column=col, value="")
        cell.fill = yellow_fill
        cell.border = thin
        cell.alignment = center
    for i, s in enumerate(stone_cols):
        col = stone_col(i)
        cell = ws.cell(row=SUBBAG_ROW, column=col, value=s["sub_bag"])
        cell.font = arial_10_bold
        cell.fill = yellow_fill
        cell.alignment = center
        cell.border = thin
    ws.cell(row=SUBBAG_ROW, column=total_col).fill = yellow_fill
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
        cell.fill = yellow_fill

        for col in stone_col_list:
            c = ws.cell(row=r_idx, column=col, value="")
            c.fill = yellow_fill
            c.border = thin
            c.alignment = center

        for i, s in enumerate(stone_cols):
            col = stone_col(i)
            c = ws.cell(row=r_idx, column=col, value=s[key])
            c.font = arial_10_bold
            c.fill = yellow_fill
            c.alignment = center
            c.border = thin

        ws.cell(row=r_idx, column=total_col).fill = yellow_fill
        ws.cell(row=r_idx, column=total_col).border = thin

    # Total Carat Weight header
    cell = ws.cell(row=QLTY_ROW, column=total_col, value="Total Carat Weight")
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = yellow_fill

    # Pre-calculated carat row
    ws.cell(row=CARAT_ROW, column=2, value=1).font = arial_10_bold
    ws.cell(row=CARAT_ROW, column=2).alignment = center
    ws.cell(row=CARAT_ROW, column=2).border = thin

    for col in stone_col_list:
        c = ws.cell(row=CARAT_ROW, column=col, value=0)
        c.font = arial_10_bold
        c.alignment = center
        c.border = thin

    for i, s in enumerate(stone_cols):
        col = stone_col(i)
        val = round(s["carat_per"], 6) if isinstance(s["carat_per"], float) else s["carat_per"]
        cell = ws.cell(row=CARAT_ROW, column=col, value=val)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        if isinstance(s["carat_per"], float):
            cell.number_format = '0.######'

    # Flute # header
    ws.cell(row=FLUTE_HDR_ROW, column=2, value="Flute #").font = arial_10_bold
    ws.cell(row=FLUTE_HDR_ROW, column=2).alignment = center
    ws.cell(row=FLUTE_HDR_ROW, column=2).border = thin
    ws.cell(row=FLUTE_HDR_ROW, column=2).fill = yellow_fill

    for col in stone_col_list:
        c = ws.cell(row=FLUTE_HDR_ROW, column=col)
        c.border = thin
        c.fill = yellow_fill
    ws.cell(row=FLUTE_HDR_ROW, column=total_col).border = thin
    ws.cell(row=FLUTE_HDR_ROW, column=total_col).fill = yellow_fill

    first_stone_letter = get_column_letter(FIRST_STONE_COL)
    last_stone_letter = get_column_letter(last_stone_col)

    # Flute data rows
    for u in range(1, units + 1):
        r = FIRST_FLUTE_ROW + u - 1
        cell = ws.cell(row=r, column=2, value=u)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = yellow_fill

        for col in stone_col_list:
            c = ws.cell(row=r, column=col, value="")
            c.border = thin
            c.alignment = center

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
    within_formula = f'AND({total_letter}{FIRST_FLUTE_ROW}>=$E$9*0.9,{total_letter}{FIRST_FLUTE_ROW}<=$E$9*1.1)'
    outside_formula = f'OR({total_letter}{FIRST_FLUTE_ROW}<$E$9*0.9,{total_letter}{FIRST_FLUTE_ROW}>$E$9*1.1)'
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[within_formula], fill=green_fill_cf))
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[outside_formula], fill=red_fill_cf))


    # Total row
    cell = ws.cell(row=total_row, column=2, value="Total")
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = pale_green_fill

    for col in stone_col_list:
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}{FIRST_FLUTE_ROW}:{col_letter}{last_flute_row})"
        cell = ws.cell(row=total_row, column=col, value=formula)
        cell.font = arial_10_bold
        cell.alignment = center
        cell.border = thin
        cell.fill = pale_green_fill

    total_letter = get_column_letter(total_col)
    formula = f"=SUM({total_letter}{FIRST_FLUTE_ROW}:{total_letter}{last_flute_row})"
    cell = ws.cell(row=total_row, column=total_col, value=formula)
    cell.font = arial_10_bold
    cell.alignment = center
    cell.border = thin
    cell.fill = pale_green_fill

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{order_no}.xlsx")
    wb.save(out_path)
    return out_path, n_stones, units, min_t_weight


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
