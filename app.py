import streamlit as st
import tempfile
import zipfile
import os
from io import StringIO

import openpyxl
from generate_order_files import process_file

st.set_page_config(
    page_title="Order File Generator",
    page_icon="📋",
    layout="centered",
)

st.title("Order File Generator")
st.caption("Upload a file **or** paste rows from Excel → generate individual order files")

tab_upload, tab_paste = st.tabs(["📁 Upload Excel file", "📋 Paste from Excel"])

# Expected columns (same as sample data)
EXPECTED_HINT = (
    "Ship To | Style# | Order No | PO Num | Sub Bag# | Item Num | Qlty | Size | "
    "Units | Stones# | Carats | Price/Cts | Amt | Due_date | MM Size"
)


def paste_to_xlsx(text: str, dest_path: str) -> int:
    """
    Parse tab- or comma-separated paste from Excel into an .xlsx file.
    Returns number of data rows written (excluding header).
    """
    text = text.strip()
    if not text:
        raise ValueError("Nothing pasted.")

    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 2:
        raise ValueError("Need a header row and at least one data row.")

    # Detect delimiter: tab (Excel default) or comma
    sample = lines[0]
    delim = "\t" if sample.count("\t") >= sample.count(",") else ","

    rows = []
    for ln in lines:
        # Excel sometimes wraps fields in quotes
        parts = [p.strip().strip('"') for p in ln.split(delim)]
        rows.append(parts)

    # Normalize column count
    max_cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < max_cols:
            r.append("")

    wb = openpyxl.Workbook()
    ws = wb.active
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            # Try numeric conversion for data rows
            if r_idx > 1 and val != "":
                try:
                    if "." in val or "e" in val.lower():
                        val = float(val)
                    else:
                        val = int(val)
                except ValueError:
                    pass
            ws.cell(row=r_idx, column=c_idx, value=val)

    wb.save(dest_path)
    return len(rows) - 1  # data rows only


def run_generation(input_path: str):
    """Shared generate + zip download logic."""
    with tempfile.TemporaryDirectory() as tmp:
        out_dir = os.path.join(tmp, "output")
        os.makedirs(out_dir, exist_ok=True)

        logs = []
        def log_cb(msg):
            logs.append(msg)

        created = process_file(input_path, out_dir, log_callback=log_cb)

        st.text_area("Log", value="\n".join(logs), height=220)

        if created:
            zip_path = os.path.join(tmp, "orders.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in created:
                    zf.write(p, arcname=os.path.basename(p))

            with open(zip_path, "rb") as f:
                st.download_button(
                    label=f"⬇ Download {len(created)} order file(s) as ZIP",
                    data=f.read(),
                    file_name="order_files.zip",
                    mime="application/zip",
                )
            st.success(f"Created {len(created)} file(s).")
        else:
            st.error("No files were created. Check the log above.")


# ----- Tab 1: Upload -----
with tab_upload:
    uploaded = st.file_uploader(
        "Daily sample data (.xlsx)",
        type=["xlsx", "xls"],
        help="The daily flat Excel file with Order No, Style#, Stones#, etc.",
        key="uploader",
    )
    if uploaded is not None:
        st.success(f"Loaded: **{uploaded.name}** ({uploaded.size:,} bytes)")
        if st.button("Generate order files", type="primary", key="gen_upload"):
            with st.spinner("Generating order files…"):
                with tempfile.TemporaryDirectory() as tmp:
                    input_path = os.path.join(tmp, uploaded.name)
                    with open(input_path, "wb") as f:
                        f.write(uploaded.getbuffer())
                    run_generation(input_path)
    else:
        st.info("Upload your daily sample Excel file to begin.")


# ----- Tab 2: Paste -----
with tab_paste:
    st.markdown(
        """
        **How to paste**
        1. In Excel, select your data **including the header row**
        2. Copy (`Ctrl+C` / `Cmd+C`)
        3. Click in the box below and paste (`Ctrl+V` / `Cmd+V`)
        """
    )
    st.caption(f"Expected columns (order can vary as long as headers match):\n`{EXPECTED_HINT}`")

    pasted = st.text_area(
        "Paste Excel data here",
        height=260,
        placeholder="Ship To\tStyle#\tOrder No\t...\n100050\t31-D36867SY-E.00\tP2009515\t...",
        key="paste_area",
    )

    if pasted and pasted.strip():
        # Live preview of row count
        n_lines = len([ln for ln in pasted.strip().splitlines() if ln.strip()])
        st.caption(f"Detected **{max(0, n_lines - 1)}** data row(s) (+ 1 header).")

        if st.button("Generate order files", type="primary", key="gen_paste"):
            with st.spinner("Parsing paste and generating…"):
                try:
                    with tempfile.TemporaryDirectory() as tmp:
                        input_path = os.path.join(tmp, "pasted_data.xlsx")
                        n = paste_to_xlsx(pasted, input_path)
                        st.info(f"Parsed **{n}** data row(s) from clipboard.")
                        run_generation(input_path)
                except Exception as e:
                    st.error(f"Could not parse pasted data: {e}")
    else:
        st.info("Copy rows from Excel (with headers) and paste them above.")


st.divider()
st.caption("Powered by openpyxl · Deployed on Streamlit Community Cloud")
