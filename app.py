import streamlit as st
import tempfile
import zipfile
import os

from generate_order_files import process_file

st.set_page_config(
    page_title="Order File Generator",
    page_icon="📋",
    layout="centered",
)

st.title("Order File Generator")
st.caption("Upload daily sample Excel → generate individual order files")

uploaded = st.file_uploader(
    "Daily sample data (.xlsx)",
    type=["xlsx", "xls"],
    help="The daily flat Excel file with Order No, Style#, Stones#, etc.",
)

if uploaded is not None:
    st.success(f"Loaded: **{uploaded.name}** ({uploaded.size:,} bytes)")

    if st.button("Generate order files", type="primary"):
        with st.spinner("Generating order files…"):
            with tempfile.TemporaryDirectory() as tmp:
                input_path = os.path.join(tmp, uploaded.name)
                with open(input_path, "wb") as f:
                    f.write(uploaded.getbuffer())

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
else:
    st.info("Upload your daily sample Excel file to begin.")

st.divider()
st.caption("Powered by openpyxl · Deployed on Streamlit Community Cloud")
