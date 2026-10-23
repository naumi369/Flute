# Order File Generator

Upload a daily sample Excel file and generate individual order BOM Excel files.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. New app → select this repo → Main file: `app.py`
4. Deploy

## Files

- `app.py` – Streamlit web UI
- `generate_order_files.py` – Excel generation logic
- `requirements.txt` – Python dependencies
