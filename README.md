# Streamlit Excel/PDF Generator - Large Data Version

This version is optimized for input files with 20,000+ rows.

## Performance changes
- Reads the source workbook in read-only mode.
- Scans source rows only once.
- Groups rows by GST STATE during the same scan.
- Copies only required values, not source styles.
- Applies standardized output formatting after data creation.
- Shows row-processing progress in Streamlit.
- PDF generation is optional and disabled by default for very large files.

## Deployment
Upload `app.py`, `requirements.txt`, and `packages.txt` to the GitHub repository root.
