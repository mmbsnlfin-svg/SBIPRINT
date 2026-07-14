# Excel Splitter & PDF Generator

A Streamlit version of the two desktop Python programs. It:

1. Uploads one `.xlsx` or `.xlsm` file.
2. Splits the active worksheet using a selected column.
3. Adds serial numbers and total formulas.
4. Generates one Excel file for every unique split value.
5. Optionally converts all generated Excel files to PDF using LibreOffice.
6. Provides ZIP downloads for Excel files, PDF files, or both.

## GitHub files

Keep these three deployment files in the repository root:

- `app.py`
- `requirements.txt`
- `packages.txt`

`packages.txt` installs LibreOffice on Streamlit Community Cloud. Without LibreOffice, Excel generation works but PDF conversion does not.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

For PDF generation, install LibreOffice on the computer and ensure `libreoffice` or `soffice` is available in the system PATH.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and `packages.txt` to the repository root.
3. In Streamlit Community Cloud, select the repository and choose `app.py` as the main file.
4. Deploy the app.

## Workbook assumptions

- Row 1 and row 2 are headers.
- Data starts from row 3.
- The active worksheet is processed.
- Total formulas are inserted in the selected start-to-end column range.
