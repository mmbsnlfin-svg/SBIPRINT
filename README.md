# Excel Splitter, State-Parent Summary & PDF Generator

## Main changes

- Creates state-wise Excel files and state-wise PDFs.
- Creates `Summary.xlsx` only. No Summary PDF is generated.
- Summary rows are grouped by **GST State Code + Parent BA**.
- Circuit count is the count of nonblank **LC ID** values for each State/Parent group.
- BSNL GST No. and SBI GST No. are read from the uploaded prior summary/master workbook.
- GST mapping first uses State Code + Parent; if no exact match is found, it uses the State Code GST entry.
- Invoice No. and Invoice Date are intentionally blank for manual entry after generation.
- Summary cells use wrap text and print-friendly widths/heights.

## GitHub / Streamlit files

Place these files in the repository root:

- `app.py`
- `requirements.txt`
- `packages.txt`

Run locally:

```bash
pip install -r requirements.txt
streamlit run app.py
```
