# Excel Splitter & PDF Generator

Streamlit application for splitting an Excel workbook by GST State Code and creating print-ready Excel and PDF files.

## Main features

- User enters unwanted source columns such as `A,B,C` on the front page.
- Unwanted columns are removed from every generated file.
- Remaining data font automatically becomes larger when fewer columns remain.
- Row 1 is rebuilt as a large bold title with the GST State Code at the right.
- Rows 1 and 2 repeat on every printed page.
- All remaining columns fit on one Legal landscape page width.
- Additional rows continue automatically to page 2, page 3, and further pages.
- Dark Arial fonts and strong borders improve print visibility.

## GitHub / Streamlit deployment

Upload these files to the repository root:

- `app.py`
- `requirements.txt`
- `packages.txt`

Deploy `app.py` from Streamlit Community Cloud.
