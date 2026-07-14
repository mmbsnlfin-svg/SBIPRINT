import copy
import io
import os
import re
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# -----------------------------------------------------------------------------
# Fixed output structure selected by column HEADER NAME, not by column letter.
# This avoids formula breakage when unwanted source columns are removed.
# -----------------------------------------------------------------------------
REQUIRED_OUTPUT_HEADERS = [
    "Sr. No.",
    "ACCOUNT _NUM",
    "LC ID",
    "Bill From",
    "Bill To",
    "Days",
    "Branch Code",
    "Branch Name",
    "C Type",
    "Port BW",
    "Annual Recurring Charges",
    "Quarterly Charges",
    "NTU Chg /Modem Chg",
    "NOFN charges",
    "IDR / Submarine Charges",
    "Total Quarterly charges Gross",
    "GST 18%",
    "Net Payable After Tax",
    "GST STATE",
    "Parent BA",
    "PO NO",
    "PO Date",
]

# Display header -> accepted normalized source-header aliases.
HEADER_ALIASES = {
    "ACCOUNT _NUM": {"ACCOUNTNUM", "ACCOUNTNUMBER", "ACCOUNTNO"},
    "LC ID": {"LCID"},
    "Bill From": {"BILLFROM"},
    "Bill To": {"BILLTO"},
    "Days": {"DAYS"},
    "Branch Code": {"BRANCHCODE"},
    "Branch Name": {"BRANCHNAME"},
    "C Type": {"CTYPE"},
    "Port BW": {"PORTBW", "PORTBANDWIDTH"},
    "Annual Recurring Charges": {
        "ANNUALRECURRINGCHARGES",
        "ANNUALRECURRINGCHARGESAFTERDISCOUNT",
        "ARC",
    },
    "Quarterly Charges": {"QUARTERLYCHARGES"},
    "NTU Chg /Modem Chg": {
        "NTUCHGMODEMCHG",
        "NTUCHARGEMODEMCHARGE",
        "NTUCHARGESMODEMCHARGES",
    },
    "NOFN charges": {"NOFNCHARGES"},
    "IDR / Submarine Charges": {
        "IDRSUBMARINECHARGES",
        "IDRCHARGES",
        "SUBMARINECHARGES",
    },
    "Total Quarterly charges Gross": {
        "TOTALQUARTERLYCHARGESGROSS",
        "TOTALQUARTERLYCHARGES",
    },
    "GST 18%": {"GST18", "GST18PERCENT"},
    "Net Payable After Tax": {"NETPAYABLEAFTERTAX"},
    "GST STATE": {"GSTSTATE", "GSTSTATECODE", "STATE"},
    "Parent BA": {"PARENTBA"},
    "PO NO": {"PONO", "PONUMBER"},
    "PO Date": {"PODATE"},
}

# Columns intentionally excluded from output:
# PO, Type, CGST 9%, SGST 9%, Remarks, Loopback, WAN IP.

HEADER_FILL = PatternFill("solid", fgColor="92D050")
BLACK = "000000"
THIN_SIDE = Side(style="thin", color=BLACK)
MEDIUM_SIDE = Side(style="medium", color=BLACK)
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MEDIUM_BORDER = Border(left=MEDIUM_SIDE, right=MEDIUM_SIDE, top=MEDIUM_SIDE, bottom=MEDIUM_SIDE)

# Widths are keyed by final output header. Title row is never considered.
COLUMN_WIDTHS = {
    "Sr. No.": 6,
    "ACCOUNT _NUM": 14,
    "LC ID": 13,
    "Bill From": 12,
    "Bill To": 12,
    "Days": 6,
    "Branch Code": 10,
    "Branch Name": 25,
    "C Type": 8,
    "Port BW": 10,
    "Annual Recurring Charges": 15,
    "Quarterly Charges": 15,
    "NTU Chg /Modem Chg": 14,
    "NOFN charges": 12,
    "IDR / Submarine Charges": 15,
    "Total Quarterly charges Gross": 17,
    "GST 18%": 13,
    "Net Payable After Tax": 16,
    "GST STATE": 9,
    "Parent BA": 12,
    "PO NO": 13,
    "PO Date": 12,
}

FORMULA_HEADERS = {
    "Days",
    "Quarterly Charges",
    "Total Quarterly charges Gross",
    "GST 18%",
    "Net Payable After Tax",
}

TOTAL_HEADERS = {
    "Annual Recurring Charges",
    "Quarterly Charges",
    "NTU Chg /Modem Chg",
    "NOFN charges",
    "IDR / Submarine Charges",
    "Total Quarterly charges Gross",
    "GST 18%",
    "Net Payable After Tax",
}


def normalize_header(value: Any) -> str:
    """Normalize spaces, NBSP and punctuation for reliable header matching."""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u00a0", " ").replace("\u2007", " ").replace("\u202f", " ")
    return re.sub(r"[^A-Z0-9]", "", text.upper())


def clean_filename(value: Any) -> str:
    text = re.sub(r'[\\/*?:"<>|]', "", str(value).strip())
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:100] or "Blank_Value"


def unique_filename(base_name: str, used: set[str], extension: str) -> str:
    candidate = f"{base_name}{extension}"
    count = 2
    while candidate.lower() in used:
        candidate = f"{base_name}_{count}{extension}"
        count += 1
    used.add(candidate.lower())
    return candidate


def copy_cell_style(src, dst) -> None:
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)
    if src.hyperlink:
        dst._hyperlink = copy.copy(src.hyperlink)
    if src.comment:
        dst.comment = copy.copy(src.comment)


def find_header_row(ws, max_scan_rows: int = 10) -> int:
    """Find the row containing the largest number of required headers."""
    required_aliases = set().union(*HEADER_ALIASES.values())
    best_row = 0
    best_count = 0
    for row in range(1, min(ws.max_row, max_scan_rows) + 1):
        found = sum(
            1
            for cell in ws[row]
            if normalize_header(cell.value) in required_aliases
        )
        if found > best_count:
            best_count = found
            best_row = row
    if best_count < 5:
        raise ValueError("Could not identify the Excel header row. At least five required column names must be present.")
    return best_row


def build_source_header_map(ws, header_row: int) -> dict[str, int]:
    """Return display output header -> source column index."""
    normalized_to_col: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        normalized = normalize_header(ws.cell(header_row, col).value)
        if normalized and normalized not in normalized_to_col:
            normalized_to_col[normalized] = col

    mapping: dict[str, int] = {}
    missing: list[str] = []
    for display_header in REQUIRED_OUTPUT_HEADERS[1:]:  # Sr. No. is generated by code.
        aliases = HEADER_ALIASES[display_header]
        matched_col = next((normalized_to_col[a] for a in aliases if a in normalized_to_col), None)
        if matched_col is None:
            missing.append(display_header)
        else:
            mapping[display_header] = matched_col

    if missing:
        raise ValueError(
            "The following required columns were not found by header name: "
            + ", ".join(missing)
            + ". Please use the original Excel file containing all required columns."
        )
    return mapping


def output_column_map() -> dict[str, int]:
    return {header: index for index, header in enumerate(REQUIRED_OUTPUT_HEADERS, start=1)}


def excel_ref(header: str, row: int, out_map: dict[str, int]) -> str:
    return f"{get_column_letter(out_map[header])}{row}"


def data_formula(header: str, row: int, out_map: dict[str, int]) -> str:
    """Rebuild formulas from stable output headers, so deleted source columns cannot break them."""
    if header == "Days":
        return f"={excel_ref('Bill To', row, out_map)}-{excel_ref('Bill From', row, out_map)}+1"
    if header == "Quarterly Charges":
        return f"={excel_ref('Annual Recurring Charges', row, out_map)}/4"
    if header == "Total Quarterly charges Gross":
        refs = [
            excel_ref("Quarterly Charges", row, out_map),
            excel_ref("NTU Chg /Modem Chg", row, out_map),
            excel_ref("NOFN charges", row, out_map),
            excel_ref("IDR / Submarine Charges", row, out_map),
        ]
        return f"=SUM({refs[0]}:{refs[-1]})"
    if header == "GST 18%":
        return f"={excel_ref('Total Quarterly charges Gross', row, out_map)}*18%"
    if header == "Net Payable After Tax":
        return (
            f"={excel_ref('Total Quarterly charges Gross', row, out_map)}"
            f"+{excel_ref('GST 18%', row, out_map)}"
        )
    raise KeyError(header)


def title_text(source_title: Any, split_value: Any) -> str:
    text = str(source_title or "").strip()
    state = str(split_value).strip()
    # Avoid duplicating the state if already present at the end.
    if state and not re.search(rf"(?:-|\s){re.escape(state)}\s*$", text, flags=re.IGNORECASE):
        text = f"{text} - {state}" if text else f"GST STATE: {state}"
    return text


def style_output_sheet(ws, total_row: int) -> None:
    max_col = len(REQUIRED_OUTPUT_HEADERS)

    # Title row.
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = ws.cell(1, 1)
    title_cell.font = Font(name="Arial", size=15, bold=True, color=BLACK)
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_cell.border = MEDIUM_BORDER
    ws.row_dimensions[1].height = 34

    # Headers.
    for col, header in enumerate(REQUIRED_OUTPUT_HEADERS, start=1):
        cell = ws.cell(2, col)
        cell.value = header
        cell.font = Font(name="Arial", size=9, bold=True, color=BLACK)
        cell.fill = HEADER_FILL
        cell.border = MEDIUM_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = COLUMN_WIDTHS[header]
    ws.row_dimensions[2].height = 56

    # Body and totals.
    for row in range(3, total_row + 1):
        is_total = row == total_row
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Arial", size=9, bold=is_total, color=BLACK)
            cell.border = MEDIUM_BORDER if is_total else THIN_BORDER
            cell.alignment = Alignment(
                horizontal="left" if REQUIRED_OUTPUT_HEADERS[col - 1] == "Branch Name" else "center",
                vertical="center",
                wrap_text=True,
            )
        if not is_total:
            ws.row_dimensions[row].height = 22

    # Date formats.
    out_map = output_column_map()
    for header in ("Bill From", "Bill To", "PO Date"):
        col = out_map[header]
        for row in range(3, total_row):
            ws.cell(row, col).number_format = "dd-mmm-yyyy"

    # Amount formats.
    for header in TOTAL_HEADERS:
        col = out_map[header]
        for row in range(3, total_row + 1):
            ws.cell(row, col).number_format = '#,##0.00;[Red]-#,##0.00;"-"'

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{total_row - 1}"

    # Repeat title and header rows. Fit columns to one page, allow unlimited vertical pages.
    ws.print_title_rows = "1:2"
    ws.print_area = f"A1:{get_column_letter(max_col)}{total_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LEGAL
    ws.page_setup.scale = None
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.15
    ws.oddFooter.center.text = "Page &P of &N"
    ws.evenFooter.center.text = "Page &P of &N"
    ws.sheet_view.showGridLines = False


def split_excel_file(uploaded_bytes: bytes, original_name: str) -> tuple[dict[str, bytes], list[str]]:
    keep_vba = original_name.lower().endswith(".xlsm")
    source_wb = load_workbook(io.BytesIO(uploaded_bytes), data_only=False, keep_vba=keep_vba)
    warnings: list[str] = []

    try:
        source_ws = source_wb.active
        header_row = find_header_row(source_ws)
        source_map = build_source_header_map(source_ws, header_row)
        out_map = output_column_map()
        data_start_row = header_row + 1

        split_source_col = source_map["GST STATE"]
        unique_states: list[Any] = []
        seen: set[str] = set()
        for row in range(data_start_row, source_ws.max_row + 1):
            value = source_ws.cell(row, split_source_col).value
            text = str(value).strip() if value is not None else ""
            if not text or text.lower() in {"none", "nan", "total"}:
                continue
            if text not in seen:
                seen.add(text)
                unique_states.append(value)

        if not unique_states:
            raise ValueError("No GST STATE values were found below the header row.")

        # Assume the title is in the row directly above the header, otherwise use A1.
        title_row = max(1, header_row - 1)
        source_title = source_ws.cell(title_row, 1).value or source_ws.cell(1, 1).value

        output_files: dict[str, bytes] = {}
        used_names: set[str] = set()

        for state in unique_states:
            new_wb = Workbook()
            ws = new_wb.active
            ws.title = clean_filename(state)[:31]
            ws.cell(1, 1).value = title_text(source_title, state)

            for col, header in enumerate(REQUIRED_OUTPUT_HEADERS, start=1):
                ws.cell(2, col).value = header

            destination_row = 3
            serial = 1
            state_text = str(state).strip()

            for source_row in range(data_start_row, source_ws.max_row + 1):
                source_state = source_ws.cell(source_row, split_source_col).value
                if source_state is None or str(source_state).strip() != state_text:
                    continue

                # Skip source total/footer rows.
                first_nonempty = next(
                    (
                        str(source_ws.cell(source_row, col).value).strip()
                        for col in range(1, source_ws.max_column + 1)
                        if source_ws.cell(source_row, col).value not in (None, "")
                    ),
                    "",
                )
                if first_nonempty.lower() == "total":
                    continue

                ws.cell(destination_row, out_map["Sr. No."]).value = serial

                for header in REQUIRED_OUTPUT_HEADERS[1:]:
                    dst = ws.cell(destination_row, out_map[header])
                    src = source_ws.cell(source_row, source_map[header])
                    copy_cell_style(src, dst)

                    if header in FORMULA_HEADERS:
                        dst.value = data_formula(header, destination_row, out_map)
                    else:
                        # Copy values only for non-derived columns. This prevents hidden
                        # references to excluded columns from being carried into output.
                        dst.value = src.value if not (isinstance(src.value, str) and src.value.startswith("=")) else None

                serial += 1
                destination_row += 1

            if destination_row == 3:
                warnings.append(f"No data rows found for GST STATE {state_text}.")
                new_wb.close()
                continue

            total_row = destination_row
            ws.cell(total_row, out_map["Port BW"]).value = "Total"

            for header in TOTAL_HEADERS:
                col = out_map[header]
                letter = get_column_letter(col)
                ws.cell(total_row, col).value = f"=SUM({letter}3:{letter}{total_row - 1})"

            style_output_sheet(ws, total_row)

            out = io.BytesIO()
            new_wb.save(out)
            new_wb.close()
            filename = unique_filename(clean_filename(state), used_names, ".xlsx")
            output_files[filename] = out.getvalue()

        if not output_files:
            raise ValueError("No split Excel files were created.")
        return output_files, warnings
    finally:
        source_wb.close()


def find_libreoffice() -> str | None:
    for command in ("libreoffice", "soffice"):
        found = shutil.which(command)
        if found:
            return found
    for path in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if os.path.exists(path):
            return path
    return None


def convert_excels_to_pdfs(excel_files: dict[str, bytes]) -> tuple[dict[str, bytes], list[str]]:
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice is not installed. Add 'libreoffice-calc' to packages.txt on Streamlit Cloud."
        )

    pdf_files: dict[str, bytes] = {}
    errors: list[str] = []

    with tempfile.TemporaryDirectory(prefix="excel_pdf_header_based_") as temp_dir:
        root = Path(temp_dir)
        excel_dir = root / "excel"
        pdf_dir = root / "pdf"
        excel_dir.mkdir()
        pdf_dir.mkdir()

        for index, (filename, data) in enumerate(excel_files.items(), start=1):
            excel_path = excel_dir / filename
            excel_path.write_bytes(data)
            profile_dir = root / f"profile_{index}_{uuid.uuid4().hex}"
            profile_dir.mkdir()

            try:
                command = [
                    soffice,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                    "--convert-to",
                    "pdf:calc_pdf_Export",
                    "--outdir",
                    str(pdf_dir),
                    str(excel_path),
                ]
                result = subprocess.run(command, capture_output=True, text=True, timeout=180, check=False)
                expected = pdf_dir / f"{excel_path.stem}.pdf"
                if result.returncode != 0 or not expected.exists():
                    detail = (result.stderr or result.stdout or "Unknown LibreOffice error").strip()
                    raise RuntimeError(detail)
                pdf_files[expected.name] = expected.read_bytes()
            except Exception as exc:
                errors.append(f"{filename}: {exc}")

    return pdf_files, errors


def make_zip(files: dict[str, bytes], folder_name: str | None = None) -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, data in files.items():
            arcname = f"{folder_name}/{filename}" if folder_name else filename
            archive.writestr(arcname, data)
    return stream.getvalue()


def make_combined_zip(excels: dict[str, bytes], pdfs: dict[str, bytes]) -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, data in excels.items():
            archive.writestr(f"Excel/{filename}", data)
        for filename, data in pdfs.items():
            archive.writestr(f"PDF/{filename}", data)
    return stream.getvalue()


def initialize_state() -> None:
    for key, value in {
        "excel_files": {},
        "pdf_files": {},
        "conversion_errors": [],
    }.items():
        if key not in st.session_state:
            st.session_state[key] = value


def main() -> None:
    st.set_page_config(page_title="Excel Splitter & PDF Generator", page_icon="📊", layout="wide")
    initialize_state()

    st.title("Excel Splitter & PDF Generator")
    st.caption(
        "The output columns are selected by header name. GST STATE is used automatically for splitting. "
        "Dependent formulas are rebuilt after excluded columns are removed."
    )

    with st.expander("Fixed output columns", expanded=False):
        st.write(REQUIRED_OUTPUT_HEADERS)
        st.info("Excluded automatically: PO, Type, CGST 9%, SGST 9%, Remarks, Loopback and WAN IP.")

    uploaded_file = st.file_uploader("Upload the original Excel file", type=["xlsx", "xlsm"])
    generate_pdf = st.checkbox("Also generate PDF files", value=True)

    if st.button("Generate Files", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.error("Please upload the original Excel file.")
        else:
            try:
                with st.status("Processing workbook...", expanded=True) as status:
                    st.write("Matching required columns by header name...")
                    excel_files, warnings = split_excel_file(uploaded_file.getvalue(), uploaded_file.name)
                    st.session_state.excel_files = excel_files
                    st.session_state.pdf_files = {}
                    st.session_state.conversion_errors = []
                    st.write(f"Created {len(excel_files)} Excel file(s).")

                    if generate_pdf:
                        st.write("Converting Excel files to PDF...")
                        pdf_files, errors = convert_excels_to_pdfs(excel_files)
                        st.session_state.pdf_files = pdf_files
                        st.session_state.conversion_errors = errors
                        st.write(f"Created {len(pdf_files)} PDF file(s).")

                    for warning in warnings:
                        st.warning(warning)
                    status.update(label="Processing completed", state="complete")
            except Exception as exc:
                st.session_state.excel_files = {}
                st.session_state.pdf_files = {}
                st.session_state.conversion_errors = []
                st.error(str(exc))

    excel_files = st.session_state.excel_files
    pdf_files = st.session_state.pdf_files
    errors = st.session_state.conversion_errors

    if excel_files:
        st.divider()
        st.subheader("Download Results")
        c1, c2, c3 = st.columns(3)
        c1.metric("Excel files", len(excel_files))
        c2.metric("PDF files", len(pdf_files))
        c3.metric("PDF errors", len(errors))

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button(
                "Download Excel ZIP",
                data=make_zip(excel_files),
                file_name="Split_Excel_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with d2:
            if pdf_files:
                st.download_button(
                    "Download PDF ZIP",
                    data=make_zip(pdf_files),
                    file_name="Generated_PDF_Files.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
        with d3:
            if pdf_files:
                st.download_button(
                    "Download Excel + PDF ZIP",
                    data=make_combined_zip(excel_files, pdf_files),
                    file_name="Excel_and_PDF_Files.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

        if errors:
            st.warning("Some PDF files could not be generated.")
            st.code("\n".join(errors), language=None)

    st.divider()
    st.caption("Created for BSNL Maharashtra Circle | Header-based formula-safe version")


if __name__ == "__main__":
    main()
