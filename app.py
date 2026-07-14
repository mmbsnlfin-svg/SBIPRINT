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

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import column_index_from_string, get_column_letter


# -----------------------------
# PRINT / STYLE SETTINGS
# -----------------------------
BLACK = "000000"
MEDIUM_SIDE = Side(border_style="medium", color=BLACK)
MEDIUM_BORDER = Border(
    left=MEDIUM_SIDE,
    right=MEDIUM_SIDE,
    top=MEDIUM_SIDE,
    bottom=MEDIUM_SIDE,
)

# Minimum widths from Column A onward. Auto-fit may increase these widths.
MIN_COLUMN_WIDTHS = [
    6, 12, 12, 12, 12, 7, 9, 25, 8, 10,
    15, 15, 15, 15, 15, 15, 15, 15, 10, 12, 11, 13,
]

DEFAULT_FONT_NAME = "Arial"
DEFAULT_DATA_FONT_SIZE = 10
DEFAULT_HEADER_FONT_SIZE = 10
MAX_AUTO_WIDTH = 35
MIN_AUTO_WIDTH = 6


# -----------------------------
# GENERAL HELPERS
# -----------------------------
def letter_to_index(col_letter: str) -> int:
    """Convert an Excel column letter to a 1-based column number."""
    value = col_letter.strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", value):
        raise ValueError(f"Invalid Excel column: {col_letter!r}")
    return column_index_from_string(value)


def clean_filename(value: object) -> str:
    """Create a safe, non-empty output filename."""
    text = str(value).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:100] or "Blank_Value"


def unique_filename(base_name: str, used_names: set[str], extension: str) -> str:
    """Avoid overwriting files when cleaned split values are identical."""
    candidate = f"{base_name}{extension}"
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{base_name}_{counter}{extension}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def copy_cell(src_cell, dst_cell) -> None:
    """Copy a cell value, style, hyperlink and comment."""
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)
    if src_cell.hyperlink:
        dst_cell._hyperlink = copy.copy(src_cell.hyperlink)
    if src_cell.comment:
        dst_cell.comment = copy.copy(src_cell.comment)


def copy_row(src_ws, dst_ws, src_row_idx: int, dst_row_idx: int, max_col: int) -> None:
    """Copy an entire row including styles and row height."""
    for col_idx in range(1, max_col + 1):
        copy_cell(
            src_ws.cell(row=src_row_idx, column=col_idx),
            dst_ws.cell(row=dst_row_idx, column=col_idx),
        )

    source_height = src_ws.row_dimensions[src_row_idx].height
    if source_height is not None:
        dst_ws.row_dimensions[dst_row_idx].height = source_height


def copy_column_dimensions(src_ws, dst_ws, max_col: int) -> None:
    """Copy original widths and hidden status."""
    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        src_dim = src_ws.column_dimensions.get(letter)
        if src_dim:
            if src_dim.width is not None:
                dst_ws.column_dimensions[letter].width = src_dim.width
            dst_ws.column_dimensions[letter].hidden = src_dim.hidden


def copy_header_merged_cells(src_ws, dst_ws) -> None:
    """Copy merged ranges fully contained within header rows 1 and 2."""
    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.min_row >= 1 and merged_range.max_row <= 2:
            dst_ws.merge_cells(str(merged_range))


def value_display_length(value: object) -> int:
    """Estimate visible character length for column auto-fit."""
    if value is None:
        return 0
    text = str(value)
    if "\n" in text:
        return max(len(part) for part in text.splitlines())
    return len(text)


# -----------------------------
# DARKER PRINT AND WIDTH FIXES
# -----------------------------
def make_font_dark(font: Font, *, bold: bool | None = None, size: float | None = None) -> Font:
    """Return an Arial, pure-black font while preserving useful font attributes."""
    return Font(
        name=DEFAULT_FONT_NAME,
        size=size if size is not None else (font.size or DEFAULT_DATA_FONT_SIZE),
        bold=font.bold if bold is None else bold,
        italic=font.italic,
        vertAlign=font.vertAlign,
        underline=font.underline,
        strike=font.strike,
        color=BLACK,
    )


def apply_dark_print_style(ws, max_row: int, max_col: int) -> None:
    """
    Make all printable data pure black and strengthen borders.
    Header rows and total row are bold.
    """
    for row_idx in range(1, max_row + 1):
        is_header = row_idx in (1, 2)
        is_total = row_idx == max_row

        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell):
                continue

            font_size = DEFAULT_HEADER_FONT_SIZE if is_header else DEFAULT_DATA_FONT_SIZE
            cell.font = make_font_dark(
                cell.font,
                bold=True if (is_header or is_total) else None,
                size=font_size,
            )

            # Use medium borders for clear printout.
            cell.border = MEDIUM_BORDER

            # Keep data readable; don't shrink text invisibly.
            old = cell.alignment
            cell.alignment = Alignment(
                horizontal=old.horizontal,
                vertical=old.vertical or "center",
                text_rotation=old.text_rotation,
                wrap_text=old.wrap_text,
                shrink_to_fit=False,
                indent=old.indent,
            )


def auto_fit_columns(ws, max_row: int, max_col: int) -> None:
    """
    Increase column widths according to displayed content.
    This prevents Excel/LibreOffice from showing ### for numbers and dates.
    """
    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        min_width = (
            MIN_COLUMN_WIDTHS[col_idx - 1]
            if col_idx <= len(MIN_COLUMN_WIDTHS)
            else MIN_AUTO_WIDTH
        )

        longest = 0
        numeric_or_date_seen = False

        for row_idx in range(1, max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell):
                continue
            longest = max(longest, value_display_length(cell.value))
            if isinstance(cell.value, (int, float)) or cell.is_date:
                numeric_or_date_seen = True

        padding = 4 if numeric_or_date_seen else 2
        calculated = longest + padding
        final_width = max(min_width, min(calculated, MAX_AUTO_WIDTH))
        ws.column_dimensions[letter].width = final_width


def wrap_header_row(ws, max_col: int) -> None:
    """Wrap and center Row 2 headers."""
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=2, column=col_idx)
        if isinstance(cell, MergedCell):
            continue
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
            shrink_to_fit=False,
        )
    ws.row_dimensions[2].height = 72


def set_data_alignment(ws, first_data_row: int, last_data_row: int, max_col: int) -> None:
    """Center short values and keep long descriptions wrapped."""
    for row_idx in range(first_data_row, last_data_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell):
                continue

            text_length = value_display_length(cell.value)
            horizontal = "left" if text_length > 18 else "center"
            cell.alignment = Alignment(
                horizontal=horizontal,
                vertical="center",
                wrap_text=text_length > 18,
                shrink_to_fit=False,
            )


# -----------------------------
# PAGE SETTINGS AND PDF
# -----------------------------
def apply_pdf_page_settings(workbook_path: Path, paper_size: str = "Legal") -> None:
    """Apply print settings before LibreOffice converts the workbook to PDF."""
    wb = load_workbook(workbook_path)
    try:
        for ws in wb.worksheets:
            ws.print_title_rows = "1:2"
            ws.page_setup.orientation = "landscape"

            if paper_size.upper() == "A3":
                ws.page_setup.paperSize = ws.PAPERSIZE_A3
            else:
                ws.page_setup.paperSize = ws.PAPERSIZE_LEGAL

            ws.page_margins.left = 0.15
            ws.page_margins.right = 0.15
            ws.page_margins.top = 0.20
            ws.page_margins.bottom = 0.20
            ws.page_margins.header = 0.10
            ws.page_margins.footer = 0.10

            # Fit every column on ONE page across. Do not fit the sheet vertically.
            # LibreOffice/Excel will automatically continue excess rows on page 2,
            # page 3, and so on, while repeating Rows 1 and 2 on every page.
            ws.page_setup.scale = None
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.sheet_view.showGridLines = False
            ws.oddFooter.center.text = "Page &P"
            ws.evenFooter.center.text = "Page &P"

            # Print only the used range.
            ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

        wb.save(workbook_path)
    finally:
        wb.close()


def find_libreoffice() -> str | None:
    """Find LibreOffice on Linux, macOS or Windows."""
    for command in ("libreoffice", "soffice"):
        found = shutil.which(command)
        if found:
            return found

    windows_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in windows_paths:
        if os.path.exists(path):
            return path
    return None


def convert_excels_to_pdfs(
    excel_files: dict[str, bytes],
    paper_size: str,
) -> tuple[dict[str, bytes], list[str]]:
    """Convert generated Excel files to PDF through LibreOffice."""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice is not installed. For Streamlit Community Cloud, "
            "add 'libreoffice-calc' to packages.txt and redeploy."
        )

    pdf_files: dict[str, bytes] = {}
    errors: list[str] = []

    with tempfile.TemporaryDirectory(prefix="excel_pdf_app_") as temp_dir:
        root = Path(temp_dir)
        input_dir = root / "excel"
        output_dir = root / "pdf"
        input_dir.mkdir()
        output_dir.mkdir()

        for index, (filename, data) in enumerate(excel_files.items(), start=1):
            excel_path = input_dir / filename
            excel_path.write_bytes(data)
            profile_dir = root / f"profile_{index}_{uuid.uuid4().hex}"
            profile_dir.mkdir()

            try:
                apply_pdf_page_settings(excel_path, paper_size=paper_size)
                profile_url = profile_dir.resolve().as_uri()
                command = [
                    soffice,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    f"-env:UserInstallation={profile_url}",
                    "--convert-to",
                    "pdf:calc_pdf_Export",
                    "--outdir",
                    str(output_dir),
                    str(excel_path),
                ]
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=180,
                    check=False,
                )
                expected_pdf = output_dir / f"{excel_path.stem}.pdf"
                if result.returncode != 0 or not expected_pdf.exists():
                    detail = (result.stderr or result.stdout or "Unknown conversion error").strip()
                    raise RuntimeError(detail)
                pdf_files[expected_pdf.name] = expected_pdf.read_bytes()
            except Exception as exc:
                errors.append(f"{filename}: {exc}")

    return pdf_files, errors


# -----------------------------
# EXCEL SPLITTING
# -----------------------------
def split_excel_file(
    uploaded_bytes: bytes,
    original_name: str,
    split_col_letter: str,
    total_from_letter: str,
    total_to_letter: str,
    auto_widths: bool,
) -> tuple[dict[str, bytes], list[str]]:
    split_col = letter_to_index(split_col_letter)
    total_start_col = letter_to_index(total_from_letter)
    total_end_col = letter_to_index(total_to_letter)

    if total_start_col > total_end_col:
        raise ValueError("Total Start Column must come before Total End Column.")

    keep_vba = original_name.lower().endswith(".xlsm")
    source_stream = io.BytesIO(uploaded_bytes)
    workbook = load_workbook(source_stream, data_only=False, keep_vba=keep_vba)

    try:
        source_ws = workbook.active
        max_row = source_ws.max_row
        max_col = max(source_ws.max_column, split_col, total_end_col)

        if max_row < 3:
            raise ValueError("The selected Excel file has fewer than 3 rows.")

        unique_values: list[object] = []
        seen_values: set[str] = set()

        for row_idx in range(3, max_row + 1):
            value = source_ws.cell(row=row_idx, column=split_col).value
            if value is None:
                continue
            value_text = str(value).strip()
            if not value_text or value_text.lower() in {"nan", "none"}:
                continue
            if value_text not in seen_values:
                seen_values.add(value_text)
                unique_values.append(value)

        if not unique_values:
            raise ValueError(
                f"No valid split values were found in column "
                f"{split_col_letter.upper()} from row 3 onward."
            )

        output_files: dict[str, bytes] = {}
        used_names: set[str] = set()
        warnings: list[str] = []
        total_label_col = max(1, total_start_col - 1)

        for split_value in unique_values:
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.title = source_ws.title

            copy_column_dimensions(source_ws, new_ws, max_col)
            copy_row(source_ws, new_ws, 1, 1, max_col)
            copy_row(source_ws, new_ws, 2, 2, max_col)
            copy_header_merged_cells(source_ws, new_ws)

            split_header = new_ws.cell(row=1, column=split_col)
            split_header.value = split_value
            split_header.font = make_font_dark(split_header.font, bold=True, size=11)
            split_header.border = MEDIUM_BORDER
            split_header.alignment = Alignment(horizontal="center", vertical="center")

            new_ws.cell(row=2, column=1).value = "Sr. No."
            wrap_header_row(new_ws, max_col)

            destination_row = 3
            serial_number = 1
            split_text = str(split_value).strip()

            for source_row in range(3, max_row + 1):
                source_value = source_ws.cell(row=source_row, column=split_col).value
                if source_value is not None and str(source_value).strip() == split_text:
                    copy_row(source_ws, new_ws, source_row, destination_row, max_col)
                    new_ws.cell(row=destination_row, column=1).value = serial_number
                    serial_number += 1
                    destination_row += 1

            total_row = destination_row
            new_ws.cell(row=total_row, column=total_label_col).value = "Total"

            for col_idx in range(total_start_col, total_end_col + 1):
                total_cell = new_ws.cell(row=total_row, column=col_idx)
                col_letter = get_column_letter(col_idx)
                total_cell.value = (
                    f"=SUM({col_letter}3:{col_letter}{total_row - 1})"
                    if total_row > 3
                    else 0
                )
                total_cell.number_format = source_ws.cell(row=3, column=col_idx).number_format

            # Improve visibility before saving.
            set_data_alignment(new_ws, 3, total_row, max_col)
            apply_dark_print_style(new_ws, total_row, max_col)

            if auto_widths:
                auto_fit_columns(new_ws, total_row, max_col)

            # Keep title and header rows visible when scrolling/printing.
            new_ws.freeze_panes = "A3"
            new_ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{total_row}"

            safe_name = clean_filename(split_value)
            filename = unique_filename(safe_name, used_names, ".xlsx")
            out_stream = io.BytesIO()
            new_wb.save(out_stream)
            new_wb.close()
            output_files[filename] = out_stream.getvalue()

        return output_files, warnings
    finally:
        workbook.close()


# -----------------------------
# ZIP AND STREAMLIT STATE
# -----------------------------
def make_zip(files: dict[str, bytes], folder_name: str | None = None) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, data in files.items():
            archive_name = f"{folder_name}/{filename}" if folder_name else filename
            archive.writestr(archive_name, data)
    return output.getvalue()


def make_combined_zip(excel_files: dict[str, bytes], pdf_files: dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, data in excel_files.items():
            archive.writestr(f"Excel/{filename}", data)
        for filename, data in pdf_files.items():
            archive.writestr(f"PDF/{filename}", data)
    return output.getvalue()


def initialize_state() -> None:
    defaults = {
        "excel_files": {},
        "pdf_files": {},
        "conversion_errors": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# -----------------------------
# STREAMLIT APP
# -----------------------------
def main() -> None:
    st.set_page_config(
        page_title="Excel Splitter & PDF Generator",
        page_icon="📊",
        layout="wide",
    )
    initialize_state()

    st.title("Excel Splitter & PDF Generator")
    st.caption(
        "Creates separate Excel and PDF files with dark print, stronger borders, "
        "one-page-wide printing, repeated headers, and automatic continuation of rows "
        "onto additional pages."
    )

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm"])

    col1, col2, col3 = st.columns(3)
    with col1:
        split_col = st.text_input("Split column", value="U", max_chars=3).upper()
    with col2:
        total_from = st.text_input("Total starts at column", value="D", max_chars=3).upper()
    with col3:
        total_to = st.text_input("Total ends at column", value="F", max_chars=3).upper()

    option1, option2, option3 = st.columns(3)
    with option1:
        auto_widths = st.checkbox(
            "Auto-fit columns to prevent ###",
            value=True,
        )
    with option2:
        generate_pdf_now = st.checkbox("Also generate PDF files", value=True)
    with option3:
        paper_size = st.selectbox("PDF paper size", ["Legal", "A3"], index=0)

    st.info(
        "Recommended to match the sample: Legal Landscape and Auto-fit ON. "
        "All columns are fitted on one page across; additional rows automatically "
        "continue on page 2 and further pages."
    )

    if st.button("Generate Files", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.error("Please upload an Excel file.")
        else:
            try:
                uploaded_bytes = uploaded_file.getvalue()
                with st.status("Processing workbook...", expanded=True) as status:
                    st.write("Reading the workbook and identifying split values...")
                    excel_files, warnings = split_excel_file(
                        uploaded_bytes=uploaded_bytes,
                        original_name=uploaded_file.name,
                        split_col_letter=split_col,
                        total_from_letter=total_from,
                        total_to_letter=total_to,
                        auto_widths=auto_widths,
                    )
                    st.session_state.excel_files = excel_files
                    st.session_state.pdf_files = {}
                    st.session_state.conversion_errors = []
                    st.write(f"Created {len(excel_files)} split Excel file(s).")

                    if generate_pdf_now:
                        st.write("Applying A3/Legal landscape settings and generating PDFs...")
                        pdf_files, errors = convert_excels_to_pdfs(
                            excel_files,
                            paper_size=paper_size,
                        )
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

        m1, m2, m3 = st.columns(3)
        m1.metric("Excel files", len(excel_files))
        m2.metric("PDF files", len(pdf_files))
        m3.metric("PDF errors", len(errors))

        download_col1, download_col2, download_col3 = st.columns(3)
        with download_col1:
            st.download_button(
                "Download Excel ZIP",
                data=make_zip(excel_files),
                file_name="Split_Excel_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with download_col2:
            if pdf_files:
                st.download_button(
                    "Download PDF ZIP",
                    data=make_zip(pdf_files),
                    file_name="Generated_PDF_Files.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
        with download_col3:
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

        with st.expander("Generated file names"):
            st.write("**Excel**")
            st.write(list(excel_files.keys()))
            if pdf_files:
                st.write("**PDF**")
                st.write(list(pdf_files.keys()))

    st.divider()
    st.caption("Created for BSNL Maharashtra Circle | Streamlit version")


if __name__ == "__main__":
    main()
