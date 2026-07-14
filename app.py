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
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import column_index_from_string, get_column_letter


BLACK = "000000"
THIN_SIDE = Side(border_style="thin", color=BLACK)
MEDIUM_SIDE = Side(border_style="medium", color=BLACK)
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MEDIUM_BORDER = Border(left=MEDIUM_SIDE, right=MEDIUM_SIDE, top=MEDIUM_SIDE, bottom=MEDIUM_SIDE)

# Preferred widths of source columns A onward. These widths are automatically
# scaled after unwanted columns are removed.
BASE_COLUMN_WIDTHS = [
    6, 12, 12, 12, 12, 5, 9, 25, 6, 9,
    15, 15, 15, 15, 15, 15, 15, 15, 8, 12, 9, 12,
]


def letter_to_index(col_letter: str) -> int:
    value = col_letter.strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", value):
        raise ValueError(f"Invalid Excel column: {col_letter!r}")
    return column_index_from_string(value)


def parse_skipped_columns(text: str) -> set[int]:
    """Convert A,B,C or A B C into a set of source column numbers."""
    if not text.strip():
        return set()

    parts = [part for part in re.split(r"[\s,;]+", text.strip().upper()) if part]
    skipped: set[int] = set()
    invalid: list[str] = []

    for part in parts:
        if re.fullmatch(r"[A-Z]{1,3}", part):
            skipped.add(column_index_from_string(part))
        else:
            invalid.append(part)

    if invalid:
        raise ValueError(
            "Invalid skipped column value(s): " + ", ".join(invalid) +
            ". Enter columns like A,B,C."
        )
    return skipped


def clean_filename(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:100] or "Blank_Value"


def unique_filename(base_name: str, used_names: set[str], extension: str) -> str:
    candidate = f"{base_name}{extension}"
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{base_name}_{counter}{extension}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def copy_cell(src_cell, dst_cell) -> None:
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


def copy_selected_row(
    src_ws,
    dst_ws,
    src_row_idx: int,
    dst_row_idx: int,
    visible_source_columns: list[int],
) -> None:
    for dst_col_idx, src_col_idx in enumerate(visible_source_columns, start=1):
        copy_cell(
            src_ws.cell(row=src_row_idx, column=src_col_idx),
            dst_ws.cell(row=dst_row_idx, column=dst_col_idx),
        )

    source_height = src_ws.row_dimensions[src_row_idx].height
    if source_height is not None:
        dst_ws.row_dimensions[dst_row_idx].height = source_height


def choose_font_sizes(visible_column_count: int) -> tuple[float, float, float]:
    """Return body, header and title font sizes based on remaining columns."""
    if visible_column_count <= 10:
        return 11.0, 11.0, 16.0
    if visible_column_count <= 14:
        return 10.0, 10.0, 15.0
    if visible_column_count <= 18:
        return 9.0, 9.0, 14.0
    if visible_column_count <= 22:
        return 8.0, 8.0, 13.0
    if visible_column_count <= 26:
        return 7.0, 7.0, 12.0
    return 6.5, 6.5, 11.0


def source_title_text(source_ws) -> str:
    """Get the main title from row 1, ignoring short state-code-like values."""
    candidates: list[str] = []
    for cell in source_ws[1]:
        if cell.value is None:
            continue
        text = str(cell.value).strip()
        if text:
            candidates.append(text)
    if not candidates:
        return "Statement"
    return max(candidates, key=len)


def apply_column_widths(
    source_ws,
    target_ws,
    visible_source_columns: list[int],
    body_font_size: float,
    data_last_row: int,
) -> None:
    """Set widths from headers/data only; row 1 title is deliberately ignored.

    This prevents the long merged title from making the Sr. No. or other columns
    unnecessarily wide. Widths are capped so all columns remain printable.
    """
    for dst_col, src_col in enumerate(visible_source_columns, start=1):
        header = target_ws.cell(row=2, column=dst_col).value
        header_len = max((len(part) for part in str(header or "").split()), default=0)

        max_len = header_len
        # Measure only actual table rows, never the title row.
        for row_idx in range(3, data_last_row + 1):
            value = target_ws.cell(row=row_idx, column=dst_col).value
            if value is None:
                continue
            text = str(value)
            # Numeric columns need enough room for formatted values but not unlimited width.
            max_len = max(max_len, min(len(text), 24))

        header_text = str(header or "").strip().lower()
        if dst_col == 1 or header_text in {"sr. no.", "sr no.", "sr.no."}:
            width = 5.5
        elif "branch name" in header_text or "remarks" in header_text:
            width = min(max(max_len + 1.5, 14), 24)
        elif any(k in header_text for k in ("charges", "payable", "cgst", "sgst", "gst 18")):
            width = min(max(max_len + 1.2, 11), 15)
        elif any(k in header_text for k in ("account", "lc id", "po no", "loopback", "wan ip")):
            width = min(max(max_len + 1.2, 10), 15)
        elif any(k in header_text for k in ("bill from", "bill to", "po date")):
            width = min(max(max_len + 1.0, 10), 12)
        else:
            width = min(max(max_len + 1.0, 6), 14)

        target_ws.column_dimensions[get_column_letter(dst_col)].width = width


def style_generated_sheet(
    ws,
    visible_column_count: int,
    body_font_size: float,
    header_font_size: float,
    title_font_size: float,
    total_row: int,
) -> None:
    """Use solid black print-friendly fonts and borders."""
    for row_idx in range(2, total_row + 1):
        for col_idx in range(1, visible_column_count + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            is_header = row_idx == 2
            is_total = row_idx == total_row
            old = cell.font
            cell.font = Font(
                name="Arial",
                size=header_font_size if is_header else body_font_size,
                bold=is_header or is_total,
                italic=old.italic,
                color=BLACK,
            )
            cell.border = MEDIUM_BORDER if is_header or is_total else THIN_BORDER

            if is_header:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )
            else:
                old_alignment = cell.alignment
                cell.alignment = Alignment(
                    horizontal=old_alignment.horizontal,
                    vertical="center",
                    wrap_text=old_alignment.wrap_text,
                    shrink_to_fit=old_alignment.shrink_to_fit,
                )

    ws.row_dimensions[1].height = max(38, title_font_size * 2.8)
    ws.row_dimensions[2].height = max(55, header_font_size * 7)


def build_repeating_title_row(
    ws,
    title_text: str,
    gst_state_code: object,
    visible_column_count: int,
    title_font_size: float,
) -> None:
    """Build one full-width repeating title row.

    The GST State Code is appended to the title so the entire heading, including
    the state code, is large and bold. This also avoids reserving wide blank
    columns at the right side of the report.
    """
    full_title = f"{title_text} - {gst_state_code}"
    if visible_column_count > 1:
        ws.merge_cells(
            start_row=1, start_column=1,
            end_row=1, end_column=visible_column_count
        )

    title_cell = ws.cell(row=1, column=1)
    title_cell.value = full_title
    title_cell.font = Font(
        name="Arial",
        size=max(title_font_size, 14),
        bold=True,
        color=BLACK,
    )
    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
        shrink_to_fit=False,
    )
    title_cell.border = MEDIUM_BORDER

    # Apply the outside/bottom border across the whole merged title range.
    for col_idx in range(1, visible_column_count + 1):
        ws.cell(row=1, column=col_idx).border = MEDIUM_BORDER


def apply_pdf_page_settings(workbook_path: Path) -> None:
    """Fit all columns on one page width; allow rows to flow to additional pages."""
    wb = load_workbook(workbook_path)
    try:
        for ws in wb.worksheets:
            ws.print_title_rows = "1:2"
            ws.page_setup.orientation = "landscape"
            ws.page_setup.paperSize = ws.PAPERSIZE_LEGAL
            ws.page_margins.left = 0.15
            ws.page_margins.right = 0.15
            ws.page_margins.top = 0.20
            ws.page_margins.bottom = 0.25
            ws.page_margins.header = 0.05
            ws.page_margins.footer = 0.10
            ws.page_setup.scale = None
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.sheet_properties.pageSetUpPr.autoPageBreaks = True
            ws.oddFooter.center.text = "Page &P of &N"
            ws.evenFooter.center.text = "Page &P of &N"
            ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        wb.save(workbook_path)
    finally:
        wb.close()


def split_excel_file(
    uploaded_bytes: bytes,
    original_name: str,
    split_col_letter: str,
    total_from_letter: str,
    total_to_letter: str,
    skipped_columns_text: str,
) -> tuple[dict[str, bytes], list[str]]:
    split_col = letter_to_index(split_col_letter)
    total_start_col = letter_to_index(total_from_letter)
    total_end_col = letter_to_index(total_to_letter)
    skipped_columns = parse_skipped_columns(skipped_columns_text)

    if total_start_col > total_end_col:
        raise ValueError("Total Start Column must come before Total End Column.")
    if split_col in skipped_columns:
        raise ValueError(
            f"Split column {split_col_letter.upper()} cannot also be skipped."
        )

    keep_vba = original_name.lower().endswith(".xlsm")
    workbook = load_workbook(io.BytesIO(uploaded_bytes), data_only=False, keep_vba=keep_vba)

    try:
        source_ws = workbook.active
        max_row = source_ws.max_row
        source_max_col = max(source_ws.max_column, split_col, total_end_col)

        if max_row < 3:
            raise ValueError("The selected Excel file has fewer than 3 rows.")

        visible_source_columns = [
            col for col in range(1, source_max_col + 1) if col not in skipped_columns
        ]
        if not visible_source_columns:
            raise ValueError("All columns have been skipped. Keep at least one column.")

        source_to_destination = {
            source_col: destination_col
            for destination_col, source_col in enumerate(visible_source_columns, start=1)
        }

        total_source_columns = [
            col for col in range(total_start_col, total_end_col + 1)
            if col in source_to_destination
        ]
        if not total_source_columns:
            raise ValueError(
                "All columns in the selected total range have been skipped."
            )

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

        visible_count = len(visible_source_columns)
        body_size, header_size, title_size = choose_font_sizes(visible_count)
        title_text = source_title_text(source_ws)

        output_files: dict[str, bytes] = {}
        used_names: set[str] = set()
        warnings: list[str] = []

        skipped_sorted = sorted(skipped_columns)
        if skipped_sorted:
            skipped_display = ", ".join(get_column_letter(col) for col in skipped_sorted)
            warnings.append(f"Skipped source columns: {skipped_display}")
        warnings.append(
            f"Remaining columns: {visible_count}; body font automatically set to {body_size} pt."
        )

        for split_value in unique_values:
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.title = source_ws.title

            # Row 2 is the original column-header row. Row 1 is rebuilt below.
            copy_selected_row(source_ws, new_ws, 2, 2, visible_source_columns)
            new_ws.cell(row=2, column=1).value = "Sr. No."

            destination_row = 3
            serial_number = 1
            split_text = str(split_value).strip()

            for source_row in range(3, max_row + 1):
                source_value = source_ws.cell(row=source_row, column=split_col).value
                if source_value is not None and str(source_value).strip() == split_text:
                    copy_selected_row(
                        source_ws,
                        new_ws,
                        source_row,
                        destination_row,
                        visible_source_columns,
                    )
                    new_ws.cell(row=destination_row, column=1).value = serial_number
                    serial_number += 1
                    destination_row += 1

            total_row = destination_row
            first_total_dest = source_to_destination[total_source_columns[0]]
            total_label_dest = max(1, first_total_dest - 1)
            new_ws.cell(row=total_row, column=total_label_dest).value = "Total"

            for source_col in total_source_columns:
                destination_col = source_to_destination[source_col]
                total_cell = new_ws.cell(row=total_row, column=destination_col)
                col_letter = get_column_letter(destination_col)
                total_cell.value = (
                    f"=SUM({col_letter}3:{col_letter}{total_row - 1})"
                    if total_row > 3 else 0
                )
                total_cell.number_format = source_ws.cell(row=3, column=source_col).number_format

            build_repeating_title_row(
                new_ws,
                title_text=title_text,
                gst_state_code=split_value,
                visible_column_count=visible_count,
                title_font_size=title_size,
            )
            apply_column_widths(
                source_ws,
                new_ws,
                visible_source_columns,
                body_font_size=body_size,
                data_last_row=total_row,
            )
            style_generated_sheet(
                new_ws,
                visible_column_count=visible_count,
                body_font_size=body_size,
                header_font_size=header_size,
                title_font_size=title_size,
                total_row=total_row,
            )

            new_ws.freeze_panes = "A3"
            new_ws.auto_filter.ref = f"A2:{get_column_letter(visible_count)}{max(2, total_row - 1)}"
            new_ws.print_title_rows = "1:2"
            new_ws.sheet_view.showGridLines = False

            safe_name = clean_filename(split_value)
            filename = unique_filename(safe_name, used_names, ".xlsx")
            out_stream = io.BytesIO()
            new_wb.save(out_stream)
            new_wb.close()
            output_files[filename] = out_stream.getvalue()

        return output_files, warnings
    finally:
        workbook.close()


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
                apply_pdf_page_settings(excel_path)
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


def make_zip(files: dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, data in files.items():
            archive.writestr(filename, data)
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


def main() -> None:
    st.set_page_config(
        page_title="Excel Splitter & PDF Generator",
        page_icon="📊",
        layout="wide",
    )
    initialize_state()

    st.title("Excel Splitter & PDF Generator")
    st.caption(
        "Split by GST State Code, remove unwanted columns, and generate print-ready Excel/PDF files."
    )

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm"])

    top1, top2 = st.columns([1, 2])
    with top1:
        split_col = st.text_input(
            "GST State Code / split column",
            value="U",
            max_chars=3,
            help="The values in this column create separate files, for example AP, AS, MH.",
        ).upper()
    with top2:
        skipped_columns = st.text_input(
            "Columns not required — enter source column letters",
            value="",
            placeholder="Example: A,B,C",
            help="These columns will be completely removed before Excel and PDF generation.",
        ).upper()

    col1, col2 = st.columns(2)
    with col1:
        total_from = st.text_input("Total starts at source column", value="D", max_chars=3).upper()
    with col2:
        total_to = st.text_input("Total ends at source column", value="F", max_chars=3).upper()

    generate_pdf_now = st.checkbox("Also generate PDF files", value=True)

    st.info(
        "The app automatically increases the font when columns are removed. "
        "Row 1 contains the main title and a large bold GST State Code, and rows 1–2 repeat on every PDF page."
    )

    if st.button("Generate Files", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.error("Please upload an Excel file.")
        else:
            try:
                with st.status("Processing workbook...", expanded=True) as status:
                    st.write("Reading the workbook and removing unwanted columns...")
                    excel_files, warnings = split_excel_file(
                        uploaded_bytes=uploaded_file.getvalue(),
                        original_name=uploaded_file.name,
                        split_col_letter=split_col,
                        total_from_letter=total_from,
                        total_to_letter=total_to,
                        skipped_columns_text=skipped_columns,
                    )
                    st.session_state.excel_files = excel_files
                    st.session_state.pdf_files = {}
                    st.session_state.conversion_errors = []
                    st.write(f"Created {len(excel_files)} split Excel file(s).")

                    if generate_pdf_now:
                        st.write("Converting Excel files to multi-page PDFs...")
                        pdf_files, errors = convert_excels_to_pdfs(excel_files)
                        st.session_state.pdf_files = pdf_files
                        st.session_state.conversion_errors = errors
                        st.write(f"Created {len(pdf_files)} PDF file(s).")

                    for warning in warnings:
                        st.info(warning)
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
