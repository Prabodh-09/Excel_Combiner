
from collections import OrderedDict
from io import BytesIO, StringIO
from pathlib import Path
import csv
import zipfile

import streamlit as st
from openpyxl import Workbook, load_workbook


st.set_page_config(page_title="Excel Combiner", page_icon="📊")
st.title("📊 Excel Combiner")
st.write("Upload Excel files and combine worksheets with matching names.")

st.info(
    "Matching worksheets are combined by column position. "
    "Formatting and macros are not preserved. Formula cells are copied "
    "as formulas and are not calculated by Python."
)


def is_blank_row(row):
    return all(
        value is None or (isinstance(value, str) and not value.strip())
        for value in row
    )


def normalise_header(row):
    return tuple(
        value.strip() if isinstance(value, str) else value
        for value in row
    )


def combine_files(files):
    combined = OrderedDict()
    report = []
    processed = 0
    errors = 0

    for uploaded in files:
        filename = uploaded.name
        suffix = Path(filename).suffix.lower()

        if Path(filename).name.startswith("~$"):
            report.append({
                "Source file": filename,
                "Worksheet": "",
                "Status": "Skipped",
                "Details": "Temporary Excel lock file.",
            })
            continue

        if suffix not in {".xlsx", ".xlsm"}:
            report.append({
                "Source file": filename,
                "Worksheet": "",
                "Status": "Skipped",
                "Details": "Unsupported file type.",
            })
            continue

        try:
            file_bytes = uploaded.getvalue()

            if not zipfile.is_zipfile(BytesIO(file_bytes)):
                raise ValueError("Invalid Excel workbook.")

            source = load_workbook(
                BytesIO(file_bytes),
                read_only=True,
                data_only=False,
                keep_links=False,
            )
        except Exception as exc:
            errors += 1
            report.append({
                "Source file": filename,
                "Worksheet": "",
                "Status": "Error",
                "Details": str(exc),
            })
            continue

        had_error = False

        try:
            for ws in source.worksheets:
                try:
                    rows = (
                        tuple(row)
                        for row in ws.iter_rows(values_only=True)
                        if not is_blank_row(row)
                    )

                    header = next(rows, None)

                    if header is None:
                        report.append({
                            "Source file": filename,
                            "Worksheet": ws.title,
                            "Status": "Skipped",
                            "Details": "Empty worksheet.",
                        })
                        continue

                    header = normalise_header(header)
                    key = ws.title.casefold()

                    if key not in combined:
                        combined[key] = {
                            "name": ws.title,
                            "header": header,
                            "rows": [],
                        }

                    target = combined[key]
                    warning = ""

                    if header != target["header"]:
                        warning = (
                            "Header mismatch. Rows were appended by column "
                            "position; verify the column order."
                        )

                    added = 0

                    for row in rows:
                        row = tuple(row)

                        if normalise_header(row) == target["header"]:
                            continue

                        target["rows"].append(row)
                        added += 1

                    report.append({
                        "Source file": filename,
                        "Worksheet": ws.title,
                        "Status": "Warning" if warning else "OK",
                        "Details": warning or f"Added {added} data rows.",
                    })

                except Exception as exc:
                    had_error = True
                    report.append({
                        "Source file": filename,
                        "Worksheet": ws.title,
                        "Status": "Error",
                        "Details": str(exc),
                    })
        finally:
            source.close()

        if had_error:
            errors += 1
        else:
            processed += 1

    if not combined:
        return None, report, processed, errors, 0

    output = Workbook()
    output.remove(output.active)

    for sheet_name, data in combined.items():
        ws = output.create_sheet(title=data["name"][:31])
        ws.append(list(data["header"]))

        for row in data["rows"]:
            if ws.max_row >= 1_048_576:
                report.append({
                    "Source file": "combined_output.xlsx",
                    "Worksheet": data["name"],
                    "Status": "Warning",
                    "Details": "Excel worksheet row limit reached.",
                })
                break

            ws.append(list(row))

    buffer = BytesIO()
    output.save(buffer)
    output.close()

    return buffer.getvalue(), report, processed, errors, len(combined)


def report_csv(report):
    buffer = StringIO()
    fields = ["Source file", "Worksheet", "Status", "Details"]
    writer = csv.DictWriter(buffer, fieldnames=fields)
    writer.writeheader()
    writer.writerows(report)
    return buffer.getvalue().encode("utf-8-sig")


# Persistent state keeps the output available after Streamlit reruns.
if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "report_rows" not in st.session_state:
    st.session_state.report_rows = []
if "summary" not in st.session_state:
    st.session_state.summary = None


st.subheader("1. Upload Excel files")

try:
    files = st.file_uploader(
        "Select a folder or multiple Excel files",
        type=["xlsx", "xlsm"],
        accept_multiple_files="directory",
        help="Folder upload requires a compatible Streamlit version.",
    )
except (TypeError, ValueError):
    st.warning(
        "Folder upload is unavailable in this Streamlit version. "
        "Select multiple files instead."
    )
    files = st.file_uploader(
        "Select multiple Excel files",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

if files:
    st.success(f"{len(files)} file(s) selected.")

    with st.expander("View selected files"):
        for file in files:
            st.write(file.name)

    if st.button("Combine Excel Files", type="primary",
                 use_container_width=True):
        with st.spinner("Combining worksheets..."):
            result = combine_files(files)

        output_bytes, report, processed, errors, sheet_count = result

        st.session_state.output_bytes = output_bytes
        st.session_state.report_rows = report
        st.session_state.summary = {
            "processed": processed,
            "errors": errors,
            "sheets": sheet_count,
        }

        if output_bytes:
            st.success("Combination completed!")
        else:
            st.error("No non-empty worksheets were found.")

# IMPORTANT: download buttons are outside the Combine button condition.
# They remain visible after Streamlit reruns the application.
if st.session_state.output_bytes:
    st.subheader("2. Download your combined workbook")

    summary = st.session_state.summary
    col1, col2, col3 = st.columns(3)
    col1.metric("Workbooks processed", summary["processed"])
    col2.metric("Worksheets combined", summary["sheets"])
    col3.metric("Files with errors", summary["errors"])

    st.download_button(
        "⬇️ Download Combined Excel Workbook",
        data=st.session_state.output_bytes,
        file_name="combined_output.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_excel",
        use_container_width=True,
    )

if st.session_state.report_rows:
    st.subheader("3. Processing report")

    st.dataframe(
        st.session_state.report_rows,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "⬇️ Download Processing Report (CSV)",
        data=report_csv(st.session_state.report_rows),
        file_name="processing_report.csv",
        mime="text/csv",
        key="download_report",
        use_container_width=True,
    )

st.caption(
    "Matching worksheets are identified by name, ignoring capitalization. "
    "Ensure matching worksheets use the same column order."
)
