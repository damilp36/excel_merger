"""Streamlit entry point for Lookup Studio."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
import re
from typing import Any

import pandas as pd
import streamlit as st

from excel_merger.excel_io import LoadedSheet, WorkbookReadError, list_sheet_names, stream_sheet
from excel_merger.exporter import ExportLimitError, build_output_workbook, clean_output_filename
from excel_merger.final_result import (
    FinalResultDetails,
    GradeScale,
    build_final_result_workbook,
    prepare_final_result,
)
from excel_merger.merge import LookupSpec, MatchOptions, default_prefix, merge_lookups
from excel_merger.quality import (
    QualityReport,
    analyze_quality,
    arrow_safe_preview,
    is_missing,
    style_preview,
)


FINAL_RESULT_CACHE_VERSION = 2


st.set_page_config(
    page_title="Lookup Studio",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="collapsed",
)


@st.cache_data(show_spinner=False, max_entries=32)
def cached_sheet_names(content: bytes, file_name: str) -> list[str]:
    return list_sheet_names(content, file_name)


@st.cache_data(show_spinner=False, max_entries=32)
def cached_sheet(
    content: bytes,
    file_name: str,
    sheet_name: str,
    header_row: int,
) -> LoadedSheet:
    return stream_sheet(content, file_name, sheet_name, header_row)


@st.cache_data(show_spinner=False, max_entries=32)
def cached_quality(loaded: LoadedSheet) -> QualityReport:
    return analyze_quality(loaded)


def file_fingerprint(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()[:12]


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
          --ink: #172033;
          --muted: #64748b;
          --navy: #17324d;
          --blue: #2563eb;
          --surface: #ffffff;
          --line: #dce4ec;
        }
        .stApp { background: #f5f7fb; color: var(--ink); }
        .block-container { max-width: 1280px; padding-top: 2rem; padding-bottom: 4rem; }
        .hero {
          padding: 2.25rem 2.5rem;
          border-radius: 22px;
          color: white;
          background:
            radial-gradient(circle at 88% 18%, rgba(96,165,250,.34), transparent 28%),
            linear-gradient(125deg, #102a43 0%, #173f67 66%, #1d4f85 100%);
          box-shadow: 0 18px 50px rgba(23, 50, 77, .18);
          margin-bottom: 1.5rem;
        }
        .eyebrow { color: #bfdbfe; font-weight: 700; letter-spacing: .13em; font-size: .76rem; }
        .hero h1 { color: white; font-size: 2.45rem; line-height: 1.08; margin: .45rem 0 .65rem; }
        .hero p { color: #dbeafe; font-size: 1.05rem; max-width: 760px; margin: 0; }
        .step-label { color: var(--blue); font-size: .76rem; font-weight: 800; letter-spacing: .12em; }
        .legend { color: var(--muted); font-size: .86rem; margin: .25rem 0 .8rem; }
        .legend span { display: inline-block; margin-right: 1rem; }
        .swatch { width: .72rem; height: .72rem; border-radius: 3px; margin-right: .28rem; }
        div[data-testid="stMetric"] {
          background: white; border: 1px solid var(--line); border-radius: 14px; padding: .8rem 1rem;
        }
        div[data-testid="stFileUploader"] {
          background: #fbfdff; border: 1px dashed #aac3dd; border-radius: 14px; padding: .35rem .75rem;
        }
        div[data-testid="stTabs"] button { font-weight: 650; }
        div[data-testid="stAlert"] { border-radius: 12px; }
        .stButton > button[kind="primary"], .stDownloadButton > button {
          border-radius: 10px; min-height: 2.9rem; font-weight: 700;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def upload_slots(file_count: int) -> list[dict[str, Any] | None]:
    uploads: list[dict[str, Any] | None] = []
    for index in range(file_count):
        role = "Base table" if index == 0 else f"Lookup source {index}"
        with st.container(border=True):
            st.caption(role.upper())
            uploaded = st.file_uploader(
                f"{index + 1}. Upload {role.lower()}",
                type=["xlsx", "xlsm", "xls"],
                key=f"upload_slot_{index}",
                help="Excel .xlsx, .xlsm, and legacy .xls files are accepted.",
            )
            if uploaded is None:
                uploads.append(None)
                continue
            content = uploaded.getvalue()
            st.caption(f"{uploaded.name} · {len(content) / 1_048_576:.1f} MB")
            uploads.append(
                {
                    "name": uploaded.name,
                    "content": content,
                    "fingerprint": file_fingerprint(content),
                    "index": index,
                }
            )
    return uploads


def render_quality_summary(report: QualityReport) -> None:
    columns = st.columns(4)
    columns[0].metric("Rows", f"{report.row_count:,}")
    columns[1].metric("Columns", f"{report.column_count:,}")
    columns[2].metric("Missing cells", f"{report.missing_cells:,}")
    columns[3].metric("Mixed-type columns", f"{report.mixed_type_columns:,}")

    if report.warnings:
        st.warning(" · ".join(report.warnings), icon="⚠️")
    else:
        st.success("No obvious structural or cell-quality issues detected.", icon="✅")

    with st.expander("Column quality details"):
        quality_frame = report.columns_frame()
        if not quality_frame.empty:
            st.dataframe(
                quality_frame,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Missing %": st.column_config.ProgressColumn(
                        "Missing %", min_value=0.0, max_value=1.0, format="percent"
                    )
                },
            )
        if report.header_changes:
            st.markdown("**Header repairs**")
            for change in report.header_changes:
                st.caption(f"• {change}")
        if report.blank_rows_removed or report.blank_columns_removed:
            st.caption(
                f"Removed {report.blank_rows_removed:,} trailing blank rows and "
                f"{report.blank_columns_removed:,} unnamed blank columns."
            )


def render_file_tab(record: dict[str, Any], is_base: bool) -> dict[str, Any] | None:
    name = str(record["name"])
    content = record["content"]
    fingerprint = str(record["fingerprint"])
    role = "BASE TABLE" if is_base else "LOOKUP SOURCE"
    st.caption(role)

    try:
        sheet_names = cached_sheet_names(content, name)
    except WorkbookReadError as exc:
        st.error(str(exc))
        return None

    selector, header_control = st.columns([2, 1])
    with selector:
        sheet_name = st.selectbox(
            "Worksheet",
            sheet_names,
            key=f"sheet_{record['index']}_{fingerprint}",
        )
    with header_control:
        header_row = int(
            st.number_input(
                "Header row",
                min_value=1,
                max_value=100,
                value=1,
                step=1,
                key=f"header_{record['index']}_{fingerprint}_{sheet_name}",
                help="Use this when titles or notes appear above the actual column headers.",
            )
        )

    try:
        with st.spinner(f"Streaming cells from {sheet_name}…"):
            loaded = cached_sheet(content, name, sheet_name, header_row)
            report = cached_quality(loaded)
    except WorkbookReadError as exc:
        st.error(str(exc))
        return None

    dataframe = loaded.dataframe
    if dataframe.empty or not len(dataframe.columns):
        st.error("This worksheet has no data rows or usable columns below the selected header.")
        return None

    control_left, control_right = st.columns([1, 2])
    with control_left:
        key_column = st.selectbox(
            "Base lookup column" if is_base else "Lookup column",
            list(dataframe.columns),
            key=f"key_{record['index']}_{fingerprint}_{sheet_name}_{header_row}",
            help=(
                "Values in this column will be matched against every lookup source."
                if is_base
                else "This column will be matched to the base lookup column."
            ),
        )

    return_columns: list[str] = []
    prefix = ""
    with control_right:
        if is_base:
            st.info(
                "Every base row is retained. The lookup column is always included in the output.",
                icon="ℹ️",
            )
            options = [column for column in dataframe.columns if column != key_column]
            return_columns = st.multiselect(
                "Columns to return from the base table",
                options,
                default=options,
                key=(
                    f"base_returns_{record['index']}_{fingerprint}_{sheet_name}_"
                    f"{header_row}_{key_column}"
                ),
                help=(
                    "Choose which base fields to keep. The selected lookup column is "
                    "included automatically."
                ),
            )
        else:
            options = [column for column in dataframe.columns if column != key_column]
            return_columns = st.multiselect(
                "Columns to return",
                options,
                default=options,
                key=(
                    f"returns_{record['index']}_{fingerprint}_{sheet_name}_"
                    f"{header_row}_{key_column}"
                ),
                help="Selected fields are added to the merged output.",
            )
            prefix = st.text_input(
                "Output column prefix",
                value=default_prefix(name),
                key=f"prefix_{record['index']}_{fingerprint}",
                help="A prefix prevents column-name collisions and shows where returned data came from.",
            )

    duplicate_keys = int(
        (dataframe[key_column].notna() & dataframe[key_column].duplicated(keep=False)).sum()
    )
    if duplicate_keys:
        st.caption(
            f"{duplicate_keys:,} rows share a duplicate value in the selected key column. "
            "They are marked with a purple edge in the preview."
        )

    st.markdown(
        """
        <div class="legend">
          <span><span class="swatch" style="background:#fee2e2"></span>Missing</span>
          <span><span class="swatch" style="background:#fef3c7"></span>Different from the column's likely type</span>
          <span><span class="swatch" style="background:#8b5cf6"></span>Duplicate selected key</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.dataframe(
        style_preview(dataframe, report, key_column=key_column),
        use_container_width=True,
        height=390,
    )
    if len(dataframe) > 200:
        st.caption(f"Previewing the first 200 of {len(dataframe):,} rows.")
    render_quality_summary(report)

    return {
        "loaded": loaded,
        "report": report,
        "key_column": key_column,
        "return_columns": return_columns,
        "prefix": prefix,
    }


def generation_fingerprint(
    records: list[dict[str, Any]],
    configurations: list[dict[str, Any]],
    options: MatchOptions,
) -> str:
    payload = {
        "files": [record["fingerprint"] for record in records],
        "configurations": [
            {
                "sheet": config["loaded"].sheet_name,
                "header": config["loaded"].header_row,
                "key": config["key_column"],
                "returns": config["return_columns"],
                "prefix": config["prefix"],
            }
            for config in configurations
        ],
        "options": options.__dict__,
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()


def merged_style(dataframe: pd.DataFrame) -> pd.io.formats.style.Styler:
    preview = arrow_safe_preview(dataframe.head(200))
    styles = pd.DataFrame("", index=preview.index, columns=preview.columns)
    for row_index in preview.index:
        for column in preview.columns:
            value = preview.at[row_index, column]
            if is_missing(value):
                styles.at[row_index, column] = "background-color: #fee2e2; color: #991b1b"
            elif str(column).endswith(".Match status"):
                if value == "Matched":
                    styles.at[row_index, column] = "background-color: #dcfce7; color: #166534"
                else:
                    styles.at[row_index, column] = "background-color: #fef3c7; color: #92400e"
    return preview.style.apply(lambda _: styles, axis=None).format(na_rep="—")


def render_result(generated: dict[str, Any], output_name: str) -> None:
    result = generated["result"]
    st.markdown("### Your merged workbook is ready")
    metrics = st.columns(4)
    metrics[0].metric("Base rows", f"{result.base_rows:,}")
    metrics[1].metric("Output rows", f"{len(result.dataframe):,}")
    metrics[2].metric("Output columns", f"{len(result.dataframe.columns):,}")
    matched, unmatched = result.matched_cells_summary
    coverage = matched / (matched + unmatched) if matched + unmatched else 0.0
    metrics[3].metric("Lookup coverage", f"{coverage:.1%}")

    st.dataframe(
        merged_style(result.dataframe),
        use_container_width=True,
        height=390,
    )
    if len(result.dataframe) > 200:
        st.caption(f"Previewing the first 200 of {len(result.dataframe):,} output rows.")

    with st.expander("Lookup audit", expanded=True):
        st.dataframe(
            result.audit_frame(),
            use_container_width=True,
            hide_index=True,
        )

    st.download_button(
        "Download Excel workbook",
        data=generated["workbook"],
        file_name=clean_output_filename(output_name),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        help="Includes Merged Data, Lookup Audit, and Data Quality worksheets.",
    )

    st.markdown("### Prepare the final result")
    st.caption(
        "Continue with the merged data to calculate totals, grades, grade points, "
        "status, and a course summary."
    )
    if st.button("Proceed to final result", use_container_width=True):
        st.session_state["active_page"] = "final_result"
        st.session_state.pop("generated_final_result", None)
        st.rerun()


def _normalized_column_name(column: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(column).casefold()).strip()


def _suggest_column(columns: list[str], phrases: list[str], fallback: int = 0) -> str:
    if not columns:
        raise ValueError("No columns are available for final-result processing.")
    normalized = {column: _normalized_column_name(column) for column in columns}
    for phrase in phrases:
        expected = _normalized_column_name(phrase)
        for column in columns:
            candidate = normalized[column]
            matches = (
                expected in candidate
                if " " in expected
                else expected in candidate.split()
            )
            if expected and matches:
                return column
    return columns[min(fallback, len(columns) - 1)]


def _final_result_fingerprint(
    source_fingerprint: str,
    identifier_column: str,
    ca_column: str,
    exam_column: str,
    details: FinalResultDetails,
) -> str:
    payload = {
        "cache_version": FINAL_RESULT_CACHE_VERSION,
        "source": source_fingerprint,
        "identifier": identifier_column,
        "ca": ca_column,
        "exam": exam_column,
        "details": asdict(details),
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()


def final_result_style(dataframe: pd.DataFrame) -> pd.io.formats.style.Styler:
    preview = arrow_safe_preview(dataframe.head(200))
    styles = pd.DataFrame("", index=preview.index, columns=preview.columns)
    for row_index in preview.index:
        for column in preview.columns:
            value = preview.at[row_index, column]
            if is_missing(value):
                styles.at[row_index, column] = "background-color: #fee2e2; color: #991b1b"
            elif column == "Student Status":
                if value == "Pass":
                    styles.at[row_index, column] = "background-color: #dcfce7; color: #166534"
                elif value == "Fail":
                    styles.at[row_index, column] = "background-color: #fee2e2; color: #991b1b"
                else:
                    styles.at[row_index, column] = "background-color: #fef3c7; color: #92400e"
    return preview.style.apply(lambda _: styles, axis=None).format(na_rep="—")


def render_final_result_page() -> None:
    generated = st.session_state.get("generated_lookup")
    if not generated:
        st.error("Create a merged workbook before preparing a final result.")
        if st.button("Return to lookup"):
            st.session_state["active_page"] = "lookup"
            st.rerun()
        return

    if st.button("Back to lookup and merge"):
        st.session_state["active_page"] = "lookup"
        st.rerun()

    st.markdown(
        """
        <section class="hero">
          <div class="eyebrow">FINAL RESULT</div>
          <h1>Turn the merged scores into a complete mark sheet.</h1>
          <p>Choose the score columns, enter the course details, review the grading
          rules, and create an editable final-result workbook.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    merged_dataframe = generated["result"].dataframe
    source_fingerprint = str(generated.get("fingerprint", "current"))
    key_suffix = source_fingerprint[:12]
    columns = [str(column) for column in merged_dataframe.columns]
    score_columns = [column for column in columns if not column.endswith(".Match status")]

    st.markdown('<div class="step-label">STEP 1 · MAP RESULT COLUMNS</div>', unsafe_allow_html=True)
    st.subheader("Choose the identifier and score fields")
    identifier_suggestion = _suggest_column(
        columns,
        ["matric number", "student number", "student id", "identifier", "id"],
    )
    ca_suggestion = _suggest_column(
        score_columns,
        ["ca score", "continuous assessment", "test score", "assessment score", "ca", "test"],
        fallback=1,
    )
    exam_suggestion = _suggest_column(
        score_columns,
        ["exam score", "examination score", "final exam", "exam"],
        fallback=2,
    )
    mapping_columns = st.columns(3)
    with mapping_columns[0]:
        identifier_column = st.selectbox(
            "Student identifier column",
            columns,
            index=columns.index(identifier_suggestion),
            key=f"final_identifier_{key_suffix}",
        )
    with mapping_columns[1]:
        ca_column = st.selectbox(
            "CA score column",
            score_columns,
            index=score_columns.index(ca_suggestion),
            key=f"final_ca_{key_suffix}",
        )
    with mapping_columns[2]:
        exam_column = st.selectbox(
            "Exam score column",
            score_columns,
            index=score_columns.index(exam_suggestion),
            key=f"final_exam_{key_suffix}",
        )
    st.caption(
        "CA and Exam are suggested from column names only. The base file is not "
        "automatically treated as the Exam file. Confirm both selections before "
        "creating the final result."
    )

    st.markdown('<div class="step-label">STEP 2 · ENTER COURSE DETAILS</div>', unsafe_allow_html=True)
    st.subheader("Add the information that belongs on the mark sheet")
    with st.container(border=True):
        detail_columns = st.columns(3)
        with detail_columns[0]:
            programme = st.text_input("Programme", key=f"programme_{key_suffix}")
            session = st.text_input(
                "Session",
                placeholder="For example, 2025/2026",
                key=f"session_{key_suffix}",
            )
            semester = st.selectbox(
                "Semester",
                ["Harmattan", "Rain"],
                key=f"semester_{key_suffix}",
            )
        with detail_columns[1]:
            course_code = st.text_input("Course code", key=f"course_code_{key_suffix}")
            course_title = st.text_input("Course title", key=f"course_title_{key_suffix}")
            credit_units = int(
                st.number_input(
                    "Credit units",
                    min_value=0,
                    max_value=30,
                    value=2,
                    step=1,
                    key=f"credit_units_{key_suffix}",
                )
            )
        with detail_columns[2]:
            course_status = st.selectbox(
                "Course status",
                ["Core", "Elective", "Required"],
                key=f"course_status_{key_suffix}",
            )
            lecturers = st.text_input("Lecturer(s)", key=f"lecturers_{key_suffix}")
            identifier_heading = st.text_input(
                "Identifier heading in the workbook",
                value=str(identifier_column),
                key=f"identifier_heading_{key_suffix}_{identifier_column}",
            )
        remarks = st.text_area(
            "Lecturer's remarks",
            key=f"remarks_{key_suffix}",
            placeholder="Optional",
        )

    st.markdown('<div class="step-label">STEP 3 · REVIEW SCORE AND GRADE RULES</div>', unsafe_allow_html=True)
    st.subheader("Set the maximum scores and grading scale")
    score_settings = st.columns(3)
    maximum_ca_score = float(
        score_settings[0].number_input(
            "Maximum CA score",
            min_value=0.0,
            max_value=1_000.0,
            value=30.0,
            step=1.0,
            key=f"max_ca_{key_suffix}",
        )
    )
    maximum_exam_score = float(
        score_settings[1].number_input(
            "Maximum exam score",
            min_value=0.0,
            max_value=1_000.0,
            value=70.0,
            step=1.0,
            key=f"max_exam_{key_suffix}",
        )
    )
    pass_mark = float(
        score_settings[2].number_input(
            "Pass mark",
            min_value=0.0,
            max_value=1_000.0,
            value=40.0,
            step=1.0,
            key=f"pass_mark_{key_suffix}",
        )
    )

    with st.expander("Grade boundaries", expanded=False):
        st.caption("Enter the minimum grand total required for each grade.")
        grade_columns = st.columns(5)
        threshold_defaults = [70.0, 60.0, 50.0, 45.0, 40.0]
        grade_letters = list("ABCDE")
        thresholds = [
            float(
                grade_columns[index].number_input(
                    f"{grade} minimum",
                    min_value=0.0,
                    max_value=1_000.0,
                    value=threshold_defaults[index],
                    step=1.0,
                    key=f"grade_{grade.lower()}_{key_suffix}",
                )
            )
            for index, grade in enumerate(grade_letters)
        ]

    scale = GradeScale(
        a_minimum=thresholds[0],
        b_minimum=thresholds[1],
        c_minimum=thresholds[2],
        d_minimum=thresholds[3],
        e_minimum=thresholds[4],
        pass_mark=pass_mark,
    )
    details = FinalResultDetails(
        programme=programme,
        session=session,
        semester=semester,
        course_code=course_code,
        course_title=course_title,
        credit_units=credit_units,
        course_status=course_status,
        lecturers=lecturers,
        remarks=remarks,
        identifier_heading=identifier_heading,
        maximum_ca_score=maximum_ca_score,
        maximum_exam_score=maximum_exam_score,
        grade_scale=scale,
    )
    validation_errors = details.validation_errors()
    if len({identifier_column, ca_column, exam_column}) != 3:
        validation_errors.append(
            "Student identifier, CA score, and exam score must use different columns."
        )
    for error in validation_errors:
        st.warning(error)

    st.markdown('<div class="step-label">STEP 4 · CREATE FINAL RESULT</div>', unsafe_allow_html=True)
    final_controls = st.columns([2, 1])
    with final_controls[0]:
        st.caption(
            "The final workbook includes an executive summary, grade chart, result table, "
            "and an editable Settings worksheet."
        )
    with final_controls[1]:
        final_filename = st.text_input(
            "Final-result filename",
            value="final_result.xlsx",
            key=f"final_filename_{key_suffix}",
        )

    fingerprint = _final_result_fingerprint(
        source_fingerprint,
        identifier_column,
        ca_column,
        exam_column,
        details,
    )
    if st.button(
        "Create final result workbook",
        type="primary",
        use_container_width=True,
        disabled=bool(validation_errors),
    ):
        try:
            with st.spinner("Calculating grades and building the final workbook…"):
                processed = prepare_final_result(
                    merged_dataframe,
                    identifier_column,
                    ca_column,
                    exam_column,
                    details,
                )
                workbook = build_final_result_workbook(processed, details)
            st.session_state["generated_final_result"] = {
                "fingerprint": fingerprint,
                "processed": processed,
                "workbook": workbook,
            }
        except ValueError as exc:
            st.error(str(exc))

    final_generated = st.session_state.get("generated_final_result")
    if final_generated and final_generated.get("fingerprint") == fingerprint:
        processed = final_generated["processed"]
        missing_ca_count = int(getattr(processed, "missing_ca_count", 0))
        missing_exam_count = int(getattr(processed, "missing_exam_count", 0))
        st.markdown("### Final result preview")
        metrics = st.columns(4)
        metrics[0].metric("Records", f"{processed.total_count:,}")
        metrics[1].metric("Passes", f"{processed.pass_count:,}")
        metrics[2].metric("Fails", f"{processed.fail_count:,}")
        metrics[3].metric("Incomplete", f"{processed.incomplete_count:,}")
        if missing_ca_count or missing_exam_count:
            st.info(
                f"{missing_ca_count:,} blank CA values and "
                f"{missing_exam_count:,} blank exam values were changed to 0. "
                "The Student Status column identifies the missing score."
            )
        if processed.invalid_ca_count or processed.invalid_exam_count:
            st.warning(
                f"{processed.invalid_ca_count:,} CA values and "
                f"{processed.invalid_exam_count:,} exam values were outside the allowed "
                "range or were not numeric. They are marked incomplete."
            )
        st.dataframe(
            final_result_style(processed.dataframe),
            use_container_width=True,
            height=420,
        )
        grade_frame = pd.DataFrame(
            {
                "Grade": list(processed.grade_counts),
                "Count": list(processed.grade_counts.values()),
            }
        ).set_index("Grade")
        st.bar_chart(grade_frame, height=260)
        st.download_button(
            "Download final result workbook",
            data=final_generated["workbook"],
            file_name=clean_output_filename(final_filename),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    elif final_generated:
        st.info("The final-result settings changed. Create the workbook again to refresh it.")


def main() -> None:
    inject_styles()
    if st.session_state.get("active_page") == "final_result":
        render_final_result_page()
        return
    st.markdown(
        """
        <section class="hero">
          <div class="eyebrow">LOOKUP STUDIO</div>
          <h1>Bring scattered Excel data together.</h1>
          <p>Match a base table against multiple workbooks, inspect data-quality issues,
          and download one traceable result—without writing a formula.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="step-label">STEP 1 · ADD FILES</div>', unsafe_allow_html=True)
    heading, count_control = st.columns([3, 1])
    with heading:
        st.subheader("Choose how many workbooks to connect")
        st.caption("File 1 is the base table. Every later file supplies matching columns.")
    with count_control:
        file_count = int(
            st.number_input(
                "Number of files",
                min_value=2,
                max_value=10,
                value=2,
                step=1,
            )
        )

    uploads = upload_slots(file_count)
    uploaded_records = [record for record in uploads if record is not None]
    missing_count = file_count - len(uploaded_records)
    if not uploaded_records:
        st.info("Upload at least two Excel files to begin.", icon="ℹ️")
        return
    if missing_count:
        st.info(
            f"Upload {missing_count} more {'file' if missing_count == 1 else 'files'} to enable the merge. "
            "You can still inspect files already added.",
            icon="ℹ️",
        )

    st.markdown('<div class="step-label">STEP 2 · MAP COLUMNS</div>', unsafe_allow_html=True)
    st.subheader("Inspect each workbook and define the lookup")
    st.caption(
        "Files are read row by row. The preview highlights missing cells, mixed data types, and duplicate keys."
    )

    labels = [
        f"{record['index'] + 1} · {'Base' if record['index'] == 0 else Path(record['name']).stem}"
        for record in uploaded_records
    ]
    tabs = st.tabs(labels)
    config_by_slot: dict[int, dict[str, Any]] = {}
    for tab, record in zip(tabs, uploaded_records, strict=True):
        with tab:
            config = render_file_tab(record, is_base=record["index"] == 0)
            if config is not None:
                config_by_slot[int(record["index"])] = config

    all_loaded = (
        not missing_count
        and len(config_by_slot) == file_count
        and all(index in config_by_slot for index in range(file_count))
    )
    if not all_loaded:
        return

    configurations = [config_by_slot[index] for index in range(file_count)]
    st.markdown('<div class="step-label">STEP 3 · MERGE & EXPORT</div>', unsafe_allow_html=True)
    st.subheader("Set matching rules and create the workbook")

    with st.container(border=True):
        strategy_column, filename_column = st.columns([2, 1])
        with strategy_column:
            strategy_label = st.radio(
                "When a lookup key appears more than once",
                ["Use the first matching row", "Return every matching row"],
                horizontal=True,
                help="Returning every match can create multiple output rows for one base row.",
            )
        with filename_column:
            output_name = st.text_input("Output filename", value="merged_lookup.xlsx")

        option_columns = st.columns(3)
        trim_spaces = option_columns[0].toggle("Ignore surrounding spaces", value=True)
        ignore_case = option_columns[1].toggle("Ignore capitalization", value=True)
        numeric_equivalence = option_columns[2].toggle(
            "Match numbers to numeric text",
            value=True,
            help="For example, numeric 42 matches text '42'. Zero-padded IDs such as '0042' stay distinct.",
        )

    options = MatchOptions(
        trim_spaces=trim_spaces,
        case_sensitive=not ignore_case,
        numeric_equivalence=numeric_equivalence,
        duplicate_strategy="first" if strategy_label.startswith("Use") else "expand",
    )
    missing_returns = [
        uploads[index]["name"]
        for index in range(1, file_count)
        if not configurations[index]["return_columns"]
    ]
    if missing_returns:
        st.warning(
            "Choose at least one return column in: " + ", ".join(str(name) for name in missing_returns)
        )

    ready_records = [record for record in uploads if record is not None]
    fingerprint = generation_fingerprint(ready_records, configurations, options)
    create_clicked = st.button(
        "Create merged workbook",
        type="primary",
        use_container_width=True,
        disabled=bool(missing_returns),
    )
    if create_clicked:
        base = configurations[0]
        specs = [
            LookupSpec(
                file_name=str(ready_records[index]["name"]),
                sheet_name=configurations[index]["loaded"].sheet_name,
                dataframe=configurations[index]["loaded"].dataframe,
                key_column=configurations[index]["key_column"],
                return_columns=configurations[index]["return_columns"],
                prefix=configurations[index]["prefix"],
            )
            for index in range(1, file_count)
        ]
        try:
            with st.spinner("Matching rows and building your Excel workbook…"):
                result = merge_lookups(
                    base["loaded"].dataframe,
                    base["key_column"],
                    specs,
                    options,
                    base_return_columns=base["return_columns"],
                )
                workbook = build_output_workbook(
                    result,
                    [config["report"] for config in configurations],
                )
            st.session_state["generated_lookup"] = {
                "fingerprint": fingerprint,
                "result": result,
                "workbook": workbook,
            }
        except (ValueError, ExportLimitError) as exc:
            st.error(str(exc))

    generated = st.session_state.get("generated_lookup")
    if generated and generated.get("fingerprint") == fingerprint:
        render_result(generated, output_name)
    elif generated:
        st.info("The inputs changed. Create the workbook again to refresh the result.", icon="ℹ️")


if __name__ == "__main__":
    main()
