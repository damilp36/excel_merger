from io import BytesIO

from openpyxl import Workbook

from excel_merger.excel_io import list_sheet_names, stream_sheet
from excel_merger.quality import analyze_quality


def workbook_bytes() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Customers"
    sheet.append(["Quarterly customer export"])
    sheet.append(["ID", "Name", "Name", None, "Empty"])
    sheet.append([1, "Ada", "Primary", "north", None])
    sheet.append([2, None, "Secondary", 99, None])
    sheet.append([3, "Lin", "Primary", "south", None])
    sheet.append([None, None, None, None, None])
    workbook.create_sheet("Archive")
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_lists_sheets_and_streams_selected_header() -> None:
    content = workbook_bytes()

    assert list_sheet_names(content, "customers.xlsx") == ["Customers", "Archive"]
    loaded = stream_sheet(content, "customers.xlsx", "Customers", header_row=2)

    assert list(loaded.dataframe.columns) == [
        "ID",
        "Name",
        "Name__2",
        "Unnamed_4",
        "Empty",
    ]
    assert loaded.dataframe.shape == (3, 5)
    assert loaded.blank_rows_removed == 1
    assert [change.reason for change in loaded.header_changes] == [
        "Duplicate header was made unique",
        "Blank header was given a name",
    ]


def test_quality_report_finds_missing_and_inconsistent_cells() -> None:
    loaded = stream_sheet(workbook_bytes(), "customers.xlsx", "Customers", header_row=2)
    report = analyze_quality(loaded)

    assert report.row_count == 3
    assert report.column_count == 5
    assert report.missing_cells == 4
    assert report.mixed_type_columns == 1
    assert report.inconsistent_cells == {(1, "Unnamed_4")}
    assert len(report.header_changes) == 2
