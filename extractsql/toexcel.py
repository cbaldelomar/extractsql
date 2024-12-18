"""
Export data to Excel file
"""

from datetime import datetime, date, time
import pyodbc
import xlsxwriter
from tqdm import tqdm
from .constants import BATCH_SIZE, ROWS_PER_SHEET


def export_to_excel(
    cursor: pyodbc.Cursor, file_path: str, batch_size=BATCH_SIZE, **kwargs
):
    """
    Export data from a pyodbc cursor to a Excel file using `XlsxWriter`.

    Args:
        cursor (pyodbc.Cursor): The cursor object for database query execution.
        file_path (str): Path to the output Excel file.
        batch_size (int): Number of rows to fetch per batch.
        **kwargs: Additional args like number of rows per Excel sheet (rows_per_sheet).
    """

    rows_per_sheet = kwargs.get("rows_per_sheet", ROWS_PER_SHEET)

    # Extract column names
    columns = [column[0] for column in cursor.description]

    # Create a new workbook using xlsxwriter
    workbook = xlsxwriter.Workbook(file_path)

    # Define formats for dates, datetimes and time
    formats = {
        date: workbook.add_format({"num_format": "yyyy-mm-dd"}),
        datetime: workbook.add_format({"num_format": "yyyy-mm-dd HH:MM:SS"}),
        time: workbook.add_format({"num_format": "HH:MM:SS"}),
    }

    # Initialize control variables
    total_rows = 0
    row_count = 0
    sheet_index = 1
    worksheet = None
    counter = None

    while True:
        rows = cursor.fetchmany(batch_size)

        if not rows:
            # If the cursor does not return any rows at all, an empty sheet is created
            if worksheet is None:
                worksheet = _create_sheet(workbook, columns, sheet_index)

            break

        # Stream data row by row
        for row in rows:
            if row_count == rows_per_sheet:
                # Create a new sheet if the row limit is reached
                row_count = 0
                sheet_index += 1
                worksheet = None
                counter.close()

            if worksheet is None:
                worksheet = _create_sheet(workbook, columns, sheet_index)

                # Initialize the counter when the first row is processed
                counter = tqdm(
                    initial=row_count,
                    total=0,
                    desc=f"Writing {worksheet.name}",
                    unit="rows",
                    leave=True,
                )

            # Write the row data
            row_index = row_count + 1  # Account for header row
            for col_idx, value in enumerate(row):
                # Check for date/datetime/time values
                value_format = next(
                    (f for k, f in formats.items() if isinstance(value, k)), None
                )

                worksheet.write(row_index, col_idx, value, value_format)

                # if isinstance(value, (date, datetime)):
                #     value_format = (
                #         datetime_format if isinstance(value, datetime) else date_format
                #     )
                #     worksheet.write_datetime(row_index, col_idx, value, value_format)
                # else:
                #     worksheet.write(row_index, col_idx, value)

            row_count += 1
            total_rows += 1

            # Update the row counter
            counter.update(1)

    if not counter is None:
        counter.close()

    print(
        f"Exported {total_rows} rows in {sheet_index} sheet{"s" if sheet_index > 1 else ""}"
    )

    print("Saving workbook...")

    # Close the workbook to save the file
    workbook.close()


def _create_sheet(workbook: xlsxwriter.Workbook, columns: list[str], index: int):
    # Create a new worksheet
    worksheet = workbook.add_worksheet(f"Sheet{index}")

    # Write column headers
    for col_idx, header in enumerate(columns):
        worksheet.write(0, col_idx, header)

    return worksheet
