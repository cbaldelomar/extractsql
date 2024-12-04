"""
Export data to Excel file
"""

from datetime import datetime, date, time
import pyodbc
import xlsxwriter
from tqdm import tqdm

# from openpyxl import Workbook
# import pandas as pd
# from sqlalchemy import create_engine


# def execute_query(server: str, database: str, query: str) -> pd.DataFrame:
#     """
#     Return a DataFrame from the query result.
#     """

#     connection_string = (
#         f"mssql+pyodbc://{server}/{database}"
#         "?driver=ODBC+Driver+17+for+SQL+Server&Trusted_Connection=yes"
#     )

#     # Create SQLAlchemy engine
#     engine = create_engine(connection_string)

#     # Load query result into a Pandas DataFrame
#     with engine.connect() as connection:
#         df = pd.read_sql(query, connection)

#     return df


def export_to_excel(
    cursor: pyodbc.Cursor, file_path: str, batch_size=100_000, **kwargs
):
    """
    Export data from a pyodbc cursor to a Excel file using `XlsxWriter`.

    Args:
        cursor (pyodbc.Cursor): The cursor object for database query execution.
        file_path (str): Path to the output Excel file.
        batch_size (int): Number of rows to fetch per batch.
        **kwargs: Additional args like number of rows per Excel sheet (rows_per_sheet).
    """

    rows_per_sheet = kwargs.get("rows_per_sheet", 1_000_000)

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


# def _save_to_excel(cursor: pyodbc.Cursor, file_path: str, rows_per_sheet=1_000_000):
#     # Extract column names
#     columns = [column[0] for column in cursor.description]

#     # Create a new workbook
#     wb = Workbook()

#     # Remove default sheet
#     default_sheet = wb.active
#     wb.remove(default_sheet)

#     # Initialize control variables
#     row_count = 0
#     sheet_index = 1
#     worksheet = None

#     print("Exporting data...")
#     print("-" * 50)

#     # Stream data row by row
#     for row in cursor:
#         if row_count == rows_per_sheet:
#             row_count = 0
#             sheet_index += 1
#             worksheet = None

#         if not worksheet:
#             # Create sheet
#             sheet_name = f"Sheet{sheet_index}"

#             print("Processing:", sheet_name)

#             worksheet = wb.create_sheet(sheet_name)

#             # Write column headers
#             worksheet.append(columns)

#         # Write the row data
#         worksheet.append(list(row))

#         row_count += 1

#     print("-" * 50)

#     # Save the workbook
#     print("Saving workbook...")

#     wb.save(file_path)


# def execute_query(server: str, database: str, query: str) -> pd.DataFrame:
#     """
#     Return a DataFrame from the query result.

#     Args:
#         server: Name of the SQL Server.
#         database: Name of the database.
#         query: SQL query to execute (can be a single-step or multi-step script).

#     Returns:
#         DataFrame with result data.
#     """

#     connection_string = (
#         "DRIVER={ODBC Driver 17 for SQL Server};"
#         f"SERVER={server};"
#         f"DATABASE={database};"
#         "Trusted_Connection=yes"
#     )

#     try:
#         conn = pyodbc.connect(connection_string)

#         cursor = conn.cursor()

#         # Execute the query
#         cursor.execute(query)

#         # Loop through all result sets and fetch data if available
#         while True:
#             # Check if the result set has data
#             if cursor.description:
#                 # Extract column names
#                 columns = [column[0] for column in cursor.description]

#                 # Fetch all rows and column names
#                 rows = cursor.fetchall()

#                 # Exit after fetching the final SELECT result
#                 break

#             # Move to the next result set or exit if no more
#             if not cursor.nextset():
#                 break

#         # # Fetch all rows from the result
#         # rows = cursor.fetchall()

#         # # Extract column names
#         # columns = [column[0] for column in cursor.description]
#     except pyodbc.Error as ex:
#         print("Database error:", ex)
#         sys.exit()
#     finally:
#         if cursor:
#             cursor.close()
#             conn.close()

#     # Load data into a DataFrame
#     # coerce_float convert numeric/decimal sql types to float python type.
#     df = pd.DataFrame.from_records(rows, columns=columns, coerce_float=True)

#     return df


# def export_to_excel(df: pd.DataFrame, file_path: str, rows_per_sheet=1_000_000):
#     """
#     Export DataFrame to Excel sheets.

#     :param df: DataFrame to export.
#     :param file_path: Absolute path of the Excel file.
#     :param rows_per_sheet: (optional) Maximum rows per sheet (default: 1_000_000 for xlsx Excel limit).
#     """

#     # Export to Excel with multiple sheets
#     with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
#         # Loop through chunks of the DataFrame
#         for i in range(0, len(df), rows_per_sheet):
#             # Define the start and end of the chunk
#             chunk = df.iloc[i : i + rows_per_sheet]

#             # Write each chunk to a separate sheet
#             sheet_name = f"Sheet{i // rows_per_sheet + 1}"
#             print("Processing:", sheet_name)
#             chunk.to_excel(writer, sheet_name=sheet_name, index=False)
