"""
Export data to Flat file
"""

# import os
import pyodbc
import pyarrow as pa
from pyarrow import csv
from tqdm import tqdm


def export_to_csv(cursor: pyodbc.Cursor, file_path: str, batch_size=100_000, **kwargs):
    """
    Export data from a pyodbc cursor to a CSV or delimited text file using `pyarrow`.

    Args:
        cursor (pyodbc.Cursor): The cursor object for database query execution.
        file_path (str): Path to the output CSV or text file.
        batch_size (int): Number of rows to fetch per batch.
        **kwargs: Additional args like column delimiter (delimiter) for the output file (e.g., ",", "\\t", "|").
    """
    # # Ensure output directory exists
    # os.makedirs(os.path.dirname(output_file), exist_ok=True)

    delimiter = kwargs.get("delimiter", ",")

    # Extract column names
    columns = [desc[0] for desc in cursor.description]

    total_rows = 0

    with open(file_path, "wb") as f:
        # Initialize the counter
        counter = tqdm(
            total=0,
            desc="Writing file",
            unit="rows",
            leave=True,
        )

        while True:
            # Fetch rows in batches
            rows = cursor.fetchmany(batch_size)

            if not rows:
                break

            batch_count = len(rows)
            total_rows += batch_count

            columns_data = list(zip(*rows))  # Transpose rows to columns

            # Convert rows to pyarrow.Table
            arrow_table = pa.table(
                [pa.array(column) for column in columns_data],
                names=columns,
            )

            write_options = csv.WriteOptions(
                delimiter=delimiter, include_header=(total_rows == batch_count)
            )

            # Write table to the file
            csv.write_csv(arrow_table, f, write_options=write_options)

            # Update the row counter
            counter.update(batch_count)

        counter.close()

        print(f"Processed {total_rows} rows")

    # print(f"Export completed: {total_rows} rows written to {file_path}")
