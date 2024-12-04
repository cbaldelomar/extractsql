"""
Extract data from database
"""

import pyodbc
from . import utils
from .tocsv import export_to_csv
from .toexcel import export_to_excel


def extract_to(
    connstring: utils.ConnString, query_file: str, output_file: str, **kwargs
):
    """
    Extract the query result to a file destination.

    Args:
        connstring: ConnString class.
        query_file: SQL query file to execute (can be a single-step or multi-step script).
        file_path: File destination.
        **kwargs: Additional arguments to pass to export function.
    """

    connection_string = utils.get_connection_string(connstring)
    query = utils.read_file(query_file)

    fn = export_to_excel if utils.is_extension(output_file, ".xlsx") else export_to_csv

    try:
        conn = pyodbc.connect(connection_string)

        cursor = conn.cursor()

        # Execute the query
        print("Executing query...")
        cursor.execute(query)

        # Loop through all result sets and fetch data if available
        while True:
            # Check if the result set has data
            if cursor.description:
                print("Exporting data...")

                print("-" * 50)

                fn(cursor, output_file, **kwargs)

                print("-" * 50)

                # Exit after fetching the final SELECT result
                break

            # Move to the next result set or exit if no more
            if not cursor.nextset():
                break

        # # Fetch all rows from the result
        # rows = cursor.fetchall()

        # # Extract column names
        # columns = [column[0] for column in cursor.description]
    except pyodbc.Error:
        print("Database error")
        raise
    except Exception:
        print("Error exporting to file")
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()
