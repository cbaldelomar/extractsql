"""
Main
"""

import sys
import logging
import argparse
from . import utils
from .extract import extract_to
from .constants import FORMAT_XLSX, FORMAT_CSV, FORMAT_TXT, BATCH_SIZE, ROWS_PER_SHEET

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def _ensure_output_file(query_file, output_file, output_format):
    new_extension = f".{output_format}" if output_format else ""

    if not output_file:
        output_file = utils.replace_extension(query_file, new_extension)
    else:
        if utils.is_extension(output_file, ""):
            output_file = utils.replace_extension(output_file, new_extension)

    if utils.is_relative_path(output_file):
        output_file = utils.replace_path(output_file, query_file)

    output_file = utils.add_timestamp_to_filename(output_file)

    return output_file


def _get_args():
    parser = argparse.ArgumentParser(description="ExtractSQL Command-Line Tool")
    parser.add_argument("--version", action="version", version="1.0-0.9")

    parser.add_argument(
        "-s", "--server", required=True, help="Server name or IP address"
    )
    parser.add_argument("-d", "--database", required=True, help="Database name")
    parser.add_argument("-u", "--user", required=False, default=None, help="User name")
    parser.add_argument(
        "-p", "--password", required=False, default=None, help="User password"
    )
    parser.add_argument(
        "-q", "--query_file", required=True, help="Path to the query file"
    )
    parser.add_argument(
        "-o",
        "--output_file",
        required=False,
        help="Path to the output file (if not specified, query file name is used)",
    )

    format_values = [FORMAT_XLSX, FORMAT_CSV, FORMAT_TXT]
    parser.add_argument(
        "-f",
        "--output_format",
        choices=format_values,
        required=False,
        help="Format of the output file (required if output file path is not specified).",
    )

    parser.add_argument(
        "-c",
        "--column_delimiter",
        required=False,
        default=",",
        help='Column delimiter for the output file (csv, txt). Use between quotes (e.g., ",", "\\t", "|")',
    )

    parser.add_argument(
        "-b",
        "--batch_size",
        required=False,
        type=int,
        default=BATCH_SIZE,
        help="Number of rows per batch to read from SQL",
    )

    parser.add_argument(
        "-r",
        "--rows_per_sheet",
        required=False,
        type=int,
        default=ROWS_PER_SHEET,
        help="Rows per sheet (xlsx)",
    )

    # Parse the arguments
    return parser.parse_args()


def main():
    """
    Main entry point for the command-line interface.
    """

    args = _get_args()

    server = args.server
    database = args.database
    user = args.user
    password = args.password
    query_file = args.query_file
    output_file = args.output_file
    output_format = args.output_format
    column_delimiter = args.column_delimiter
    batch_size = args.batch_size
    rows_per_sheet = args.rows_per_sheet

    if not output_file and not output_format:
        print(
            "Output format (-f) is required if the output file (-o) was not specified."
        )
        sys.exit(1)

    query_file_extension = ".sql"

    if not utils.is_extension(query_file, query_file_extension):
        print(f"Invalid query file extension. Expected '{query_file_extension}'.")
        sys.exit(1)

    try:
        # Log start time
        start_time = utils.start_process()

        output_file = _ensure_output_file(query_file, output_file, output_format)

        odbc_driver = utils.get_connection_driver()

        if odbc_driver is None:
            raise ValueError("No suitable ODBC driver found.")
        else:
            print(f"Connecting using {odbc_driver} driver")

        connstring = utils.ConnString(server, database, user, password, odbc_driver)

        extract_to(
            connstring,
            query_file,
            output_file,
            delimiter=column_delimiter,
            batch_size=batch_size,
            rows_per_sheet=rows_per_sheet,
        )

        # Log end time
        utils.end_process(start_time)
    except Exception as e:
        print(e)
        sys.exit(1)

    print("\nData successfully exported to file:", output_file)


if __name__ == "__main__":
    main()
