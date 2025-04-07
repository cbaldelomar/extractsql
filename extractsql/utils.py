"""
Utils module
"""

import time
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
import pyodbc
from charset_normalizer import from_bytes
from .constants import READ_BYTES, DEFAULT_ENCODING


@dataclass
class ConnString:
    """
    Define properties of a connection string.
    `Trusted_Connection=yes` if `user` and `password` are not defined.
    """

    server: str
    database: str
    user: str = None
    password: str = None
    driver: str = None
    trusted: bool = field(init=False)

    def __post_init__(self):
        self.trusted = self.user is None and self.password is None


def get_connection_string(connstring: ConnString) -> str:
    """
    Return a pyodbc connection string
    """

    credentials = (
        "Trusted_Connection=yes"
        if connstring.trusted
        else f"UID={connstring.user};PWD={connstring.password}"
    )

    connection_string = (
        f"DRIVER={connstring.driver};"
        f"SERVER={connstring.server};"
        f"DATABASE={connstring.database};"
        f"{credentials}"
    )

    return connection_string


def get_connection_driver():
    """
    Get the best available ODBC driver for SQL Server from the installed drivers.

    Returns:
        str: The name of the best available driver, or None if no suitable driver is found.
    """
    preferred_drivers = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "ODBC Driver 11 for SQL Server",
        "SQL Server",
    ]

    installed_drivers = pyodbc.drivers()

    for driver in preferred_drivers:
        if driver in installed_drivers:
            return "{" + driver + "}"

    return None


def read_file(file_path: str) -> str:
    """
    Returns a string with the content of a text file.

    :param file_path: Absolute path of the file.

    :return: String with the file content.
    :rtype: str
    """

    try:
        with open(file_path, "rb") as f:
            raw = f.read(READ_BYTES)

        result = from_bytes(raw).best()
        encoding = result.encoding if result else DEFAULT_ENCODING

        with open(file_path, "r", encoding=encoding, errors="replace") as f:
            return f.read()
    except Exception:
        print("Error reading the file:", file_path)
        raise


def is_extension(file_name: str, extension: str) -> bool:
    """
    Return `True` if file have the extension.

    :param file_name: The name/path of the file.
    :param extension: The extension (including period).

    :return: `True` if file have the extension. Otherwise `False`.
    :rtype: bool
    """
    return Path(file_name).suffix == extension


def replace_extension(file_name: str, new_extension: str) -> str:
    """
    Replace the extension of the file.

    :param file_name: Original file name (could be an absolute file path).
    :param new_extension: New file extension.

    :return: File name with new extension.
    :rtype: str
    """

    return str(Path(file_name).with_suffix(new_extension))


def add_timestamp_to_filename(file_name: str, date_format="%Y%m%d_%H_%M_%S") -> str:
    """
    Add the current date and time to a file name.

    :param file_name: The original file name (could be an absolute file path).
    :param date_format: The format for the timestamp (default: YYYYMMDD_HH_MM_SS).

    :return: New file name with the timestamp added.
    :rtype: str
    """

    # Get datetime in specified format
    timestamp = datetime.now().strftime(date_format)

    # Extract file name
    original_file = Path(file_name)

    # Add timestamp to the file name
    new_file_name = f"{original_file.stem}_{timestamp}"

    return str(original_file.with_stem(new_file_name))


def is_relative_path(path: str) -> bool:
    """
    Returns `True` if the path is relative.

    :param path: The actual path.

    :return: `True` if the path is relative. Otherwise `False`.
    :rtype: bool
    """

    return not Path(path).is_absolute()


def replace_path(file_path: str, new_path: str) -> str:
    """
    Replaces the path of a file for a new path.

    :param file_path: The file path.
    :param new_path: The new path (must be an absolute path and could be a path of other file).

    :return: The new file path.
    :rtype: str
    """

    path_old = Path(file_path)
    path_new = Path(new_path)

    return str(Path.joinpath(path_new.parent, path_old.name))


def start_process() -> float:
    """
    Log the start datetime and return the time.
    """

    # Log start time
    start_time = time.time()

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Process started     {'-' * 10} {current_time}")

    return start_time


def end_process(start_time: float):
    """
    Log the end datetime and return the elapsed between start time and end time in minutes.
    """

    # Log end time
    end_time = time.time()

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Process completed   {'-' * 10} {current_time}")

    # Calculate elapsed time
    elapsed_time_seconds = end_time - start_time
    hours, remainder = divmod(elapsed_time_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    # Format elapsed time dynamically
    if hours > 0:
        print(f"Execution Time: {int(hours)}h{int(minutes)}m{seconds:.2f}s")
    elif minutes > 0:
        print(f"Execution Time: {int(minutes)}m{seconds:.2f}s")
    else:
        print(f"Execution Time: {seconds:.2f}s")


def ensure_valid_escape_sequences(text: str) -> str:
    """
    Ensure that escape sequences are valid in the text.
    If the string contains invalid escape sequences, it is returned unchanged.

    :param text: The text to validate.

    :return: The text with valid escape sequences.
    :rtype: str
    """

    try:
        # Attempt to decode escape sequences
        return text.encode(DEFAULT_ENCODING).decode("unicode_escape")
    except UnicodeDecodeError:
        # Return the original string if decoding fails
        return text
