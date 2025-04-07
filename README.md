# ExtractSQL

A command-line utility for exporting SQL query results to multiple file formats, such as Excel (`.xlsx`), CSV, and delimited text files.

Efficiently handles large datasets by processing data in batches with [pyodbc](https://github.com/mkleehammer/pyodbc), [XlsxWriter](https://github.com/jmcnamara/XlsxWriter), and [pyarrow](https://github.com/apache/arrow/tree/main/python).

## Features

- Export SQL query results to:
  - Excel (`.xlsx`) files
  - Flat files (`.csv`, `.txt`) with configurable delimiters
- Supports batch processing for large datasets
- Handles multi-step SQL scripts
- Automatic output file naming with timestamp support
- Flexible configuration via command-line arguments

## Installation

To install ExtractSQL, clone the repository and install the package:

```bash
git clone https://github.com/cbaldelomar/extractsql.git
cd extractsql
pip install .
```

## Prerequisites

Before using this tool, ensure the following are installed:

1. **Python 3.8 or higher**

    Install Python from [python.org](https://www.python.org/downloads/).
  
2. **ODBC Driver for SQL Server**

    This tool uses ODBC Driver for SQL Server to connect to Microsoft SQL Server. Download and install the driver from the [Microsoft Download Center](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server).

## Usage

### Command-Line Arguments

```bash
extractsql [options]
```

### Required Arguments

| Argument | Description |
| -------- | ----------- |
| `-s`, `--server` | Server name or IP address |
| `-d`, `--database` | Database name |
| `-q`, `--query_file` | Path to the SQL query file (`.sql` expected) |

> **Note:** If your path contains spaces, enclose it in double quotes (`"`).

### Optional Arguments

| Argument | Description | Default |
| -------- | ----------- | ------- |
| `-u`, `--user` | Database username (for authentication).	| `None` |
| `-p`, `--password` | Database password (for authentication). | `None` |
| `-o`, `--output_file` | Path to the output file. If not specified, a default file name based on the query file name will be used. If no directory is specified, the output file will be saved in the same directory as the query file. | [Derived automatically](#output-file-naming) |
| `-f`, `--output_format` |	Format of the output file (`xlsx`, `csv`, `txt`). Required if `-o` is not specified. | `None` |
| `-c`, `--column_delimiter` | Column delimiter for flat file formats (`csv`, `txt`). Example: `","`, `"\t"`, `"\|"` | `","` 
| `-b`, `--batch_size` | Number of rows to fetch from the database in each batch. | `100,000`
| `-r`, `--rows_per_sheet` | Maximum rows per Excel sheet.	| `1,000,000`

> **Note:** If `--user` and `--password` are not specified, **Windows Authentication** is used by default.

### Examples

#### Export to Excel

```bash
extractsql -s localhost -d my_database -q query.sql -f xlsx
```

#### Export to CSV with a custom delimiter

```bash
extractsql -s localhost -d my_database -q query.sql -o output.csv -c "|"
```

#### Export to a tab-delimited text file

```bash
extractsql -s localhost -d my_database -q query.sql -o output.txt -c "\t"
```

#### Using Authentication

```bash
extractsql -s localhost -d my_database -q query.sql -u my_user -p my_password -o output.xlsx
```

## Output File Naming

- If `-o` is not specified, the output file name is derived from the query file name.
- A timestamp in the format `YYYYMMDD_HH_MM_SS` is appended to the file name to ensure uniqueness.

For example:

* Query file: `example_query.sql`
* Output file: `example_query_20241203_15_30_45.xlsx`

## Future

- Add support for other RDBMS