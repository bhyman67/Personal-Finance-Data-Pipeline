# Money Manager Script

This Python script is designed to assist in managing financial transactions, investments, and eStatements by automating various tasks such as retrieving transaction data from online banking platforms, categorizing transactions, fetching investment information from Robinhood and Coinbase, and downloading eStatements from online banking portals.

## Requirements

- Python 3.x
- Required Python packages:
    - `dateutil`
    - `coinbase`
    - `robin_stocks`
    - `selenium`
    - `xlwings`
    - `pandas`
    - `PyPDF2`

## Usage

To use the script, follow these steps:

1. Ensure that all required Python packages are installed.
2. Set up the necessary configurations in the Excel workbook referenced by the script.
3. Run the script using Python.

The script performs various tasks such as:

- Retrieving transaction data from online banking portals.
- Categorizing transactions and updating Excel spreadsheets accordingly.
- Fetching investment information from Robinhood and Coinbase.
- Downloading eStatements from online banking portals.

## Configuration

Before running the script, make sure to set up the following configurations:

- Excel Workbook: The script requires access to an Excel workbook containing necessary reference data and configurations. Update the file path in the script to point to the correct workbook.
- Account Credentials: Provide credentials for accessing online banking platforms, Robinhood, and Coinbase if required.
- Browser Driver: Ensure that the appropriate browser driver (e.g., ChromeDriver) is installed and its path is specified correctly in the script.

