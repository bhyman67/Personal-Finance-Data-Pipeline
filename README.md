<h1>
    <p align="center">Personal Finance Data Pipeline</p>
</h1>

Data pipeline in Python that's designed to retrieve, process, and integrate my personal financial data for my custom Excel Money Manager tool. I just run it locally on my computer.

This pipeline automates the retrieval of transaction data from online banking platforms and investment information from Robinhood and Coinbase. It also downloads eStatements and merges them from banking portals. All data is then transformed and categorized into standardized formats for integration with my Excel workbook.

<div style="text-align: center;">
    <img src="Money%20Manager%20Pic.webp" alt="Money Manager Logo" width="500"/>
</div>

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
   - If ran from the terminal (rather than from Excel VBA), run with the current directory being a subfolder of the folder holding the .xlsm workbook. And be sure to manually pass along the creds while instantiating the object. 
     - `from retrieve_creds import retrieve_creds_for_money_manager`
     - `from personal_finance_data_pipeline import PersonalFinanceDataPipeline`
     - `creds = retrieve_creds_for_money_manager()`
     - `pipeline = PersonalFinanceDataPipeline(creds)`
   - The creds are actually optional and not needed for calling methods that don't require API access

The data pipeline performs various tasks such as:

- **Data Retrieval**: Fetching transaction data from multiple online banking portals and investment platforms
- **Data Transformation**: Normalizing and standardizing data from different sources into consistent formats
- **Transaction Categorization**: Automatically categorizing transactions and updating Excel spreadsheets
- **Investment Data Integration**: Retrieving and consolidating investment information from Robinhood and Coinbase
- **Document Management**: Downloading and merging eStatements from online banking portals
- **Excel Integration**: Writing processed data to structured Excel workbooks for analysis and reporting

### The investment portfolio part:

| Symbol | Name | Investment Type | Sector | Industry | Current Quantity | Current Equity | All Time Net Loss or Gain |
|--------|------|-----------------|--------|----------|------------------|----------------|---------------------------|
|        |      |                 |        |          |                  |                |                           |

## Configuration

Before running the data pipeline, make sure to set up the following configurations:

- **Excel Workbook**: The pipeline requires access to an Excel workbook containing necessary reference data and configurations. Update the file path in the script to point to the correct workbook.
- **Account Credentials**: Provide credentials for accessing online banking platforms, Robinhood, and Coinbase if required.
- **Browser Driver**: Ensure that the appropriate browser driver (e.g., ChromeDriver) is installed and its path is specified correctly in the pipeline configuration.

<p align="right">Click <a href="https://github.com/bhyman67/Functionalities-for-my-Money-Manager">here</a> to view the code in this project's repository<p>
