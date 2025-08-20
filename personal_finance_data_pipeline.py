
# Personal Finance Data Pipeline
# Automated retrieval and processing of financial data from multiple sources

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from dateutil.relativedelta import relativedelta
from coinbase.wallet.client import Client
from datetime import datetime, timedelta
from coinbase.rest import RESTClient
import robin_stocks.robinhood as rh
from selenium import webdriver
import xlwings as xw
import pandas as pd
import traceback
import PyPDF2
import json
import time
import os
import re

"""
Personal Finance Data Pipeline
This module provides a comprehensive data pipeline for automating personal finance management tasks. It retrieves, processes, and categorizes financial data from 
multiple sources (FirstBank, Robinhood, Coinbase), transforms transaction data, manages Excel workbooks via xlwings, and handles PDF eStatements.

Functions:
-----------
increment(match)
    Increments a matched integer by 1 and returns it as a string prefixed with '$'.
remove_visa(txn_description)
    Removes the substring 'VISA ' from a transaction description.
assign_credit_debit_ind(amt)
    Assigns 'Credit' if the amount is non-negative, otherwise 'Debit'.
extract_and_remove_date(description, post_date)
    Extracts a date in ' MM-DD' format from a transaction description and removes it; if not found, uses the provided post_date.
check_for_existing_pdf(file_dir)
    Checks if any PDF files exist in the specified directory.
PDFmerge(pdfs, output_pdf_name)
    Merges a list of PDF files into a single output PDF.

Classes:
--------
PersonalFinanceDataPipeline
    A comprehensive data pipeline class for managing personal finance data, including retrieving account balances and transactions, 
    processing income and expense data, retrieving investment holdings, and downloading/merging eStatements.
    
    Methods:
    --------
    __init__(self, creds=None)
        Initializes the pipeline instance, loads reference data and credentials.
    __assign_exclude_ind(self, desc)
        Determines if a transaction description should be excluded from income/expense calculations.
    __categorize_description(self, desc)
        Categorizes a transaction description based on reference data.
    __del__(self)
        Cleans up resources by quitting the Excel application if necessary.
    retrieve_account_data_and_transactions(self)
        Retrieves account balances and transaction data from FirstBank and Robinhood, processes and writes them to Excel.
            -> all time data is being pulled from RH
    refresh_income_and_expense_data(self)
        Processes transaction data to classify as income or expense, categorizes descriptions, and writes results to Excel.
    get_investments_v1(self)
        Retrieves and consolidates investment holdings from Robinhood and Coinbase, and writes them to Excel.
    retrieve_estatements(self)
        Automates downloading, saving, and merging of eStatements from FirstBank online banking, and logs the process.
"""

def increment(match):

    number = int(match.group(1))
    return f'${number + 1}'

def remove_visa(txn_description):
    
    return re.sub(r'\bVISA \b', '', txn_description)

def assign_credit_debit_ind(amt):

    if amt >= 0:
        return("Credit")
    else:
        return("Debit")

def extract_and_remove_date(description, post_date):
    """
    Extracts a date from the description and infers the correct year using post_date as reference.
    Maintains the original description cleaning behavior while adding year inference.
    
    Args:
        description (str): Transaction description that may contain a date
        post_date (str/datetime): The post date of the transaction
    
    Returns:
        tuple: (formatted_date, cleaned_description)
    """
    
    # Look for both patterns as in original version
    full_pattern = r'\bON\s\d{2}-\d{2}\s\d{4}\b'
    date_pattern = r'\s(\d{2})-(\d{2})'
    match = re.search(date_pattern, description)
    
    if match:
        # Extract month and day for date inference
        month, day = map(int, match.groups())
        # Create a date with the month and day, using post_date's year
        try:
            # First try with post_date's year
            txn_date = post_date.replace(month=month, day=day)
            # If the resulting date is more than 0 days after post_date, 
            # then it's likely from the previous year
            if (txn_date - post_date).days > 0:
                txn_date = txn_date.replace(year=txn_date.year - 1)
            
            # Use the original full pattern removal for description cleaning
            description_without_date = re.sub(full_pattern, '', description)
            return txn_date.strftime('%m/%d/%Y'), description_without_date.strip()
        except ValueError:
            # If the date is invalid (e.g., Feb 30), fall back to post_date
            pass
    
    # If no valid date found in description, use post_date
    return post_date.strftime('%m/%d/%Y'), description

def check_for_existing_pdf(file_dir):
  
    exists_a_pdf = False
  
    for itm in os.listdir(file_dir):
      if itm.endswith(".pdf"):
          exists_a_pdf = True
          break
    
    return exists_a_pdf

def PDFmerge(pdfs, output_pdf_name):
  
    # create the pdf file merger object
    pdfMerger = PyPDF2.PdfFileMerger()
 
    # append pdfs one by one
    for pdf in pdfs:
        pdfMerger.append(pdf)
 
    # write combined pdf to output_pdf_name pdf file
    with open(output_pdf_name, 'wb') as f:
        pdfMerger.write(f)

class PersonalFinanceDataPipeline:

    def __init__(self, creds = None):

        # If this class is being instantiated in the VBA source code (ran by the RunPython VBA funct)
        if __name__ == "Scripts.personal_finance_data_pipeline":

            self.wb = xw.Book.caller()

        else:

            self.wb = xw.Book("../Money Manager.xlsm")

        # Set credential variables if they were passed in
        if creds:

            self.firstbank_u = creds["FirstBank"][0]
            self.firstbank_p = creds["FirstBank"][1]
            self.robinhood_u = creds["Robinhood"][0]
            self.robinhood_p = creds["Robinhood"][1]
            self.coinbase_key_id = creds["Coinbase"][0]
            self.coinbase_key_secret = creds["Coinbase"][1]

        # Read in reference data from the Script Control Center & Ref Dta sheet in the Money Management Excel workbook
        self.description_category_lookup = self.wb.sheets["Script Control Center & Ref Dta"].range("Table1").options(dict).value # dict
        self.description_excludes = self.wb.sheets["Script Control Center & Ref Dta"].range("Table2").value # list
        self.manual_descriptions = self.wb.sheets["Script Control Center & Ref Dta"].range("Table3").options(pd.DataFrame, index = False, header = False).value # dataframe

        # Set account names, which come from the Script Control Center & Ref Dta sheet
        self.account1_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_1").value
        self.account2_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_2").value
        self.account3_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_3").value
        self.account4_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_4").value
        self.credit_card_account_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Credit_Card_Account").value

    def __assign_exclude_ind(self, desc):

        # how to check if any items w/in a list are in a string
        truth_val = False
        if any(desc_exclude in desc for desc_exclude in self.description_excludes):
            truth_val = True
        return(truth_val)

    def __categorize_description(self, desc):

        # loop through the dict
        for desc_substring in self.description_category_lookup:
            if desc_substring.upper() in desc.upper():
                # found a match
                return(self.description_category_lookup[desc_substring])

        # Return an empty string if no matches were found
        return("")

    def __get_upwork_income(self):
        """
        Retrieves Upwork income data from CSV file and formats it for integration
        with other transaction data.
        
        Returns:
            pd.DataFrame: Formatted Upwork income transactions
        """
        # Get the data from the Upwork Txns worksheet and filter for income transactions
        upwork_df = self.wb.sheets["Upwork Txns"].range("A1").expand().options(pd.DataFrame, header=True, index=False).value
        upwork_income_df = upwork_df[upwork_df["Transaction Type"].isin(["Bonus", "Fixed-price", "Hourly", "Expense reimbursement"])]

        #subset cols (Date, Amount $, Transaction Type, Transaction Summary)
        upwork_income_df = upwork_income_df[["Date", "Amount $", "Transaction Type", "Transaction Summary"]]

        # Rename columns 
        upwork_income_df = upwork_income_df.rename(columns={
            "Amount $": "Amount",
            "Transaction Summary": "Description",
            "Date": "Post Date",
            "Transaction Type": "Type"
        })

        # Convert dates to proper format
        upwork_income_df["Post Date"] = pd.to_datetime(upwork_income_df["Post Date"])
        # For Upwork, transaction date is always the same as post date
        upwork_income_df["Transaction Date"] = upwork_income_df["Post Date"].dt.strftime('%m/%d/%Y')
        upwork_income_df["Post Date"] = upwork_income_df["Post Date"].dt.strftime('%m/%d/%Y')

        # Add account and income indicator
        upwork_income_df["Account"] = "Upwork"
        upwork_income_df["Income or Expense"] = "Income"
        upwork_income_df["Description Category"] = "Upwork Income"

        # Ensure cols are in the correct order
        col_order = ["Post Date", "Transaction Date", "Account", "Amount", "Description", "Type", "Income or Expense", "Description Category"]
        upwork_income_df = upwork_income_df[col_order]

        return upwork_income_df

    def __del__(self):

        if __name__ != "Scripts.personal_finance_data_pipeline":

            self.wb.app.quit()

    def retrieve_account_data_and_transactions(self): 

        # **************************************************************************************************************************************************
        # Data retrieval from Robinhood (account balances, interest income, cash card transactions, and direct deposits) - Using Robin Stocks Unofficial API
        # **************************************************************************************************************************************************

        # Robinhood API calls for data
        rh.authentication.login(self.robinhood_u, self.robinhood_p) 

        # Balances
        rh_cash_available_for_withdrawal = rh.profiles.load_account_profile()["cash_available_for_withdrawal"]
        rhy_accounts_json_resp = rh.request_get("https://bonfire.robinhood.com/rhy/accounts/")

        # RH Spending and Direct Deposits
        card_settled_transactions_json_resp = rh.request_get("https://minerva.robinhood.com/cards/settled_transactions/")
        unified_transfers_json_resp = rh.request_get("https://bonfire.robinhood.com/paymenthub/unified_transfers/")

        # RH Interest Income (eventually need to add dividend income to this...)
        brokerage_interest_income_json_resp = rh.request_get("https://api.robinhood.com/accounts/sweeps")

        rh.authentication.logout()

        # Transform and normalize the brokerage interest income data
        brokerage_interest_income = pd.json_normalize(brokerage_interest_income_json_resp["results"])
        # Filter to only include data from November 1st, 2021 and on
        brokerage_interest_income = brokerage_interest_income[
            pd.to_datetime(brokerage_interest_income['pay_date']).dt.tz_localize(None) >= pd.Timestamp('2021-11-01')
        ]
        # Convert pay_date to datetime
        brokerage_interest_income['Date'] = pd.to_datetime(brokerage_interest_income['pay_date']).dt.strftime('%m/%d/%Y')
        brokerage_interest_income = pd.DataFrame({
            'Date': brokerage_interest_income['Date'],
            'Account': ' '.join(self.account3_name.split()[:2]),
            'Amount': brokerage_interest_income['amount.amount'],
            'Description': brokerage_interest_income['reason'],
            'Type': "INTEREST",
            'Credit_Debit_Ind': brokerage_interest_income['direction'].str.capitalize(),
            'Income_Expense_Exclude': False
        })
        
        # Write interest income to Robinhood Income tab
        self.wb.sheets["Robinhood Income"].range('A1').options(pd.DataFrame, index=False).value = brokerage_interest_income
        self.wb.sheets["Robinhood Income"].range('A1').current_region.autofit()

        # Transform and normalize the cash card settled transactions data
        card_settled_transactions = pd.json_normalize(card_settled_transactions_json_resp["results"])
        card_settled_transactions = pd.DataFrame({
            'Date': pd.to_datetime(card_settled_transactions['post_date']).dt.strftime('%m/%d/%Y'),
            'Account': 'Robinhood Cash Card',  # Static account name
            'Amount': card_settled_transactions['amount.amount'],  # Transaction amount
            'Description': card_settled_transactions['merchant_description'],  # Merchant description
            'Type': "CASH CARD",  # Transaction type from RH
            'Credit_Debit_Ind': card_settled_transactions['direction'].str.capitalize(),  # Credit/Debit indicator
            'Income_Expense_Exclude': False  # Set to False for all records
        })

        # Transform and normalize the payroll transfer data
        payroll_transfers = pd.json_normalize(unified_transfers_json_resp['results'])
        payroll_transfers = payroll_transfers[payroll_transfers['details.description'] == 'PAYROLL']
        payroll_transfers = pd.DataFrame({
            'Date': pd.to_datetime(payroll_transfers['details.settlement_date']).dt.strftime('%m/%d/%Y'),
            'Account': 'Robinhood Cash Management',
            'Amount': payroll_transfers['amount'],
            'Description': payroll_transfers['details.description'].fillna('') + ' ' + payroll_transfers['details.originator_name'].fillna(''),
            'Type': payroll_transfers['transfer_type'],
            'Credit_Debit_Ind': payroll_transfers['details.direction'].str.capitalize(),
            'Income_Expense_Exclude': False
        })

        # Combine card transactions and payroll transfers and write to Robinhood Spending tab
        rh_spending_df = pd.concat([card_settled_transactions, payroll_transfers])
        rh_spending_df = rh_spending_df.sort_values(by='Date', ascending=False)
        self.wb.sheets["Robinhood Spending"].range('A1').options(pd.DataFrame, index=False).value = rh_spending_df
        self.wb.sheets["Robinhood Spending"].range('A1').current_region.autofit()

        # Get the Robinhood spending account available cash balance
        rhy_accounts = pd.json_normalize(rhy_accounts_json_resp["results"])
        spending_account_available_cash = rhy_accounts[rhy_accounts['purpose'] == 'spend'].iloc[0]['cash_available']

        # ***************************************************************************************************************
        #        Data retrieval from FirstBank (account balances and transactions) - Web scraping using Selenium
        # ***************************************************************************************************************

        # You need to put the error handling back into this scraping routine...

        # Get the current date and the first of the previous month
        crntDt = datetime.today().strftime('%m/%d/%Y')
        firstOfPrevMnth = (datetime.today() - relativedelta(months=1)).replace(day=1).strftime('%m/%d/%Y')

        # List all of your FirstBank accounts
        accounts = []
        accounts.append( '{{accountType={account_name}, selectedNumber=2d83bcf05b214c9b1b032bef309d72b4}}'.format(account_name = self.account1_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=9e720c749c446ee65976669a391134fb}}'.format(account_name = self.account2_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=8c4a6dff17073338f88e3f5b3ae117a2}}'.format(account_name = self.credit_card_account_name) )

        # Instantiate the webdriver object
        service = webdriver.chrome.service.Service(self.wb.sheets["Script Control Center & Ref Dta"].range("Chromedriver").value)
        browser = webdriver.Chrome(service=service)
        browser.implicitly_wait(30)

        # Login to OB (is there a way to use credentials that are saved in the browser???)
        browser.get('https://www.efirstbank.com/')
        browser.find_element(By.ID, 'userId').send_keys(self.firstbank_u)
        browser.find_element(By.ID, 'password').send_keys(self.firstbank_p)
        browser.find_element(By.ID, 'logIn').click()

        # Grab account totals 
        # Current balance from account 1 
        time.sleep(10)
        browser.find_element(By.XPATH, '//*[@id="js-acct-name"]/span[1]')
        account1_current_balance = browser.find_element(By.XPATH, '//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text
        # Click on account 2 and then grab the current balance from that
        browser.find_element(By.XPATH, '//*[@id="js-product-id-10620720"]/div[2]/div[1]/div/div[1]/p/span').click()
        time.sleep(3)
        account2_current_balance = browser.find_element(By.XPATH, '//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text

        # Pull data for each account
        html_tables = []
        for account in enumerate(accounts):

            # Pull up the "Download Account Info" page

            # Click on the account tab
            browser.find_element(By.XPATH, '//*[@id="obTab"]/a').click()
            time.sleep(1)
            browser.find_element(By.LINK_TEXT, 'Downloads').click()

            # Select account
            browser.find_element(By.NAME, 'accountSelected').click()
            browser.find_element(By.XPATH, f"//option[@value = '{account[1]}']").click()

            # Set the date range (format is mm/dd/yyyy)
            browser.find_element(By.ID, 'dateRangeRadio').click()
            if account[0] == 0 or account[0] == 1:
                account = account[1].split(',')[0].split('=')[1].split(" ")[1]
            else:
                account = account[1].split(',')[0].split('=')[1]
            browser.find_element(By.NAME, 'fromDate').send_keys(firstOfPrevMnth)
            browser.find_element(By.NAME, 'toDate').send_keys(crntDt)

            # click  the view txns button
            browser.find_element(By.XPATH, "//input[@value='View Transactions']").click()

            # Find one of these two elements B4 the scrape to ensure that the page loads first
            elmnt_txt = browser.find_element(By.XPATH, "//table[@class='detail dataTable'] | //*[@id='contentContainer']/div[2]/div/p").text

            if "No transactions were found in the specified range." in elmnt_txt:

                pass 

            else:

                # Data Scrape
                html_table = pd.read_html(browser.page_source)[0]

                # Add an account col and then append data table to list
                html_table["Account"] = account
                html_table = html_table[["Date","Account","Amount","Description","Type"]]
                html_tables.append(html_table)

        # Combine all of the DFs and then export
        txns_df = pd.concat(html_tables)
        characters_to_replace = {
            "$":"",
            ",":"",
            "(":"-",
            ")":""
        }
        txns_df["Amount"] = txns_df["Amount"].replace(r'[\$,)]', '', regex=True).replace(r'[(]', '-', regex=True).astype(float) 

        # Log out and close both the browser and db cnxn
        time.sleep(2)
        browser.find_element(By.XPATH, "//span[@data-i18n = 'main:Log Out']").click()
        browser.quit()

        # ****************************************************************************************************************
        # Light enrichment of the data but most processing work will be done in the refresh_income_and_expense_data method
        # ****************************************************************************************************************

        # Classify transactions as either credit or debit
        txns_df["Credit_Debit_Ind"] = ""
        txns_df["Credit_Debit_Ind"] = txns_df["Amount"].apply(assign_credit_debit_ind)
        # Indicate whether the transaction is an income or expense
        txns_df["Income_Expense_Exclude"] = ""
        txns_df["Income_Expense_Exclude"] = txns_df["Description"].apply(self.__assign_exclude_ind)

        # ***********************************************************************************************************************
        # Write account balances and FirstBank transactions to Excel
        # ***********************************************************************************************************************

        # Write data to Excel
        # -> account balances to the Overview sheet
        self.wb.sheets["Overview"].range( self.account1_name.replace(" ","_") ).value = float(account1_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Overview"].range( self.account2_name.replace(" ","_") ).value = float(account2_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Personal Investment Portfolio"].range( self.account3_name.replace(" ","_") ).value = rh_cash_available_for_withdrawal 
        self.wb.sheets["Overview"].range(self.account4_name.replace(" ","_")).value = float(spending_account_available_cash)
        # -> transactions to the All Bank Transactions sheet
        self.wb.sheets["All Bank Transactions"].range('A1').options(pd.DataFrame, index = False).value = txns_df
        self.wb.sheets["All Bank Transactions"].range('A1').current_region.autofit()

    def refresh_income_and_expense_data(self): # change this to categories, or... income/expense generator

        # Get FirstBank transactions
        df = self.wb.sheets["All Bank Transactions"].range("A1").current_region.options(pd.DataFrame).value
        df.reset_index(inplace = True)

        # Get Robinhood transactions and combine all data sets
        rh_spending_df = self.wb.sheets["Robinhood Spending"].range('A1').current_region.options(pd.DataFrame, header=True, index=False).value
        rh_income_df = self.wb.sheets["Robinhood Income"].range('A1').current_region.options(pd.DataFrame, header=True, index=False).value
        df = pd.concat([df, rh_spending_df, rh_income_df])

        # Filter out all income expense excludes
        df = df[df["Income_Expense_Exclude"] == False]

        # Classify txns as either income or expense
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] == "Robinhood Brokerage") & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] == "Robinhood Cash Card") & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Expense"

        # Flip the sign on all amounts to be positive (for credit card txns that show negetive amts)
        # Convert the column to a numeric type and format as Accounting (in Excel), too. - maybe do this later...
        df["Amount"] = df["Amount"].apply(abs)

        # drop these cols "Income_Expense_Exclude","Credit_Debit_Ind"
        df.drop(["Income_Expense_Exclude","Credit_Debit_Ind"], axis=1, inplace=True)

        # Rename the income/expense indicator col
        df.rename(columns={"Income_Expense_Ind":"Income or Expense"}, inplace=True)

        # Clean up the description col and extract transaction dates
        df["Description"] = df["Description"].apply(lambda x: remove_visa(x))
        # Convert Post Date to datetime if it isn't already
        df["Date"] = pd.to_datetime(df["Date"])
        # Extract transaction dates and clean descriptions
        df[["Txn Date", "Description"]] = df.apply(
            lambda row: extract_and_remove_date(row["Description"], row["Date"]),
            axis=1, result_type="expand"
        )

        # Add description category col
        df["Description_Category"] = ""
        df["Description_Category"] = df["Description"].apply(self.__categorize_description)
        # Add these description categories manually
        for index, row in self.manual_descriptions.iterrows():
            if ((df["Date"]==row[0]) & (df["Amount"]==row[1]) & (df["Description"]==row[2])).any():
                df.loc[ (df["Date"]==row[0]) & (df["Amount"]==row[1]) & (df["Description"]==row[2]), "Description_Category"] = row[3]

        # Rename and reorder columns
        df.rename(columns={
            "Date": "Post Date",
            "Description_Category": "Description Category",
            "Txn Date": "Transaction Date"
        }, inplace=True)
        
        # Convert dates to final format
        df["Post Date"] = df["Post Date"].dt.strftime('%m/%d/%Y')
        
        # Reorder columns
        df = df[[
            "Post Date",
            "Transaction Date",
            "Account",
            "Amount",
            "Description",
            "Type",
            "Income or Expense",
            "Description Category"
        ]]
        
        # **********************************************************************************************************
        # Replace txns for the HOA roof replacement
        incoming_txn = df[
            (df["Amount"] == 16815.39) & (df["Income or Expense"] == "Income") & (df["Post Date"] == "02/24/2025")
        ]
        outgoing_txn = df[
            (df["Amount"] == 17316.39) & (df["Income or Expense"] == "Expense") & (df["Post Date"] == "02/25/2025")
        ] 
        df.drop(incoming_txn.index, inplace = True)
        df.drop(outgoing_txn.index, inplace = True)
        new_txn = pd.DataFrame([{
            "Post Date": outgoing_txn["Post Date"].values[0],
            "Transaction Date": outgoing_txn["Post Date"].values[0],  # Using post date as transaction date since this is a manual entry
            "Account": outgoing_txn["Account"].values[0],
            "Amount": 500,
            "Description": "Safeco Insurance Deductible - HOA Roof Replacement",
            "Type": outgoing_txn["Type"].values[0],
            "Income or Expense": "Expense",
            "Description Category": ""
        }])
        # **********************************************************************************************************

        # Add upwork income 
        upwork_income = self.__get_upwork_income()

        # Combine all data sources
        df = pd.concat([df, new_txn, upwork_income])

        # Convert Post Date back to datetime for proper sorting
        df["Post Date"] = pd.to_datetime(df["Post Date"])
        
        # Sort by Post Date while it's still in datetime format
        df = df.sort_values(by="Post Date", ascending=False)
        
        # Convert back to string format for Excel
        df["Post Date"] = df["Post Date"].dt.strftime('%m/%d/%Y')
        
        # Write the df to the Income and Expenses tab and make it a data table
        self.wb.sheets["Income and Expense Tracking"].tables("transactions").range.clear()
        self.wb.sheets["Income and Expense Tracking"].range('A1').options(pd.DataFrame, index = False).value = df
        self.wb.sheets["Income and Expense Tracking"].tables.add(source = self.wb.sheets["Income and Expense Tracking"].range("A1").current_region, name = "transactions")
        self.wb.sheets["Income and Expense Tracking"].range('A1').current_region.autofit()

    def get_investments_v1(self): 

        # provide option to pull all time investment data from Robinhood and Coinbase (from file...)

        # +++ Robinhood +++

        # Login
        rh.authentication.login(self.robinhood_u, self.robinhood_p)

        # Get holdings data
        holdings_data = rh.account.build_holdings()
        df = pd.DataFrame(holdings_data)

        # Parse it out 
        df = df.transpose()
        df.reset_index(inplace = True)
        df.rename(columns = {"index":"symbol"}, inplace = True)
        df = df[["symbol","name","type","quantity","equity"]]
        df.rename(
            columns={
                "symbol":"Symbol",
                "name":"Name",
                "type":"Type",
                "quantity":"Quantity",
                "equity":"Current Equity"
            }, 
            inplace=True
        )

        # Log out
        rh.authentication.logout()

        # +++ Coinbase +++

        # Get all your crypto accounts
        client0 = Client("0", "0")
        client = RESTClient(api_key=self.coinbase_key_id, api_secret=self.coinbase_key_secret)
        crypto_accounts = client.get_accounts()["accounts"]
        # Build a list of tuples
        crypto_accounts_with_balances = []
        for crypto_account in crypto_accounts:

            if float(crypto_account["available_balance"]["value"]) > 0:
            
                crypto_symbol = crypto_account["available_balance"]["currency"]
                crypto_name = crypto_account["name"]
                crypto_quantity = crypto_account["available_balance"]["value"]
                crypto_exchange_rate = client0.get_exchange_rates(currency=crypto_symbol)["rates"]["USD"] # need to fix this
                crypto_equity = str(float(crypto_exchange_rate) * float(crypto_quantity))

                crypto_accounts_with_balances.append(
                    (
                        crypto_symbol,
                        crypto_name,
                        "cryptocurrency",
                        crypto_quantity,
                        crypto_equity  
                    )
                )

        df2 = pd.DataFrame(
            crypto_accounts_with_balances,
            columns = [
                "Symbol", "Name", "Type", "Quantity", "Current Equity"
            ]
        )

        # Pull out the USD amount
        usd_amt = df2[(df2["Symbol"]=="USD")].iloc[0]["Quantity"]
        df2.drop(index = df2[(df2["Symbol"]=="USD")].iloc[0].name, inplace = True)

        df = pd.concat([df,df2])

        # +++ Write it all to Excel +++

        # Write holdings data to the workbook and make it a table
        holdings_table_address = self.wb.sheets["Personal Investment Portfolio"].tables["holdings"].range.address
        self.wb.sheets["Personal Investment Portfolio"].range(re.sub(r'\$(\d+)$', increment, holdings_table_address)).delete(shift='up') 
        self.wb.sheets["Personal Investment Portfolio"].range("A1").options(index=False).value = df
        self.wb.sheets["Personal Investment Portfolio"].tables.add(source = self.wb.sheets["Personal Investment Portfolio"].range("A1").current_region, name = "holdings")
        self.wb.sheets["Personal Investment Portfolio"].range("A1").current_region.autofit()
        self.wb.sheets["Personal Investment Portfolio"].range("coinbase_usd_cash_bal").value = usd_amt

    def retrieve_estatements(self):

        try:
                
            # Instantiate the webdriver object 
            chromeOptions = webdriver.ChromeOptions()
            settings = {
                "recentDestinations": [
                    {
                        "id": "Save as PDF",
                        "origin": "local",
                        "account": ""
                    }
                ],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }
            downloaded_estatement_folder = self.wb.sheets["Script Control Center & Ref Dta"].range("Downloaded_eStatement_folder").value
            prefs = {
                'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                'savefile.default_directory': downloaded_estatement_folder
            }
            chromeOptions.add_experimental_option("prefs",prefs)
            chromeOptions.add_argument('--kiosk-printing')
            browser =  webdriver.Chrome(
                executable_path = self.wb.sheets["Script Control Center & Ref Dta"].range("Chromedriver").value, 
                options = chromeOptions
            )
            browser.implicitly_wait(10)

            # Login to OB (is there a way to use credentials that are saved in the browser???)
            browser.get('https://www.efirstbank.com/')
            browser.find_element_by_id('userId').send_keys(self.firstbank_u)
            browser.find_element_by_id('password').send_keys(self.firstbank_p)
            browser.find_element_by_id('logIn').click()

            # Define folder locations
            # -> root paths
            assets_and_liabilities = self.wb.sheets["Script Control Center & Ref Dta"].range("Assets_and_Liabilities_Path").value
            firstbank_asset_accounts = os.path.join(assets_and_liabilities, "Assets", "Bank Accounts", "FirstBank")
            firstbank_liability_account = os.path.join(assets_and_liabilities, "Liabilities", "FirstBank {account_name}".format(account_name = self.credit_card_account_name))
            # -> full directory paths (and a list of all those paths)
            account1_stmt_path = os.path.join(firstbank_asset_accounts, self.account1_name, "Current Statements in OB")
            account2_stmt_path = os.path.join(firstbank_asset_accounts, self.account2_name, "Current Statements in OB")
            account3_stmt_path = os.path.join(firstbank_asset_accounts, self.account3_name, "Current Statements in OB")
            credit_card_stmt_path = os.path.join(firstbank_liability_account,"Current Statements in OB")
            current_statement_in_ob_path_list = [account1_stmt_path,account2_stmt_path,account3_stmt_path,credit_card_stmt_path]

            # Navigate to the eStatements in online banking
            browser.find_element_by_xpath('//*[@id="obTab"]/a').click()
            browser.find_element_by_link_text('eStatements').click()

            xpath = '//*[@id="contentContainer"]/div[2]/div[2]/table/tbody/tr[{tr_index}]/td[{td_index}]' # /select
            for i in range(4):
                
                # Reference two siblings up from the parent to to the account name
                current_account = browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 1)).text
                print(current_account)
                
                current_account_dropdowns = browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 3)+"/select")
                date_options = current_account_dropdowns.find_elements_by_tag_name("option")
                for date_option in date_options:

                    # Select the statement date that you want to pull a statement for
                    statement_date = date_option.get_attribute("value")
                    statement_date = statement_date.replace("/","-")
                    date_option.click()
                    current_tab = browser.current_window_handle

                    # Click on the eStatement button for the estatement to show up in a new tab w/in the browser
                    browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 4) + '/div/input').click()

                    # switch into the new tab and wait for it to load
                    browser.switch_to.window(browser.window_handles[1])
                    embeded_web_element = browser.find_element_by_tag_name("embed")

                    # Print the page to pdf
                    browser.execute_script("window.print();")

                    # Folder reference will depend on...
                    if current_account == self.account1_name:
                        export_folder = account1_stmt_path
                    elif current_account == self.account2_name:
                        export_folder = account2_stmt_path
                    elif current_account == self.account3_name:
                        export_folder = account3_stmt_path
                    elif current_account == self.credit_card_account_name:
                        export_folder = credit_card_stmt_path

                    # Wait for estatementprep.do.pdf to be downloaded
                    time_threshold = 8
                    j = 1
                    # while not os.path.exists(os.path.join(downloaded_estatement_folder,"estatementprep.do.pdf")):
                    while not check_for_existing_pdf(downloaded_estatement_folder) and j < time_threshold:
                        time.sleep(2)
                        j += 1 
                    # Grab the name of the one file that should be in there
                    f_name = os.listdir(downloaded_estatement_folder)[0]
                    os.rename(
                        os.path.join(downloaded_estatement_folder,f_name),
                        os.path.join(export_folder,statement_date + ".pdf")
                    )
                    # wait for folder to be empty?
                    while check_for_existing_pdf(downloaded_estatement_folder) and j < time_threshold:
                        time.sleep(1)
                        j += 1 

                    browser.close()
                    browser.switch_to.window(current_tab)

                    #break # This is temporary

            # Log out and close both the browser and db cnxn
            browser.find_element_by_xpath("//span[@data-i18n = 'main:Log Out']").click()
            browser.quit()

            # +++ PDF merge routine +++
            
            # Loop through all folders holding bank statements
            for current_statement_in_ob_path in current_statement_in_ob_path_list:
                
                pdf_list = []
                for eStatement in os.listdir( current_statement_in_ob_path ):

                    pdf_list.append( os.path.join(current_statement_in_ob_path,eStatement) )

                # Merge all PDFs in the PDF list together
                eStatement_account = os.path.basename( os.path.split(current_statement_in_ob_path)[0] )
                PDFmerge(
                    pdf_list,
                    os.path.join(
                        os.path.abspath(os.path.join(current_statement_in_ob_path, os.pardir)),
                        'Merged {eStatement_account} eStatements.pdf'.format(eStatement_account = eStatement_account)
                    )
                )

            # write to log file
            with open(self.wb.sheets["Script Control Center & Ref Dta"].range("Log_File").value, 'w') as f:

                f.write("eStatements Retrieved Successfully")

        except Exception as e:

            # write to log file... 
            with open(self.wb.sheets["Script Control Center & Ref Dta"].range("Log_File").value, 'w') as f:
                f.write(str(e))
                f.write(traceback.format_exc())
