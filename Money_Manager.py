from coinbase.wallet.client import Client
from datetime import datetime, timedelta
from dateutil.relativedelta import *
import robin_stocks.robinhood as rh
from selenium import webdriver
import xlwings as xw
import pandas as pd
import traceback
import PyPDF2
import json
import time
import os

def assign_credit_debit_ind(amt):

    if amt >= 0:
        return("Credit")
    else:
        return("Debit")

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

class Money_Manager:

    def __init__(self, creds = None):

        # If this class is being instantiated in source ran by the RunPython VBA funct
        if __name__ == "Scripts_and_Trading_Bots.Money_Manager":

            self.wb = xw.Book.caller()

        else:

            self.wb = xw.Book("../Money Manager.xlsm") 

        # Read in reference data from the money manager Excel workbook
        self.desc_cat_lookup = self.wb.sheets["Txn Ref Data and Script Vars"].range("Table1").options(dict).value # dictionary
        self.desc_excludes = self.wb.sheets("Txn Ref Data and Script Vars").range("Table2").value # list
        self.manual_desc = self.wb.sheets("Txn Ref Data and Script Vars").range("Table3").options(pd.DataFrame, index = False, header = False).value # dataframe

        # Read in some script variables from the money manager Excel workbook
        self.account1_name = self.wb.sheets["Txn Ref Data and Script Vars"].range("M8").value
        self.account2_name = self.wb.sheets["Txn Ref Data and Script Vars"].range("M9").value
        self.account3_name = self.wb.sheets["Txn Ref Data and Script Vars"].range("M10").value
        self.credit_card_account_name = self.wb.sheets["Txn Ref Data and Script Vars"].range("M11").value

        if creds:
            self.firstbank_u = creds["FirstBank"][0]
            self.firstbank_p = creds["FirstBank"][1]
            self.robinhood_u = creds["Robinhood"][0]
            self.robinhood_p = creds["Robinhood"][1]
            self.coinbase_key_id = creds["Coinbase"][0]
            self.coinbase_key_secret = creds["Coinbase"][1]

    def __assign_exclude_ind(self, desc):

        # how to check if any items w/in a list are in a string
        truth_val = False
        if any(desc_exclude in desc for desc_exclude in self.desc_excludes):
            truth_val = True
        return(truth_val)

    def __categorize_description(self, desc):

        # loop through the dict
        for desc_substring in self.desc_cat_lookup:
            if desc_substring.upper() in desc.upper():
                # found a match
                return(self.desc_cat_lookup[desc_substring])

        # Return an empty string if no matches were found
        return("")

    def __del__(self):

        if __name__ != "Scripts_and_Trading_Bots.Money_Manager":

            self.wb.app.quit()

    def set_cash_available_for_withdrawal(self, otp):

        login = rh.authentication.login(self.robinhood_u, self.robinhood_p, mfa_code = otp)
        cash_available_for_withdrawal = rh.profiles.load_account_profile()["cash_available_for_withdrawal"]
        rh.authentication.logout()

        self.wb.sheets["Buying Power, Net Worth, Goals"].range("D11").value = cash_available_for_withdrawal

    def add_transaction_descriptions(self): # change this to categories, or... income/expense generator

        df = self.wb.sheets["Posted Transactions"].range("A1").current_region.options(pd.DataFrame).value
        df.reset_index(inplace = True)

        # Filter out all income expense excludes
        df = df[df["Income_Expense_Exclude"] == False]

        # Classify txns as either income or expense
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Income"

        # Flip the sign on all amounts to be positive (for credit card txns that show negetive amts)
        df["Amount"] = df["Amount"].apply(abs)

        # drop these cols "Income_Expense_Exclude","Credit_Debit_Ind"
        df.drop(["Income_Expense_Exclude","Credit_Debit_Ind"], axis=1, inplace=True)

        # Add description category col
        df["Description_Category"] = ""
        df["Description_Category"] = df["Description"].apply(self.__categorize_description)
        # Add these description categories manually
        for index, row in self.manual_desc.iterrows():
            df.loc[ (df["Date"]==row[0]) & (df["Amount"]==row[1]) & (df["Description"]==row[2]), "Description_Category"] = row[3]
        # Write the df to the Posted Transactions col
        self.wb.sheets["Income and Expenses"].range('A1').options(pd.DataFrame, index = False).value = df
        self.wb.sheets["Income and Expenses"].tables.add(source = self.wb.sheets["Income and Expenses"].range("A1").current_region, name = "transactions")
        self.wb.sheets["Income and Expenses"].range('A1').current_region.autofit()

    def get_investments(self, otp):

        # From both Robinhood and coinbase...

        # Login to robinhood
        login = rh.authentication.login(self.robinhood_u, self.robinhood_p, mfa_code = otp)

        # Get holdings data from Robinhood
        holdings_data = rh.account.build_holdings()
        df = pd.DataFrame(holdings_data)

        df = df.transpose()
        df.reset_index(inplace = True)
        df.rename(columns = {"index":"symbol"}, inplace = True)
        df = df[["symbol","name","equity","quantity","type"]]
        df.rename(
            columns={
                "symbol":"Symbol",
                "name":"Name",
                "equity":"Equity",
                "quantity":"Quantity",
                "type":"Type"
            }, 
            inplace=True
        )

        # Get needed data from coinbase
        client = Client(self.coinbase_key_id, self.coinbase_key_secret)
        crypto_accounts = client.get_accounts()["data"]
        # Build a list of tuples
        crypto_accounts_with_balances = []
        for crypto_account in crypto_accounts:

            if float(crypto_account["balance"]["amount"]) > 0:
            
                crypto_accounts_with_balances.append(
                    (
                        crypto_account["currency"],
                        crypto_account["name"],
                        crypto_account["native_balance"]["amount"],
                        crypto_account["balance"]["amount"],
                        "crypto"
                    )
                )

        df2 = pd.DataFrame(
            crypto_accounts_with_balances,
            columns = [
                "Symbol", "Name", "Equity", "Quantity", "Type"
            ]
        )

        # Pull out the USD amount
        usd_amt = df2[(df2["Symbol"]=="USD")].iloc[0]["Quantity"]
        df2.drop(index = df2[(df2["Symbol"]=="USD")].iloc[0].name, inplace = True)

        df = pd.concat([df,df2])

        # Write holdings data to the workbook and make it a table
        self.wb.sheets["Investment Portfolio"].range("F4").options(index=False).value = df
        self.wb.sheets["Investment Portfolio"].tables.add(source = self.wb.sheets["Investment Portfolio"].range("F4").current_region, name = "holdings")
        self.wb.sheets["Investment Portfolio"].range("F4").current_region.autofit()
        self.wb.sheets["Investment Portfolio"].range("M6").value = usd_amt

        # Log out of Robinhood
        rh.authentication.logout()

    def scrape_txns(self):

        # You need to put the error handling back into this scraping routine... 

        # Set dates (go up to 10 months back (9 at the least) depending on current date)
        #   -> goes back 9 months and then takes the 1st of that month
        crntDt = datetime.today()
        tenMnthPriorDt = datetime.today() - relativedelta(months = 9)
        tenMnthPriorDt = tenMnthPriorDt - timedelta(days = tenMnthPriorDt.day - 1)

        # List all of your accounts
        accounts = []
        accounts.append( '{{accountType={account_name}, selectedNumber=2d83bcf05b214c9b1b032bef309d72b4}}'.format(account_name = self.account1_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=6bdd4d69eb9a2b4b79df6d003b0c7244}}'.format(account_name = self.account2_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=9e720c749c446ee65976669a391134fb}}'.format(account_name = self.account3_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=8c4a6dff17073338f88e3f5b3ae117a2}}'.format(account_name = self.credit_card_account_name) )

        # Get your creds for online banking and instantiate the webdriver obj
        browser = webdriver.Chrome( executable_path = self.wb.sheets["Txn Ref Data and Script Vars"].range("M4").value )
        browser.implicitly_wait(30)

        # Login to OB (is there a way to use credentials that are saved in the browser???)
        browser.get('https://www.efirstbank.com/')
        browser.find_element_by_id('userId').send_keys(self.firstbank_u)
        browser.find_element_by_id('password').send_keys(self.firstbank_p)
        browser.find_element_by_id('logIn').click()

        # Grab account totals
        # Current balance from account 1 
        time.sleep(5)
        browser.find_element_by_xpath('//*[@id="js-acct-name"]/span[1]')
        account1_current_balance = browser.find_element_by_xpath('//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text
        # Click on account 3 and then grab the current balance from that
        browser.find_element_by_xpath('//*[@id="js-product-id-10620720"]/div[2]/div[1]/div/div[1]/p/span').click()
        time.sleep(2)
        account3_current_balance = browser.find_element_by_xpath('//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text
        # Click on account 2 and then grab the balance from that
        browser.find_element_by_xpath('//*[@id="js-product-id-5759370"]/div[2]/div[1]/div/div[1]/p/span').click()
        time.sleep(2)
        account2_current_balance = browser.find_element_by_xpath('//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text

        # Pull data for each account
        html_tables = []
        for account in accounts:

            # Pull up the "Download Account Info" page
            # browser.find_element_by_link_text('Online Banking').click()
            browser.find_element_by_xpath('//*[@id="obTab"]/a').click()
            time.sleep(1)
            browser.find_element_by_link_text('Downloads').click()

            # Select account
            browser.find_element_by_name('accountSelected').click()
            browser.find_element_by_xpath(f"//option[@value = '{account}']").click()

            # Set the date range (format is mm/dd/yyyy)
            browser.find_element_by_id('dateRangeRadio').click()
            account = account.split(',')[0].split('=')[1]
            browser.find_element_by_name('fromDate').send_keys(tenMnthPriorDt.strftime('%m/%d/%Y'))
            browser.find_element_by_name('toDate').send_keys(datetime.today().strftime('%m/%d/%Y'))

            # click  the view txns button
            browser.find_element_by_xpath("//input[@value='View Transactions']").click()

            # Find this element b4 the scrape to ensure that the page loads first
            browser.find_element_by_xpath("//table[@class='detail dataTable']")

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
        txns_df["Amount"] = txns_df["Amount"].replace('[\$,)]', '', regex=True).replace('[(]','-',regex=True).astype(float)

        # Add a credit/debit indicator column and an income/expence exclude indicator column
        desc_excludes = self.wb.sheets("Txn Ref Data and Script Vars").range("Table2").value
        # credit/debit indicator col
        txns_df["Credit_Debit_Ind"] = ""
        txns_df["Credit_Debit_Ind"] = txns_df["Amount"].apply(assign_credit_debit_ind)
        # indicator for transfers w/in internal accounts and credit card payments
        txns_df["Income_Expense_Exclude"] = ""
        txns_df["Income_Expense_Exclude"] = txns_df["Description"].apply(self.__assign_exclude_ind)

        # Log out and close both the browser and db cnxn
        time.sleep(2)
        browser.find_element_by_xpath("//span[@data-i18n = 'main:Log Out']").click()
        browser.quit()

        self.wb.sheets["Buying Power, Net Worth, Goals"].range("D8").value = float(account1_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Buying Power, Net Worth, Goals"].range("D9").value = float(account2_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Buying Power, Net Worth, Goals"].range("D10").value = float(account3_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Posted Transactions"].range('A1').options(pd.DataFrame, index = False).value = txns_df
        self.wb.sheets["Posted Transactions"].range('A1').current_region.autofit()

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
            downloaded_estatement_folder = self.wb.sheets["Txn Ref Data and Script Vars"].range("M5").value
            prefs = {
                'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                'savefile.default_directory': downloaded_estatement_folder
            }
            chromeOptions.add_experimental_option("prefs",prefs)
            chromeOptions.add_argument('--kiosk-printing')
            browser =  webdriver.Chrome(
                executable_path = self.wb.sheets["Txn Ref Data and Script Vars"].range("M4").value, 
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
            assets_and_liabilities = self.wb.sheets["Txn Ref Data and Script Vars"].range("M12").value
            firstbank_asset_accounts = os.path.join(assets_and_liabilities, "Assets", "Bank Accounts", "FirstBank")
            firstbank_liability_account = os.path.join(assets_and_liabilities, "Liabilities", "FirstBank {account_name}".format(account_name = self.credit_card_account_name))
            # -> full directory paths (and a list of all those paths)
            account1_stmt_path = os.path.join(firstbank_asset_accounts, self.account1_name, "Current Statements in OB")
            account2_stmt_path = os.path.join(firstbank_asset_accounts, self.account2_name, "Current Statements in OB")
            account3_stmt_path = os.path.join(firstbank_asset_accounts, self.account3_name, "Current Statements in OB")
            credit_card_stmt_path = os.path.join(firstbank_liability_account,"Current Statements in OB")
            current_statement_in_ob_path_list = [account1_stmt_path,account2_stmt_path,account3_stmt_path,credit_card_stmt_path]

            # Navigate to the eStatements in online banking
            # browser.find_element_by_link_text('Online Banking').click()
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
                    time_threshold = 5
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
            with open(self.wb.sheets["Txn Ref Data and Script Vars"].range("M7").value, 'w') as f:

                f.write("eStatements Retrieved Successfully")

        except Exception as e:

            # write to log file... 
            with open(self.wb.sheets["Txn Ref Data and Script Vars"].range("M7").value, 'w') as f:
                f.write(str(e))
                f.write(traceback.format_exc())
