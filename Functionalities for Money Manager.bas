Attribute VB_Name = "Module1"

Option Explicit

Dim otp As String
Dim tbl As ListObject
Dim newrow As ListRow
Dim first_empty_row As Integer
Dim python_source_code As String

' ****************************************************************************
' Form control buttons exist in the workbook for calling each of these functs.
' ****************************************************************************

Sub Retrieve_eStatements()

    ' Define Python source code
    python_source_code = "" & _
    "from Scripts.Money_Manager import Money_Manager; " & _
    "from Scripts.retrieve_creds import retrieve_creds_for_money_manager; " & _
    "creds = retrieve_creds_for_money_manager(); " & _
    "money_manager_obj = Money_Manager(creds); " & _
    "money_manager_obj.retrieve_estatements()"
    
    ' Run Python source code
    RunPython python_source_code

End Sub

Sub Add_Descriptions()
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts.Money_Manager import Money_Manager; " & _
    "money_manager_obj = Money_Manager(); " & _
    "money_manager_obj.add_transaction_descriptions()"
    
    ' Run Python source code
    RunPython python_source_code
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub

Sub Scrape_Posted_Txns()

    ' Clear all of the data in the "Raw Posted Txn Data - Scraped" worksheet
    Sheets("Posted and Archived Txns").Range("A1").CurrentRegion.Cells.Clear

    ' Prompt user for password to credential workbook
    otp = InputBox("Input", Title:="Robinhood OTP")
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts.Money_Manager import Money_Manager; " & _
    "from Scripts.retrieve_creds import retrieve_creds_for_money_manager; " & _
    "creds = retrieve_creds_for_money_manager(); " & _
    "money_manager_obj = Money_Manager(creds); " & _
    "otp = '" & otp & "'; " & _
    "money_manager_obj.set_cash_available_for_withdrawal(otp); " & _
    "money_manager_obj.scrape_txns(); "

    ' Run Python source code
    RunPython python_source_code
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub

Sub Refresh_Investment_Portfolio()

    ' Clear the old data
    Sheets("Personal Investment Portfolio").Range("holdings").Clear
    
    ' Prompt user for password to credential workbook
    otp = InputBox("Input", Title:="Robinhood OTP")
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts.Money_Manager import Money_Manager; " & _
    "from Scripts.retrieve_creds import retrieve_creds_for_money_manager; " & _
    "creds = retrieve_creds_for_money_manager(); " & _
    "money_manager_obj = Money_Manager(creds); " & _
    "otp = '" & otp & "'; " & _
    "money_manager_obj.get_investments(otp); "

    ' Run Python source code
    RunPython python_source_code
    
    ' Update this formula
    Sheets("Personal Investment Portfolio").Range("K12").Formula = "=SUM(holdings[equity],J6,K6)"
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub









