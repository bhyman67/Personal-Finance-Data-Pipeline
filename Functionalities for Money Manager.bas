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
    "from Scripts_and_Trading_Bots.Money_Manager import Money_Manager; " & _
    "from Scripts_and_Trading_Bots.retrieve_creds import retrieve_creds_for_money_manager; " & _
    "creds = retrieve_creds_for_money_manager(); " & _
    "money_manager_obj = Money_Manager(creds); " & _
    "money_manager_obj.retrieve_estatements()"
    
    ' Run Python source code
    RunPython python_source_code

End Sub

Sub add_archived_posted_txns()

    ' This really isn't a huge time saver, but it's fine...

    ' Row index of first empty row after Posted txn data set
    first_empty_row = Worksheets("Posted Transactions").Range("A1").End(xlDown).Row + 1
    
    ' Move archived posted txns to that row
    Worksheets("Archived Posted Txn Data").Range("A1").CurrentRegion.Copy Worksheets("Posted Transactions").Range("A" & first_empty_row)
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub

Sub Add_Descriptions()

    ' Clear the old data
    Sheets("Income and Expenses").Range("A:J").Clear
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts_and_Trading_Bots.Money_Manager import Money_Manager; " & _
    "money_manager_obj = Money_Manager(); " & _
    "money_manager_obj.add_transaction_descriptions()"
    
    ' Run Python source code
    RunPython python_source_code
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub

Sub Scrape_Posted_Txns()

    ' Clear all of the data in the "Raw Posted Txn Data - Scraped" worksheet
    Sheets("Posted Transactions").Cells.Clear

    ' Prompt user for password to credential workbook
    otp = InputBox("Input", Title:="Robinhood OTP")
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts_and_Trading_Bots.Money_Manager import Money_Manager; " & _
    "from Scripts_and_Trading_Bots.retrieve_creds import retrieve_creds_for_money_manager; " & _
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
    Sheets("Investment Portfolio").Range("F:J").Clear
    
    ' Prompt user for password to credential workbook
    otp = InputBox("Input", Title:="Robinhood OTP")
    
    ' Define Python source code
    python_source_code = "" & _
    "from Scripts_and_Trading_Bots.Money_Manager import Money_Manager; " & _
    "from Scripts_and_Trading_Bots.retrieve_creds import retrieve_creds_for_money_manager; " & _
    "creds = retrieve_creds_for_money_manager(); " & _
    "money_manager_obj = Money_Manager(creds); " & _
    "otp = '" & otp & "'; " & _
    "money_manager_obj.get_investments(otp); "

    ' Run Python source code
    RunPython python_source_code
    
    ' Add two new records to the holdings Excel table for your staked crypto
    Set tbl = Sheets("Investment Portfolio").ListObjects("holdings")
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1).Formula = "=M14"
        .Range(2).Formula = "=M9"
        .Range(3).Formula = "=[@Quantity]*N8"
        .Range(4).Formula = "=N9"
        .Range(5).Value = "crypto"
    End With
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1).Formula = "=M14"
        .Range(2).Formula = "=M10"
        .Range(3).Formula = "=[@Quantity]*N8"
        .Range(4).Formula = "=N10"
        .Range(5).Value = "crypto"
    End With
    
    ' Update this formula
    Sheets("Investment Portfolio").Range("M12").Formula = "=SUM(holdings[equity],L6,M6)"
    
    ' Tell the user the macro is done
    MsgBox ("Done")

End Sub








