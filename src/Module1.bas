Attribute VB_Name = "Module1"
Option Explicit

Dim otp As String
Dim tbl As ListObject
Dim newrow As ListRow
Dim first_empty_row As Integer
Dim python_source_code As String

Private Function PyBootstrap() As String
    PyBootstrap = "" & _
    "import sys, importlib; " & _
    "repo_parent = r'D:\...'; " & _ ' specify this path
    "repo_root = r'D:\...\personal_finance_data_pipeline'; " & _ ' specify this path
    "sys.path.insert(0, repo_parent) if repo_parent not in sys.path else None; " & _
    "sys.path.insert(0, repo_root) if repo_root not in sys.path else None; " & _
    "from personal_finance_data_pipeline.src.personal_finance_data_pipeline import PersonalFinanceDataPipeline; "
End Function

Private Function PyCredsLoader() As String
    PyCredsLoader = "" & _
    "_m = importlib.import_module('retrieve_creds'); " & _
    "_get_creds = getattr(_m, 'retrieve_creds_for_money_manager', None); " & _
    "_get_creds = _get_creds if _get_creds is not None else getattr(_m, 'retrieve_creds_for_PersonalFinanceDataPipeline'); " & _
    "creds = _get_creds(); "
End Function

Sub Retrieve_Account_Data_and_Transactions()

    Sheets("All FirstBank Transactions").Range("A1").CurrentRegion.Cells.Clear

    python_source_code = PyBootstrap() & _
        PyCredsLoader() & _
        "pipeline = PersonalFinanceDataPipeline(creds); " & _
        "pipeline.retrieve_account_data_and_transactions(); "

    RunPython python_source_code

    MsgBox ("Done")

End Sub

Sub Refresh_Income_and_Expenses()

    python_source_code = PyBootstrap() & _
        "pipeline = PersonalFinanceDataPipeline(); " & _
        "pipeline.refresh_income_and_expense_data(); "

    RunPython python_source_code

    MsgBox ("Done")

End Sub

Sub Refresh_Investment_Portfolio()

    Sheets("Personal Investment Portfolio").Range("holdings").Clear

    python_source_code = PyBootstrap() & _
        PyCredsLoader() & _
        "pipeline = PersonalFinanceDataPipeline(creds); " & _
        "pipeline.get_investments_v1(); "

    RunPython python_source_code

    Sheets("Overview").Range("personal_investment_portfolio").Formula = "=SUM(holdings[Current Equity])+coinbase_usd_cash_bal"

    MsgBox ("Done")

End Sub

Sub Retrieve_eStatements()

    python_source_code = PyBootstrap() & _
        PyCredsLoader() & _
        "pipeline = PersonalFinanceDataPipeline(creds); " & _
        "pipeline.retrieve_estatements(); "

    RunPython python_source_code

    MsgBox ("Done")

End Sub




