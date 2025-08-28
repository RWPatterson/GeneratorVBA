Attribute VB_Name = "ReportFillMod"
Option Explicit
'***************************************************************************************************************
'This module is responsible for filling the report page labels and data from the SaveData and Language Tables.
'***************************************************************************************************************


'Replaced L_Con in the new report writer versions using data in a table and more modern select method.
'Supply the given ID and it will return the label based on the current language setting.

'Old L_Con 10,000 requests in 2.12 seconds
'GetLanguageResult 10,000 requests in .35s
Public Function GetLanguageResult(ID As Integer) As String
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim LangColumn As Integer
        
On Error GoTo ErrorHandler
        
    Set ws = ThisWorkbook.Worksheets("Translations_Table") 'This function is hardcoded to use the Save_Data Sheet.
    Set tbl = ws.ListObjects("TranslationsDataTable") 'The SaveDataTable is always present in the workbook.
    LangColumn = GetSaveResult(57) + 1 'Get the current language offset
    
    GetLanguageResult = tbl.DataBodyRange(ID, LangColumn).Value
    
ExitFunction:
    Exit Function

ErrorHandler:
    ' Set value to "Not Found" if an error occurs
    GetLanguageResult = "Not Found"
    Resume ExitFunction
    
End Function

' Helper function to check if worksheet exists
Public Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(wsName).Name = wsName)
    On Error GoTo 0
End Function

''This sub returns the value of a field from the Save_Data table.
''The Save_Data table has an ID, a Display Name, a User Entry, a Custom Default, and a From Data column
''The priority of return is:
'    'User_Entry is highest
'    'Custom_Default fallback
'    'From_RawDat last.
'    'if no value is present in all 3 columns, return 0.
Public Function GetSaveResult(ID As Long) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("SaveDataTable")
    
    ' This just returns what's already in the Report Value column
    GetSaveResult = tbl.DataBodyRange(ID, 3).Value
    
ExitFunction:
    Exit Function

ErrorHandler:
    ' Do nothing on an error, dev probably has an ID that doesn't exist.
    Resume ExitFunction
End Function

'This sub sets the value of a user entry field in the Save_Data table.
'The Save_Data table has an ID, a Display Name, a User Entry, a Custom Default, and a From Data column

'Optional SaveValue to supply value to be saved.
Public Sub SetSaveUserEntry(ID As Integer, SaveValue As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data") 'This function is hardcoded to use the Save_Data Sheet.
    Set tbl = ws.ListObjects("SaveDataTable") 'The SaveDataTable is always present in the workbook.

    tbl.DataBodyRange(ID, 4).Value = SaveValue
        
ExitFunction:
    Exit Sub

ErrorHandler:
    ' Do nothing on an error, dev probably has an ID that doesn't exist.
    Resume ExitFunction
End Sub

'This sub sets the value of a custom default field in the Save_Data table.
'The Save_Data table has an ID, a Display Name, a User Entry, a Custom Default, and a From Data column

'Optional SaveValue to supply value to be saved.
Public Sub SetSaveCustomDefault(ID As Integer, SaveValue As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data") 'This function is hardcoded to use the Save_Data Sheet.
    Set tbl = ws.ListObjects("SaveDataTable") 'The SaveDataTable is always present in the workbook.
    
    tbl.DataBodyRange(ID, 5).Value = SaveValue
        
ExitFunction:
    Exit Sub

ErrorHandler:
    ' Do nothing on an error, dev probably has an ID that doesn't exist.
    Resume ExitFunction
End Sub




' Helper function to check if a Variant array variable has data
Public Function hasData(arr As Variant) As Boolean
    On Error GoTo NoData
    If IsArray(arr) Then
        hasData = (UBound(arr) >= LBound(arr))
    Else
        hasData = False
    End If
    Exit Function
    
NoData:
    hasData = False
End Function

    


'Locates the last used row in A
Function GetLastRow(Worksheet As String) As Integer

GetLastRow = ThisWorkbook.Worksheets(Worksheet).Range("A" & Rows.count).End(xlUp).Row

End Function
