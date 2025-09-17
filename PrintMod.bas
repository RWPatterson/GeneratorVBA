Attribute VB_Name = "PrintMod"
Option Explicit

Sub PrintSelectedSheets()
    Dim wsPrintTable As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim printSheetName As String
    Dim printFlag As Boolean
    Dim sheetsToPrint As Sheets
    Dim shNames() As String
    Dim cnt As Long
    
    ' Set the worksheet containing the table
    Set wsPrintTable = ThisWorkbook.Sheets("Save_Data")
    
    ' Set PrintFormatting to 1 (show entry colors)
    ThisWorkbook.Names("PrintFormatting").RefersToRange.value = 1

    ' Get the table object
    On Error Resume Next
    Set tbl = wsPrintTable.ListObjects("ISO16889PrintTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Table 'ISO16889PrintTable' not found on sheet '" & wsPrintTable.Name & "'.", vbCritical
        ThisWorkbook.Names("PrintFormatting").RefersToRange.value = 0
        Exit Sub
    End If
    
    cnt = 0
    For i = 1 To tbl.ListRows.count
        printFlag = tbl.DataBodyRange(i, tbl.ListColumns("Print? True/False").index).value
        If printFlag = True Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        MsgBox "No sheets are marked for printing.", vbInformation
        ThisWorkbook.Names("PrintFormatting").RefersToRange.value = 0
        Exit Sub
    End If
    
    ReDim shNames(1 To cnt)
    cnt = 0
    For i = 1 To tbl.ListRows.count
        printFlag = tbl.DataBodyRange(i, tbl.ListColumns("Print? True/False").index).value
        If printFlag = True Then
            printSheetName = tbl.DataBodyRange(i, tbl.ListColumns("Display Name").index).value
            On Error Resume Next
            Dim sht As Worksheet
            Set sht = ThisWorkbook.Sheets(printSheetName)
            On Error GoTo 0
            If Not sht Is Nothing Then
                cnt = cnt + 1
                shNames(cnt) = printSheetName
                Set sht = Nothing
            Else
                MsgBox "Sheet '" & printSheetName & "' not found. It will be skipped.", vbExclamation
            End If
        End If
    Next i
    
    If cnt = 0 Then
        MsgBox "No valid sheets found to print.", vbInformation
        ThisWorkbook.Names("PrintFormatting").RefersToRange.value = 0
        Exit Sub
    End If
    
    ThisWorkbook.Sheets(shNames).PrintPreview
    
    MsgBox "Print preview process completed.", vbInformation
    ' Reset PrintFormatting back to 0 (hide entry colors)
    ThisWorkbook.Names("PrintFormatting").RefersToRange.value = 0
End Sub
