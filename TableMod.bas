Attribute VB_Name = "TableMod"
Option Explicit

'This module is dedicated to CRUD operations related to Tables.
'The array to Range functions are how we construct the table data

'********************************************************
'Simple array to sheet functions, assign range on sheet values to array values.
'********************************************************


'This sub takes in a 2d array of data, a worksheet destination and a cell
Sub Array2DToRange(InputArray As Variant, tgtWksheet As String, TgtCell As String)

    Worksheets(tgtWksheet).Range(TgtCell).Resize(UBound(InputArray, 1), UBound(InputArray, 2)).Value = InputArray
    
End Sub


'This sub takes in a 1d array of data, a worksheet destination and a cell and writes it as a column

Sub Array1DToRangeCol(InputArray As Variant, tgtWksheet As String, TgtCell As String)
    
    Worksheets(tgtWksheet).Range(TgtCell).Resize(UBound(InputArray, 1), 1).Value = Application.Transpose(InputArray)

End Sub


'This sub takes in a 1d array of data, a worksheet destination and a cell and writes it as a row
Sub Array1DToRangeRow(InputArray As Variant, tgtWksheet As String, TgtCell As String)
    
    Worksheets(tgtWksheet).Range(TgtCell).Resize(1, UBound(InputArray, 1)).Value = InputArray

End Sub

'********************************************************************************
'Composite array to sheet functions, assign range values to array values over multiple sheets.
'********************************************************************************

'Set the elapsed time on all data sheets with non-empty TestData arrays
Sub TimeArrayToDataSheets(InputArray As Variant, TgtCell As String)

        If Not IsEmpty(DataFileMod.TestData.analogData) Then
            Sheets("AnalogData").Range("A1").Value = "Elapsed Time"
            Call Array1DToRangeCol(InputArray, "AnalogData", TgtCell)
        End If
        
        'Typical count and analog data share elapsed time, but cycle time has a different frequency
        If Not IsEmpty(DataFileMod.TestData.cycleAnalogData) Then
            Sheets("CycleAnalogData").Range("A1").Value = "Elapsed Time"
            Call Array1DToRangeCol(TestData.CycleTimes, "CycleAnalogData", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LB_Sizes) Then
            Sheets("LB_Up_Counts").Range("A1").Value = "Elapsed Time"
            Sheets("LB_Down_Counts").Range("A1").Value = "Elapsed Time"
            Call Array1DToRangeCol(InputArray, "LB_Up_Counts", TgtCell)
            Call Array1DToRangeCol(InputArray, "LB_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LBE_Sizes) Then
            Sheets("LBE_Down_Counts").Range("A1").Value = "Elapsed Time"
            Call Array1DToRangeCol(InputArray, "LBE_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LS_Sizes) Then
            Sheets("LS_Up_Counts").Range("A1").Value = "Elapsed Time"
            Sheets("LS_Down_Counts").Range("A1").Value = "Elapsed Time"
            Call Array1DToRangeCol(InputArray, "LS_Up_Counts", TgtCell)
            Call Array1DToRangeCol(InputArray, "LS_Down_Counts", TgtCell)
        End If
End Sub

'This sub takes all of the arrays which have data in the class module
'Then writes them to their respective worksheets.
Sub TestDataToSheets(TgtCell As String)
        
        If Not IsEmpty(DataFileMod.TestData.analogData) Then
            Call Array2DToRange(TestData.analogData, "AnalogData", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.cycleAnalogData) Then
            Call Array2DToRange(TestData.cycleAnalogData, "CycleAnalogData", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LB_Sizes) Then
            Call Array2DToRange(TestData.LBU_CountsData, "LB_Up_Counts", TgtCell)
            Call Array2DToRange(TestData.LBD_CountsData, "LB_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LBE_Sizes) Then
            Call Array2DToRange(TestData.LBE_CountsData, "LBE_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LS_Sizes) Then
            Call Array2DToRange(TestData.LSU_CountsData, "LS_Up_Counts", TgtCell)
            Call Array2DToRange(TestData.LSD_CountsData, "LS_Down_Counts", TgtCell)
        End If
End Sub

Sub DataTagsToSheets(TgtCell As String)
        If Not IsEmpty(DataFileMod.TestData.AnalogTags) Then
            Call Array1DToRangeRow(TestData.AnalogTags, "AnalogData", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.CycleAnalogTags) Then
            Call Array1DToRangeRow(TestData.CycleAnalogTags, "CycleAnalogData", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LB_Sizes) Then
            Call Array1DToRangeRow(TestData.LB_Sizes, "LB_Up_Counts", TgtCell)
            Call Array1DToRangeRow(TestData.LB_Sizes, "LB_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LBE_Sizes) Then
            Call Array1DToRangeRow(TestData.LBE_Sizes, "LBE_Down_Counts", TgtCell)
        End If
        
        If Not IsEmpty(DataFileMod.TestData.LS_Sizes) Then
            Call Array1DToRangeRow(TestData.LS_Sizes, "LS_Up_Counts", TgtCell)
            Call Array1DToRangeRow(TestData.LS_Sizes, "LS_Down_Counts", TgtCell)
        End If
End Sub


'********************************************************************************
'Simple table functions
'********************************************************************************
'This cleans up the worksheets so we can place new data.
'This will also hide empty tabs, which will be revealed when filled.
Sub DeleteDataTables(TgtCell As String)
        
    If Not IsEmpty(Sheets("HeaderData").Range(TgtCell)) Then
        Sheets("HeaderData").UsedRange.Clear
        Sheets("HeaderData").Visible = xlSheetHidden
    End If
    
    If Not IsEmpty(Sheets("AnalogData").Range(TgtCell)) Then
        Sheets("AnalogData").Range(TgtCell).CurrentRegion.Clear
        Sheets("AnalogData").Visible = xlSheetHidden
    End If
    
    If Not IsEmpty(Sheets("CycleAnalogData").Range(TgtCell)) Then
        Sheets("CycleAnalogData").Range(TgtCell).CurrentRegion.Clear
        Sheets("CycleAnalogData").Visible = xlSheetHidden
    End If
    
    If Not IsEmpty(Sheets("LB_Up_Counts").Range(TgtCell)) Then
        Sheets("LB_Up_Counts").Range(TgtCell).CurrentRegion.Clear
        Sheets("LB_Up_Counts").Visible = xlSheetHidden
        Sheets("LB_Down_Counts").Range(TgtCell).CurrentRegion.Clear
        Sheets("LB_Down_Counts").Visible = xlSheetHidden
    End If
    
    If Not IsEmpty(Sheets("LBE_Down_Counts").Range(TgtCell)) Then
        Sheets("LBE_Down_Counts").Range(TgtCell).CurrentRegion.Clear
        Sheets("LBE_Down_Counts").Visible = xlSheetHidden
    End If
    
    If Not IsEmpty(Sheets("LS_Up_Counts").Range(TgtCell)) Then
        Sheets("LS_Up_Counts").Range(TgtCell).CurrentRegion.Clear
        Sheets("LS_Up_Counts").Visible = xlSheetHidden
        Sheets("LS_Down_Counts").Range(TgtCell).CurrentRegion.Clear
        Sheets("LS_Down_Counts").Visible = xlSheetHidden
    End If

End Sub



Public Sub DisposeData(TgtCell As String)
    Call TableMod.DeleteDataTables(TgtCell)
    Sheets("HeaderData").UsedRange.Clear
    Set DataFileMod.TestData = Nothing
End Sub

'This takes in a worksheet and a cell where a table is expected to be found
'The expectation is that you've written data to the sheet programmatically in a contiguous blob
'The function will expand the cell reference to the entire range, then convert that range into a named table for data operations.
Sub CreateTable(ByVal tgtWksheet As String, ByVal TgtCell As String, Optional ByVal TableName As String = "", Optional AvgCalc As Boolean)
    Dim rng As Range
    Dim tbl As ListObject
    
    Set rng = Sheets(tgtWksheet).Range(TgtCell).CurrentRegion
    
    Set tbl = Sheets(tgtWksheet).ListObjects.Add(xlSrcRange, rng, , xlYes)
    
    If TableName <> "" Then
        tbl.Name = Replace(TableName, ";", "") ' Use the provided table name but remove semi colon leader
    Else
        tbl.Name = Sheets(tgtWksheet).Name & "Tbl" ' Use the sheet name with "Table" suffix
    End If
    
    ' Show headers and remove autofilter dropdown
    tbl.ShowHeaders = True
    tbl.ShowAutoFilterDropDown = False
    tbl.ShowTableStyleRowStripes = False
    tbl.TableStyle = "TableStyleMedium20"
    
    tbl.HeaderRowRange.HorizontalAlignment = xlCenter
    tbl.DataBodyRange.HorizontalAlignment = xlCenter

    ' Show average rows
    If AvgCalc = True Then
        tbl.ShowTotals = True
        Dim col As ListColumn
        For Each col In tbl.ListColumns
            If (col.index > 2) Then
            col.TotalsCalculation = xlTotalsCalculationAverage
            Else
            col.TotalsCalculation = xlTotalsCalculationNone
            End If
        Next col
    End If
    
    'Show tab containing table
    Sheets(tgtWksheet).Visible = xlSheetVisible

    
End Sub

'********************************************************************************
'Composite table functions
'********************************************************************************

Sub ConvertDataToNamedTables(TgtCell As String)
        If Not IsEmpty(Sheets("AnalogData").Range(TgtCell)) Then
            Call CreateTable("AnalogData", "A1")
        End If
        
        If Not IsEmpty(Sheets("CycleAnalogData").Range(TgtCell)) Then
            Call CreateTable("CycleAnalogData", "A1")
        End If
        
        If Not IsEmpty(Sheets("LB_Up_Counts").Range(TgtCell)) Then
                    Call CreateTable("LB_Up_Counts", "A1")
                    Call CreateTable("LB_Down_Counts", "A1")
        End If
        
        If Not IsEmpty(Sheets("LBE_Down_Counts").Range(TgtCell)) Then
                    Call CreateTable("LBE_Down_Counts", "A1")
        End If
        
        If Not IsEmpty(Sheets("LS_Up_Counts").Range(TgtCell)) Then
                    Call CreateTable("LS_Up_Counts", "A1")
                    Call CreateTable("LS_Down_Counts", "A1")
        End If
End Sub

'This function takes the HeaderData Array saved in the DataFile class module, chops it into sections, then creates tables of those sections on the target tab.
'This function should be refactored into 2 or 3 subs down the road.
Sub HeaderDataToSheet(TgtSheet As String)
    Dim SectNames() As Variant ' Row 0 will be filled with section names, Row 1 will be filled with section indexes
    Dim SectCount As Integer ' This will keep the number of sections
    
    'loop counters
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    If Not IsEmpty(DataFileMod.TestData.HeaderData) Then
        For i = 0 To UBound(DataFileMod.TestData.HeaderData, 2) - 1
            If Left(DataFileMod.TestData.HeaderData(0, i), 1) = ";" Then
                SectCount = SectCount + 1
                'Expand the Section Name array each time a section name is found
                ReDim Preserve SectNames(1 To 2, 1 To SectCount)
                'Save the section name
                SectNames(1, SectCount) = DataFileMod.TestData.HeaderData(0, i)
                'Save the section name index
                SectNames(2, SectCount) = i
            End If
        Next i
    End If
    
    Dim SectStart As Integer
    Dim SectEnd As Integer
    Dim SectArray() As Variant ' This array is used as temporary storage to extract a header section and write it to the target sheet
    Dim WriteIndex As Integer 'This will store the row to write the next header section table
    
    SectStart = 0
    SectEnd = 0
    SectCount = 0
    WriteIndex = 1
    
    If Not IsEmpty(SectNames) Then
        'loop through each section
        For i = 0 To UBound(SectNames, 2) - 1 ' Updated dimension
                
                'Check for last section, if so then set max column to end of headerdata array
                If i = UBound(SectNames, 2) - 1 Then  ' Last section
                    SectStart = SectNames(2, i + 1) + 1
                    SectEnd = UBound(DataFileMod.TestData.HeaderData, 2)
                Else
                    SectStart = SectNames(2, i + 1) + 1 'typical sections
                    SectEnd = SectNames(2, i + 2) - 1
                End If
                
                ' Redim SectArray
                ReDim SectArray(0 To UBound(DataFileMod.TestData.HeaderData, 2), 0 To (SectEnd - SectStart + 1))
            
            'loop through all the columns of the header array between the two section name indexes
            For j = SectStart To SectEnd
                
                'loop through all the rows of the area between each section name in the headerdata array
                For k = 0 To UBound(DataFileMod.TestData.HeaderData, 1)
                    
                    'Save the headerdata values into the SectArray
                    SectArray(k, j - SectStart) = DataFileMod.TestData.HeaderData(k, j)
    
                Next k
                
            Next j
            'label the table on the sheet
            With Sheets(TgtSheet).Range("A" & WriteIndex)
            .Value = Replace(SectNames(1, i + 1), ";", "")
            .Font.Bold = True
            .Font.Underline = xlUnderlineStyleSingle
            End With
            
            WriteIndex = WriteIndex + 2
            
            'write temp array to headerdata tab at WriteIndex
            Call Array2DToRange(SectArray, TgtSheet, "A" & WriteIndex)
            
            'convert data in tab to table named after section at row index
            Call TableMod.CreateTable(TgtSheet, "A" & WriteIndex, SectNames(1, i + 1))
            
            'increment writeIndex to end of  + 1 empty row
            WriteIndex = Sheets(TgtSheet).Cells(Rows.count, 1).End(xlUp).Row + 2
        Next i
    End If
    
End Sub

'********************************************************************************
'Table Read Function
'********************************************************************************

'WkSheet As String, TblName As String, TblKey As String)
Function GetArrayFromTable(wkSheet As String, tblName As String, TblKey As String) As Variant
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim columnRange As Range
    Dim dataArray As Variant
    Dim counter As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = Sheets(wkSheet)
    Set tbl = ws.ListObjects(tblName)
    
    Set columnRange = tbl.ListColumns(TblKey).DataBodyRange

    dataArray = columnRange.Value
        
    GetArrayFromTable = dataArray

ExitFunction:
    Exit Function

ErrorHandler:
    ' Set value to "Not Found" if an error occurs
    GetArrayFromTable = "Not Found"
    Resume ExitFunction
    
End Function

'WkSheet As String, TblName As String, TblKey As String)
Public Function GetValueFromTable(wkSheet As String, tblName As String, TblKey As String, ValueIndex As Integer) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim columnRange As Range
    Dim dataArray As Variant
    
    On Error GoTo ErrorHandler
    
    Set ws = Sheets(wkSheet)
    Set tbl = ws.ListObjects(tblName)
    
    GetValueFromTable = tbl.ListColumns(TblKey).DataBodyRange(ValueIndex, 1)
    
ExitFunction:
    Exit Function

ErrorHandler:
    ' Set value to "Not Found" if an error occurs
    GetValueFromTable = "Not Found"
    Resume ExitFunction
    
End Function

Function GetValueFromTableObj(tbl As ListObject, TblKey As String, ValueIndex As Long) As Variant
    On Error GoTo ErrorHandler
    GetValueFromTableObj = tbl.ListColumns(TblKey).DataBodyRange(ValueIndex, 1)
    Exit Function
ErrorHandler:
    GetValueFromTableObj = "Not Found"
End Function

Sub SetValueInTableObj(tbl As ListObject, TblKey As String, ValueIndex As Long, newValue As Variant)
    On Error GoTo ErrorHandler
    tbl.ListColumns(TblKey).DataBodyRange(ValueIndex, 1).Value = newValue
    Exit Sub
ErrorHandler:
    MsgBox "SetValueInTable Error: cell not found or other error.", vbExclamation
End Sub


Sub GroupTableData(wkSheet As String, TableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Sheets(wkSheet)
    Set tbl = ws.ListObjects(TableName)

    ' Group the Data range leaving the headers for viewing
    tbl.DataBodyRange.Rows.Group
    
End Sub



Sub CollapseTableData(wkSheet As String)

    Dim ws As Worksheet
    
    ' Set the worksheet object
    Set ws = Sheets(wkSheet)
    
    ' Collapse all groups in the worksheet to level 1
    ws.Outline.ShowLevels RowLevels:=1
    
End Sub
