Attribute VB_Name = "Charts"
Option Explicit

' Note: OptimizePerformance is now centralized in DevToolsMod module

'----------------------------------------------------------------------------------------
' FormatChart - Optimized version with improved error handling and performance
'----------------------------------------------------------------------------------------
Sub FormatChart(wkSheet As String, chartName As String, TableName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim parameterArray(1 To 11) As Variant
    Dim i As Integer
    
    Call DevToolsMod.OptimizePerformance(True)
    
    ' Validate worksheet
    Set ws = ThisWorkbook.Worksheets(wkSheet)
    If ws Is Nothing Then
        MsgBox "Worksheet '" & wkSheet & "' not found.", vbExclamation
        GoTo ExitHandler
    End If

    ' Validate chart - try by name first, then by index if name fails
    On Error Resume Next
    Set chartObj = ws.ChartObjects(chartName)
    If chartObj Is Nothing Then
        Set chartObj = ws.ChartObjects(1) ' Fallback to first chart
    End If
    On Error GoTo ErrorHandler
    
    If chartObj Is Nothing Then
        MsgBox "Chart '" & chartName & "' not found on '" & wkSheet & "'.", vbExclamation
        GoTo ExitHandler
    End If

    ' Bulk load all parameters in single operation instead of 11 individual calls
    For i = 1 To 11
        parameterArray(i) = GetChartSaveResult(wkSheet, TableName, i)
    Next i
    
    ' Apply formatting with single chart reference
    With chartObj.Chart
        ' Apply titles
        .HasTitle = True
        .ChartTitle.Text = parameterArray(1)  ' GraphTitle
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = parameterArray(2)  ' YAxisTitle
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = parameterArray(3)  ' XAxisTitle

        ' ==== Y AXIS ====
        Call FormatAxis(.Axes(xlValue), CBool(parameterArray(4)), parameterArray(5), parameterArray(6), parameterArray(7))
        
        ' ==== X AXIS ====
        Call FormatAxis(.Axes(xlCategory), CBool(parameterArray(8)), parameterArray(9), parameterArray(10), parameterArray(11))
    End With
    
ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Set ws = Nothing
    Set chartObj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error formatting chart '" & chartName & "': " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' FormatAxis - Helper function to reduce code duplication
'----------------------------------------------------------------------------------------
Private Sub FormatAxis(axis As axis, isLog As Boolean, minVal As Variant, maxVal As Variant, majorVal As Variant)
    On Error GoTo ErrorHandler
    
    With axis
        If isLog Then
            .ScaleType = xlScaleLogarithmic
            ' Min value
            If Not IsEmpty(minVal) And minVal <> "" And IsNumeric(minVal) And minVal > 0 Then
                .MinimumScale = 10 ^ Application.WorksheetFunction.RoundDown(Application.WorksheetFunction.Log10(minVal), 0)
                .MinimumScaleIsAuto = False
            Else
                .MinimumScaleIsAuto = True
            End If
            ' Max value
            If Not IsEmpty(maxVal) And maxVal <> "" And IsNumeric(maxVal) And maxVal > 0 Then
                .MaximumScale = 10 ^ Application.WorksheetFunction.RoundUp(Application.WorksheetFunction.Log10(maxVal), 0)
                .MaximumScaleIsAuto = False
            Else
                .MaximumScaleIsAuto = True
            End If
            ' Major unit
            If Not IsEmpty(majorVal) And majorVal <> "" And IsNumeric(majorVal) And majorVal > 0 Then
                .MajorUnit = 10 ^ Application.WorksheetFunction.RoundDown(Application.WorksheetFunction.Log10(majorVal), 0)
                .MajorUnitIsAuto = False
            Else
                .MajorUnitIsAuto = True
            End If
        Else
            .ScaleType = xlScaleLinear
            ' Min value
            If Not IsEmpty(minVal) And minVal <> "" And IsNumeric(minVal) Then
                .MinimumScale = minVal
                .MinimumScaleIsAuto = False
            Else
                .MinimumScaleIsAuto = True
            End If
            ' Max value
            If Not IsEmpty(maxVal) And maxVal <> "" And IsNumeric(maxVal) Then
                .MaximumScale = maxVal
                .MaximumScaleIsAuto = False
            Else
                .MaximumScaleIsAuto = True
            End If
            ' Major unit
            If Not IsEmpty(majorVal) And majorVal <> "" And IsNumeric(majorVal) Then
                .MajorUnit = majorVal
                .MajorUnitIsAuto = False
            Else
                .MajorUnitIsAuto = True
            End If
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    ' Silently handle axis formatting errors to prevent cascade failures
End Sub

'----------------------------------------------------------------------------------------
' SetChartSeriesByRange - Optimized version with better error handling
'----------------------------------------------------------------------------------------
Sub SetChartSeriesByRange(wkSheet As String, chartName As String, TopLeftAddress As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim chrt As ChartObject
    Dim topLeft As Range
    Dim lastCol As Long, lastRow As Long
    Dim dataRange As Range
    Dim xData As Range, yData As Range
    Dim hdrCell As Range
    Dim refName As String
    Dim iCol As Long
    Dim i As Long
    
    Call DevToolsMod.OptimizePerformance(True)

    ' Set worksheet containing chart and reference cell
    Set ws = ThisWorkbook.Worksheets(wkSheet)
    Set chrt = ws.ChartObjects(chartName)
    Set topLeft = ws.Range(TopLeftAddress)

    ' Find the last column and row of contiguous data
    lastCol = topLeft.End(xlToRight).Column
    lastRow = topLeft.End(xlDown).Row

    ' Set the full data range (including headers)
    Set dataRange = ws.Range(topLeft, ws.Cells(lastRow, lastCol))

    ' Clear all existing series in single loop (more efficient than individual deletes)
    With chrt.Chart
        For i = .SeriesCollection.count To 1 Step -1
            .SeriesCollection(i).Delete
        Next i
    End With

    ' Define X-data (first column, below header) - single range reference
    Set xData = dataRange.Columns(1).offset(1, 0).Resize(dataRange.Rows.count - 1, 1)

    ' Loop through each Y series (columns 2 to N) with optimized range operations
    For iCol = 2 To dataRange.Columns.count
        Set yData = dataRange.Columns(iCol).offset(1, 0).Resize(dataRange.Rows.count - 1, 1)
        Set hdrCell = dataRange.Cells(1, iCol)
        refName = "='" & ws.Name & "'!" & hdrCell.Address(ReferenceStyle:=xlA1)
        
        With chrt.Chart.SeriesCollection.NewSeries
            .xValues = xData
            .values = yData
            .Name = refName   ' Reference the header cell dynamically
        End With
    Next iCol
    
ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Set ws = Nothing
    Set chrt = Nothing
    Set topLeft = Nothing
    Set dataRange = Nothing
    Set xData = Nothing
    Set yData = Nothing
    Set hdrCell = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error setting chart series: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' GetChartSaveResult - Optimized version with better error handling
'----------------------------------------------------------------------------------------
Public Function GetChartSaveResult(wkSheet As String, TableName As String, ID As Integer) As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    
    Set ws = ThisWorkbook.Worksheets(wkSheet)
    Set tbl = ws.ListObjects(TableName)
    
    ' Check columns 3 and 4 for values (User_Entry has priority over Calculated)
    For i = 3 To 4
        If Not IsEmpty(tbl.DataBodyRange(ID, i).Value) And tbl.DataBodyRange(ID, i).Value <> "" Then
            GetChartSaveResult = tbl.DataBodyRange(ID, i).Value
            GoTo ExitFunction
        End If
    Next i
    
    ' If no value found, return empty
    GetChartSaveResult = Empty
        
ExitFunction:
    Set ws = Nothing
    Set tbl = Nothing
    Exit Function

ErrorHandler:
    GetChartSaveResult = Empty
    Resume ExitFunction
End Function

'----------------------------------------------------------------------------------------
' UpdateCharts - Optimized version with better performance control
'----------------------------------------------------------------------------------------
Public Sub UpdateCharts()
    On Error GoTo ErrorHandler
    
    Call DevToolsMod.OptimizePerformance(True)
    
    Application.Calculate
    
    
    Call SetISO16889C1DPvMassSI
    Call SetISO16889C2SizevBetaSI
    Call SetISO16889C3TimevBeta
    Call SetISO16889C4PressureSIvBeta
    Call SetISO16889C5UpCountsVsTime
    Call SetISO16889C6DnCountsVsTime

ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating charts: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' NiceRoundUp - Enhanced version with better performance
'----------------------------------------------------------------------------------------
Function NiceRoundUp(x As Double, Optional isLogScale As Boolean = False) As Double
    On Error GoTo ErrorHandler
    
    Dim niceVals As Variant
    Dim i As Integer
    Dim exp As Long, base As Double, mult As Double
    
    If x <= 0 Then
        NiceRoundUp = 0
        Exit Function
    End If
    
    If isLogScale Then
        ' For log scale: next higher power of 10
        exp = Application.WorksheetFunction.Ceiling_Math(Log(x) / Log(10), 1)
        NiceRoundUp = 10 ^ exp
    Else
        ' For linear: use extended "nice" series
        niceVals = Array(1, 2, 2.5, 5, 6, 8, 10, 12, 15, 20, 25, 30, 40, 50, 60, 75, 100)
        exp = Int(Log(x) / Log(10))
        base = 10 ^ exp
        For i = LBound(niceVals) To UBound(niceVals)
            mult = niceVals(i) * base
            If mult >= x Then
                NiceRoundUp = mult
                Exit Function
            End If
        Next i
        ' Fallback in case x is huge
        NiceRoundUp = 10 * base
    End If
    
    Exit Function
    
ErrorHandler:
    NiceRoundUp = x ' Return original value on error
End Function

