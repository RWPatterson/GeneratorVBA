VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmchart 
   Caption         =   "Modify Charts"
   ClientHeight    =   12780
   ClientLeft      =   15
   ClientTop       =   285
   ClientWidth     =   13860
   OleObjectBlob   =   "frmchart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChartInfo As Object
Private chartNeedsUpdate As Boolean


'----------------------------------------------------------------------------------------
' 1. UserForm_Initialize
' Purpose: Set up form controls, load chart selection list, initialize mapping between
'          chart names and their data locations/tables.
'----------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Set ChartInfo = CreateObject("Scripting.Dictionary")
    chartNeedsUpdate = False

    ' Map Excel tab names (as shown in dropdown) to array: [Worksheet, Table, ChartName]
    ChartInfo.Add "C1_DP_v_Mass", Array("C1_DP_v_Mass", "ISO16889C1SITable", "ISO16889Chart1SI")
    ChartInfo.Add "C2_Beta_v_Size", Array("C2_Beta_v_Size", "ISO16889C2Table", "ISO16889C2Chart")
    ChartInfo.Add "C3_Beta_v_Time", Array("C3_Beta_v_Time", "ISO16889C3Table", "ISO16889C3Chart")
    ChartInfo.Add "C4_Beta_v_Press", Array("C4_Beta_v_Press", "ISO16889C4SITable", "ISO16889C4Chart")
    ChartInfo.Add "C5_Up_Counts", Array("C5_Up_Counts", "ISO16889C5Table", "ISO16889C5UpCountsVsTime")
    ChartInfo.Add "C6_Down_Counts", Array("C6_Down_Counts", "ISO16889C6Table", "ISO16889C6DnCountsVsTimeChart")

    ' Populate dropdown with Excel tab names
    Dim key As Variant
    For Each key In ChartInfo.Keys
        Me.ChartChoice.AddItem key
    Next key

    ' Select first chart and load it by default
    If Me.ChartChoice.ListCount > 0 Then
        Me.ChartChoice.ListIndex = 0
        Call ChartChoice_Click ' This will load the first chart
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing form: " & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------------------------------
' 2. ChartChoice_Click
' Purpose: Handle when the user selects a chart from the dropdown, fetch all relevant
'          parameter values (original and user-entered) from the correct table/sheet and
'          populate controls on the form.
'----------------------------------------------------------------------------------------
Private Sub ChartChoice_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedTabName As String
    Dim chartInfoArray As Variant
    Dim wkSheet As String
    Dim tblName As String
    Dim chartName As String

    Call DevToolsMod.OptimizePerformance(True)

    ' Get selected tab name from dropdown
    selectedTabName = Me.ChartChoice.value

    If selectedTabName = "" Then
        MsgBox "Please select a chart from the list.", vbExclamation
        GoTo ExitHandler
    End If

    ' Verify tab exists in mapping
    If Not ChartInfo.Exists(selectedTabName) Then
        MsgBox "Selected tab not found in ChartInfo mapping.", vbCritical
        GoTo ExitHandler
    End If

    ' Single lookup - store result in array [Worksheet, Table, ChartName]
    chartInfoArray = ChartInfo(selectedTabName)
    wkSheet = chartInfoArray(0)
    tblName = chartInfoArray(1)
    chartName = chartInfoArray(2)

    ' Load the parameters into UserForm controls
    Call LoadChartParameters(wkSheet, tblName)

    ' Enable controls for editing as appropriate
    Me.ResetChartsCB.enabled = True
    Me.ModChartCB.enabled = True

    Me.NewMinX.Locked = False
    Me.NewMaxX.Locked = False
    Me.NewMajorX.Locked = False
    Me.NewXLog.Locked = False

    Me.NewMinY.Locked = False
    Me.NewMaxY.Locked = False
    Me.NewMajorY.Locked = False
    Me.NewYLog.Locked = False

    ' Update chart preview immediately when selection changes
    Call UpdateChartPreview(wkSheet, chartName)
    
ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error selecting chart: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' 3. LoadChartParameters
' Purpose: Generic sub to populate UserForm controls with both "From Data" (original)
'          and "User Entry" (user-modified) values for the selected chart.
'----------------------------------------------------------------------------------------
Private Sub LoadChartParameters(wkSheet As String, tblName As String)
    On Error GoTo ErrorHandler

    ' IDs per your rows 1 to 11
    Const ChartTitleID As Integer = 1
    Const YTitleID As Integer = 2
    Const XTitleID As Integer = 3
    Const YAxisLogID As Integer = 4
    Const YAxisMinID As Integer = 5
    Const YAxisMaxID As Integer = 6
    Const YAxisMajorTickID As Integer = 7
    Const XAxisLogID As Integer = 8
    Const XAxisMinID As Integer = 9
    Const XAxisMaxID As Integer = 10
    Const XAxisMajorTickID As Integer = 11

    ' Load all parameters in sequence for better performance
    Me.TitleChart.value = GetChartSaveResult(wkSheet, tblName, ChartTitleID)
    Me.TitleY.value = GetChartSaveResult(wkSheet, tblName, YTitleID)
    Me.TitleX.value = GetChartSaveResult(wkSheet, tblName, XTitleID)

    Me.YLog.value = GetChartSaveResult(wkSheet, tblName, YAxisLogID)
    Me.MinY.value = GetChartSaveResult(wkSheet, tblName, YAxisMinID)
    Me.MaxY.value = GetChartSaveResult(wkSheet, tblName, YAxisMaxID)
    Me.MajorY.value = GetChartSaveResult(wkSheet, tblName, YAxisMajorTickID)

    Me.XLog.value = GetChartSaveResult(wkSheet, tblName, XAxisLogID)
    Me.MinX.value = GetChartSaveResult(wkSheet, tblName, XAxisMinID)
    Me.MaxX.value = GetChartSaveResult(wkSheet, tblName, XAxisMaxID)
    Me.MajorX.value = GetChartSaveResult(wkSheet, tblName, XAxisMajorTickID)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading chart parameters: " & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------------------------------
' 4. SaveChartParameters
' Purpose: Take values from UserForm input controls and write them to the "User Entry"
'          column in the relevant parameter table for the selected chart.
'----------------------------------------------------------------------------------------
Private Sub SaveChartParameters(wkSheet As String, tblName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArray(1 To 11, 1 To 1) As Variant
    Dim currentValues As Variant
    Dim hasChanges As Boolean
    
    Set ws = Sheets(wkSheet)
    Set tbl = ws.ListObjects(tblName)
    
    ' Get current values to check for changes
    currentValues = tbl.DataBodyRange.Columns(3).value
    
    ' Build array with all new values
    dataArray(1, 1) = Me.TitleChart.value
    dataArray(2, 1) = Me.TitleY.value
    dataArray(3, 1) = Me.TitleX.value
    dataArray(4, 1) = IIf(Me.NewYLog.value, "TRUE", "FALSE")
    dataArray(5, 1) = Me.NewMinY.value
    dataArray(6, 1) = Me.NewMaxY.value
    dataArray(7, 1) = Me.NewMajorY.value
    dataArray(8, 1) = IIf(Me.NewXLog.value, "TRUE", "FALSE")
    dataArray(9, 1) = Me.NewMinX.value
    dataArray(10, 1) = Me.NewMaxX.value
    dataArray(11, 1) = Me.NewMajorX.value
    
    ' Check if any values actually changed
    Dim i As Integer
    For i = 1 To 11
        If IsEmpty(currentValues(i, 1)) Or currentValues(i, 1) <> dataArray(i, 1) Then
            hasChanges = True
            Exit For
        End If
    Next i
    
    ' Only write if there are actual changes
    If hasChanges Then
        tbl.DataBodyRange.Columns(3).value = dataArray
        chartNeedsUpdate = True
    End If
    
ExitHandler:
    Set ws = Nothing
    Set tbl = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving chart parameters: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' 5. ResetChartParameters
' Purpose: Clear/restore all user-modified entries in the parameter table for the selected
'          chart, reverting UserForm controls to the original "From Data" values.
'----------------------------------------------------------------------------------------
Private Sub ResetChartParameters(wkSheet As String, tblName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = Sheets(wkSheet)
    Set tbl = ws.ListObjects(tblName)
    
    ' Clear all data under "User Entry" column (column 3) in single operation
    tbl.DataBodyRange.Columns(3).ClearContents
    
    ' Mark chart for update
    chartNeedsUpdate = True
    
ExitHandler:
    Set ws = Nothing
    Set tbl = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error resetting chart parameters: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' 6. UpdateChartPreview
' Purpose: Refresh the chart preview image on the form by exporting the most recent chart
'          from the relevant worksheet, after any changes.
'----------------------------------------------------------------------------------------
Private Sub UpdateChartPreview(wsName As String, chartName As String)
    On Error GoTo ErrorHandler
    
    Dim CurrentChart As Chart
    Dim Fname As String
    Dim MyPath As String

    MyPath = ThisWorkbook.Path
    Fname = MyPath & "\chartPreview.gif"

    Set CurrentChart = Worksheets(wsName).ChartObjects(1).Chart

    ' Export chart as gif
    CurrentChart.Export FileName:=Fname, FilterName:="GIF"

    ' Display image
    Me.Image1.Picture = LoadPicture(Fname)
    
    ' Reset flag
    chartNeedsUpdate = False
    
ExitHandler:
    Set CurrentChart = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating chart preview: " & Err.Description, vbCritical
    chartNeedsUpdate = False ' Reset flag even on error
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' 11. ResetChartsCB_Click
' Purpose: Handler for "Reset" button; calls ResetChartParameters and updates the form.
'----------------------------------------------------------------------------------------
Private Sub ResetChartsCB_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedTabName As String
    Dim chartInfoArray As Variant
    Dim wkSheet As String
    Dim tblName As String
    Dim chartName As String
    
    Call DevToolsMod.OptimizePerformance(True)
    
    selectedTabName = Me.ChartChoice.value
    
    If selectedTabName = "" Then
        MsgBox "Please select a chart from the list before resetting.", vbExclamation
        GoTo ExitHandler
    End If
    
    If Not ChartInfo.Exists(selectedTabName) Then
        MsgBox "Selected tab not found in ChartInfo mapping.", vbCritical
        GoTo ExitHandler
    End If
    
    ' Single lookup [Worksheet, Table, ChartName]
    chartInfoArray = ChartInfo(selectedTabName)
    wkSheet = chartInfoArray(0)
    tblName = chartInfoArray(1)
    chartName = chartInfoArray(2)
    
    ' Clear user entries
    Call ResetChartParameters(wkSheet, tblName)
    
    ' Reload the UserForm fields to reflect original parameters
    Call LoadChartParameters(wkSheet, tblName)
    
    ' Update chart preview to reflect reset
    Call UpdateChartPreview(wkSheet, chartName)
    
ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error resetting chart: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' 12. ModChartCB_Click
' Purpose: Handler for "View Modified Chart"/apply changes button: validates, updates
'          table, triggers chart preview refresh.
'----------------------------------------------------------------------------------------
Private Sub ModChartCB_Click()
    On Error GoTo ErrorHandler
    
    Dim chartInfoArray As Variant
    Dim wkSheet As String
    Dim tblName As String
    Dim chartName As String
    Dim selectedTabName As String

    Call DevToolsMod.OptimizePerformance(True)

    ' Ensure a chart is selected
    selectedTabName = Me.ChartChoice.value
    If selectedTabName = "" Then
        MsgBox "Please select a chart from the list.", vbExclamation
        GoTo ExitHandler
    End If

    If Not ChartInfo.Exists(selectedTabName) Then
        MsgBox "Selected tab not found in ChartInfo mapping.", vbCritical
        GoTo ExitHandler
    End If

    ' Single lookup [Worksheet, Table, ChartName]
    chartInfoArray = ChartInfo(selectedTabName)
    wkSheet = chartInfoArray(0)
    tblName = chartInfoArray(1)
    chartName = chartInfoArray(2)

    Call SaveChartParameters(wkSheet, tblName)
    
    Call FormatChart(wkSheet, chartName, tblName)
    
    ' Update preview with modified chart
    Call UpdateChartPreview(wkSheet, chartName)

    Me.ResetChartsCB.enabled = True

ExitHandler:
    Call DevToolsMod.OptimizePerformance(False)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error modifying chart: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------------------------------------------------------
' UserForm_Terminate
' Purpose: Clean up objects when form is closed
'----------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error Resume Next
    Set ChartInfo = Nothing
End Sub
