Attribute VB_Name = "ISO16889Mod"
Option Explicit

'************************************************************************
'*************  ISO16889 Report Subs and Funcs - REFACTORED  ***********
'************************************************************************

' CENTRALIZED CONFIGURATION CONSTANTS
Private Const CHART_C1_SHEET As String = "C1_DP_v_Mass"
Private Const CHART_C1_TABLE As String = "ISO16889C1SITable"
Private Const CHART_C1_CHART As String = "ISO16889Chart1SI"
Private Const CHART_C2_SHEET As String = "C2_Beta_v_Size"
Private Const CHART_C2_TABLE As String = "ISO16889C2Table"
Private Const CHART_C2_CHART As String = "ISO16889C2Chart"
Private Const CHART_C3_SHEET As String = "C3_Beta_v_Time"
Private Const CHART_C3_TABLE As String = "ISO16889C3Table"
Private Const CHART_C3_CHART As String = "ISO16889C3Chart"
Private Const CHART_C4_SHEET As String = "C4_Beta_v_Press"
Private Const CHART_C4_TABLE As String = "ISO16889C4SITable"
Private Const CHART_C4_CHART As String = "ISO16889C4Chart"
Private Const CHART_C5_SHEET As String = "C5_Up_Counts"
Private Const CHART_C5_TABLE As String = "ISO16889C5Table"
Private Const CHART_C5_CHART As String = "ISO16889C5UpCountsVsTime"
Private Const CHART_C6_SHEET As String = "C6_Down_Counts"
Private Const CHART_C6_TABLE As String = "ISO16889C6Table"
Private Const CHART_C6_CHART As String = "ISO16889C6DnCountsVsTimeChart"

Private Const DEFAULT_COUNT_TIME As Long = 60
Private Const DEFAULT_HOLD_TIME As Long = 0
Private Const MAX_BETA_VALUE As Double = 100000
Private Const MIN_DATA_WINDOW As Long = 2
Private Const ISO16889_SKIP_ROWS As Long = 3

Public ISO16889ReportData As ISO16889ClassMod

'======================================================================
'================== MAIN PUBLIC INTERFACE ===========================
'======================================================================

Public Sub SetupISO16889ClassModule()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    ' Always create a fresh instance (disposal handled in CleanupBeforeNewFile)
    Set ISO16889ReportData = New ISO16889ClassMod
    Set ISO16889ReportData.WorkbookInstance = ThisWorkbook

    If DataFileMod.EnsureTestDataReady() Then
        ' STEP 1: Initialize and validate
        If Not InitializeAndValidateData() Then GoTo CleanExit
        
        ' STEP 2: Apply user overrides and generate analysis
        Call ApplyUserInterventions
        Call GenerateISO16889Analysis
        
    Else
        MsgBox "No valid test data found. Please load a valid .DAT file first.", vbExclamation
    End If
    
CleanExit:
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "ISO16889 Setup Complete"
    If Err.Number <> 0 Then
        MsgBox "Fatal Error in ISO16889 setup: " & Err.Description, vbCritical
    End If
End Sub

Public Sub ForceRebuildAnalysis()
    Call ISO16889ReportData.InvalidateCache
    Call SetupISO16889ClassModule
End Sub

'======================================================================
'================== CONSOLIDATED FUNCTIONS ==========================
'======================================================================

' CONSOLIDATED: Replaces 4 separate SaveData functions
Public Function ManageISO16889SaveData(ID As Integer, operation As String, Optional value As Variant) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim targetColumn As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    
    Select Case UCase(operation)
        Case "GET"
            ManageISO16889SaveData = tbl.DataBodyRange(ID, 3).value ' Report Value column
        Case "SET_USER"
            targetColumn = 4 ' User Entry column
        Case "SET_DEFAULT"
            targetColumn = 5 ' Custom Default column
        Case "SET_DATA"
            targetColumn = 6 ' From Data column
            Debug.Print "Setting ID " & ID & " to value: " & value
        Case Else
            Err.Raise vbObjectError + 1001, "ManageISO16889SaveData", "Invalid operation: " & operation
    End Select
    
    If operation <> "GET" Then
        tbl.DataBodyRange(ID, targetColumn).value = value
    End If
    
    Exit Function
    
ErrorHandler:
    If operation = "GET" Then
        ManageISO16889SaveData = "ID Not Found"
    Else
        Debug.Print "Error in ManageISO16889SaveData ID " & ID & ": " & Err.Description
    End If
End Function

' CONSOLIDATED: Replaces 6 nearly identical chart setup functions
Private Sub SetupISO16889Chart(chartID As String)
    Dim config As Variant
    config = GetChartConfig(chartID)
    
    If Not IsArray(config) Then
        Debug.Print "Unknown chart ID: " & chartID
        Exit Sub
    End If
    
    ' FIX: Use explicit variables instead of array elements directly in ByRef calls
    Dim sheetName As String
    Dim chartName As String
    Dim tableName As String
    
    sheetName = config(0)
    tableName = config(1)
    chartName = config(2)
    
    Call Charts.FormatChart(sheetName, chartName, tableName)
    Call Charts.SetChartSeriesByRange(sheetName, chartName, "V3")
End Sub

' CONSOLIDATED: Chart configuration lookup
Private Function GetChartConfig(chartID As String) As Variant
    Select Case UCase(chartID)
        Case "C1", "C1_DP_V_MASS"
            GetChartConfig = Array(CHART_C1_SHEET, CHART_C1_TABLE, CHART_C1_CHART)
        Case "C2", "C2_BETA_V_SIZE"
            GetChartConfig = Array(CHART_C2_SHEET, CHART_C2_TABLE, CHART_C2_CHART)
        Case "C3", "C3_BETA_V_TIME"
            GetChartConfig = Array(CHART_C3_SHEET, CHART_C3_TABLE, CHART_C3_CHART)
        Case "C4", "C4_BETA_V_PRESS"
            GetChartConfig = Array(CHART_C4_SHEET, CHART_C4_TABLE, CHART_C4_CHART)
        Case "C5", "C5_UP_COUNTS"
            GetChartConfig = Array(CHART_C5_SHEET, CHART_C5_TABLE, CHART_C5_CHART)
        Case "C6", "C6_DOWN_COUNTS"
            GetChartConfig = Array(CHART_C6_SHEET, CHART_C6_TABLE, CHART_C6_CHART)
        Case Else
            GetChartConfig = Empty
    End Select
End Function

' CONSOLIDATED: Replaces 3 similar validation functions
Private Function ValidateAndApplyUserOverride(overrideType As String, newValue As Variant) As Boolean
    ValidateAndApplyUserOverride = False
    
    ' FIX: Handle empty/zero values properly
    If IsEmpty(newValue) Or newValue = "" Then
        Debug.Print overrideType & " override is empty - skipping validation"
        Exit Function
    End If
    
    Select Case UCase(overrideType)
        Case "DP"
            If Not IsNumeric(newValue) Then
                Debug.Print "DP override is not numeric: " & newValue
                Exit Function
            End If
            
            Dim dpValue As Double
            dpValue = CDbl(newValue)
            
            ' FIX: Handle zero values gracefully
            If dpValue <= 0 Then
                Debug.Print "DP override is zero or negative - skipping validation"
                Exit Function
            End If
            
            If ISO16889ReportData.IsValidDPOverride(dpValue) Then
                Call ISO16889ReportData.InvalidateCache
                Call PromptForISO16889Rebuild("DP override")
                ValidateAndApplyUserOverride = True
            Else
                ShowValidationError "DP", dpValue, ISO16889ReportData.GetActualTerminationDP()
            End If
            
        Case "FILTER"
            If Not IsNumeric(newValue) Then
                Debug.Print "Filter override is not numeric: " & newValue
                Exit Function
            End If
            
            Dim filterValue As Integer
            filterValue = CInt(newValue)
            
            ' FIX: Handle zero filter values gracefully
            If filterValue <= 0 Then
                Debug.Print "Filter override is zero or negative - skipping validation"
                Exit Function
            End If
            
            If ISO16889ReportData.IsValidFilterChoice(filterValue) Then
                Call ISO16889ReportData.InvalidateCache
                Call PromptForISO16889Rebuild("filter selection")
                ValidateAndApplyUserOverride = True
            Else
                ShowValidationError "FILTER", filterValue, ISO16889ReportData.GetAvailableFilterOptions()
            End If
            
        Case "SENSOR"
            Dim sensorValue As String
            sensorValue = CStr(newValue)
            
            If sensorValue = "" Then
                Debug.Print "Sensor override is empty - skipping validation"
                Exit Function
            End If
            
            If ISO16889ReportData.IsValidSensorChoice(sensorValue) Then
                Call PromptForSensorChange(sensorValue)
                ValidateAndApplyUserOverride = True
            Else
                ShowValidationError "SENSOR", sensorValue, ISO16889ReportData.GetAvailableSensorOptions()
            End If
    End Select
End Function

'======================================================================
'================== PRIVATE IMPLEMENTATION FUNCTIONS ================
'======================================================================

Private Function InitializeAndValidateData() As Boolean
    DevToolsMod.TimerStartCount
    InitializeAndValidateData = False
    
    On Error GoTo InitError
    
    ' Validate file compatibility first
    If Not ISO16889ReportData.ValidateFileCompatibility() Then GoTo InitError
    
    ' Initialize termination parameters
    If Not ISO16889ReportData.EvaluateByTestType() Then GoTo InitError
    If Not ISO16889ReportData.EvaluateSelectedSensors() Then GoTo InitError
    
    ' Validate termination parameters were set correctly
    If ISO16889ReportData.TerminationTime <= 0 Then
        MsgBox "Unable to determine test termination time.", vbCritical
        GoTo InitError
    End If
    
    If ISO16889ReportData.TerminationDP <= 0 Then
        MsgBox "Unable to determine test termination pressure.", vbCritical
        GoTo InitError
    End If
    
    InitializeAndValidateData = True
    DevToolsMod.TimerEndCount "Data Validation"
    Exit Function
    
InitError:
    DevToolsMod.TimerEndCount "Data Validation (FAILED)"
    InitializeAndValidateData = False
End Function

Private Sub ApplyUserInterventions()
    DevToolsMod.TimerStartCount
    
    Dim userChoice As String
    Dim rebuildRequired As Boolean
    
    ' Check filter override
    userChoice = ManageISO16889SaveData(7, "GET")
    If IsNumeric(userChoice) And userChoice <> "" Then
        If CInt(userChoice) <> ISO16889ReportData.TerminationFilter Then
            If ValidateAndApplyUserOverride("FILTER", userChoice) Then
                Call ISO16889ReportData.ApplyFilterOverride(CInt(userChoice))
                rebuildRequired = True
            End If
        End If
    End If
    
    ' Check DP override
    userChoice = ManageISO16889SaveData(2, "GET")
    If IsNumeric(userChoice) And userChoice <> "" Then
        If CDbl(userChoice) <> ISO16889ReportData.TerminationDP Then
            If ValidateAndApplyUserOverride("DP", userChoice) Then
                ISO16889ReportData.TerminationDP = CDbl(userChoice)
                Call ISO16889ReportData.SetTerminationTime
                rebuildRequired = True
            End If
        End If
    End If
    
    ' Check sensor override (display only)
    userChoice = ManageISO16889SaveData(8, "GET")
    If userChoice <> "" And userChoice <> ISO16889ReportData.TerminationSizePhrase Then
        If ValidateAndApplyUserOverride("SENSOR", userChoice) Then
            ISO16889ReportData.TerminationSizePhrase = userChoice
        End If
    End If
    
    If rebuildRequired Then Call ISO16889ReportData.InvalidateCache
    
    DevToolsMod.TimerEndCount "User Interventions Applied"
End Sub

Private Sub GenerateISO16889Analysis()
    DevToolsMod.TimerStartCount
    
    ' CRITICAL FIX: Check if rebuild is actually needed
    If Not ISO16889ReportData.IsRebuildRequired() Then
        Debug.Print "Using cached analysis - no rebuild needed"
        DevToolsMod.TimerEndCount "ISO16889 Analysis (cached)"
        Exit Sub
    End If
    
    Debug.Print "Rebuilding ISO16889 analysis..."
    
    ' Calculate clump arrays
    Call ISO16889ReportData.SetClumpTimes
    Call ISO16889ReportData.SetClumpPressures
    
    ' Clear and rebuild tables
    Sheets("ISO16889Data").UsedRange.Clear
    Call FillISO16889Tables(DataFileMod.TestData, ISO16889ReportData)
    
    ' CRITICAL FIX: Only calculate beta sizes ONCE after tables are built
    Call ISO16889ReportData.CalculateBetaSizes
    
    ' Record final values
    Call FillISO16889SaveData
    
    ' Update cache to prevent repeated calculations
    Call ISO16889ReportData.UpdateCache
    
    Debug.Print "Analysis rebuild completed"
    DevToolsMod.TimerEndCount "ISO16889 Analysis (rebuilt)"
End Sub

Private Sub ShowValidationError(errorType As String, value As Variant, availableOptions As Variant)
    Select Case UCase(errorType)
        Case "DP"
            MsgBox "Invalid DP override: " & value & vbCrLf & _
                   "Maximum allowed: " & availableOptions, vbExclamation
        Case "FILTER"
            MsgBox "Invalid filter selection: " & value & vbCrLf & _
                   "Available filters: " & availableOptions, vbExclamation
        Case "SENSOR"
            MsgBox "Invalid sensor selection: " & value & vbCrLf & _
                   "Available sensors: " & availableOptions, vbExclamation
    End Select
End Sub

Private Sub PromptForISO16889Rebuild(changeType As String)
    Dim result As VbMsgBoxResult
    result = MsgBox("Your " & changeType & " change requires rebuilding ISO 16889 analysis. Rebuild now?", _
                   vbYesNo + vbQuestion, "Rebuild Required")
    If result = vbYes Then Call ForceRebuildAnalysis
End Sub

Private Sub PromptForSensorChange(newSensor As String)
    Dim result As VbMsgBoxResult
    result = MsgBox("Change particle counter sensor to " & newSensor & "?", _
                   vbYesNo + vbQuestion, "Sensor Change")
    If result = vbYes Then
        MsgBox "Sensor selection updated. Report displays will now use " & newSensor & " data.", vbInformation
    End If
End Sub

'======================================================================
'================== EXCEL INTEGRATION FUNCTIONS =======================
'======================================================================

Sub FillISO16889SaveData()
    Call SaveDataMod.BeginAutomatedUpdate
    
    On Error GoTo CleanupEvents
    
    ' Core termination data
    Call ManageISO16889SaveData(1, "SET_DATA", ISO16889ReportData.TerminationTag)
    Call ManageISO16889SaveData(2, "SET_DATA", ISO16889ReportData.TerminationDP)
    Call ManageISO16889SaveData(3, "SET_DATA", ISO16889ReportData.TerminationTime)
    Call ManageISO16889SaveData(4, "SET_DATA", ISO16889ReportData.RetainedMassValid)
    Call ManageISO16889SaveData(7, "SET_DATA", ISO16889ReportData.TerminationFilter)
    Call ManageISO16889SaveData(8, "SET_DATA", ISO16889ReportData.TerminationSizePhrase)
    
    ' Beta size calculations
    Call ManageISO16889SaveData(9, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta2)
    Call ManageISO16889SaveData(10, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta10)
    Call ManageISO16889SaveData(11, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta75)
    Call ManageISO16889SaveData(12, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta100)
    Call ManageISO16889SaveData(13, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta200)
    Call ManageISO16889SaveData(14, "SET_DATA", ISO16889ReportData.ISO16889SizeAtBeta1000)
    
CleanupEvents:
    Call SaveDataMod.EndAutomatedUpdate
End Sub

Sub FillISO16889Tables(TestData As DataFileClassMod, ReportData As ISO16889ClassMod)
    DevToolsMod.TimerStartCount
    
    Dim currentRow As Long
    currentRow = ReportFillMod.GetLastRow("ISO16889Data")

    ' Process all available sensor types
    If ReportFillMod.hasData(TestData.LBU_CountsData) Then
        Call GetISO16889TableData(TestData, ReportData, "LB")
        currentRow = CreateISO16889Tables(TestData, ReportData, "LB", currentRow)
    End If
    
    If ReportFillMod.hasData(TestData.LSU_CountsData) Then
        Call GetISO16889TableData(TestData, ReportData, "LS")
        currentRow = CreateISO16889Tables(TestData, ReportData, "LS", currentRow)
    End If
    
    If ReportFillMod.hasData(TestData.LBE_CountsData) Then
        Call GetISO16889TableData(TestData, ReportData, "LBE")
        currentRow = CreateISO16889Tables(TestData, ReportData, "LBE", currentRow)
    End If
    
    DevToolsMod.TimerEndCount "ISO16889 Tables Generation"
End Sub

Sub GetISO16889TableData(TestData As DataFileClassMod, ReportData As ISO16889ClassMod, sensorType As String)
    DevToolsMod.TimerStartCount
    
    Dim rowsPerBeta As Double
    Dim BetaStartRow As Long, BetaStopRow As Long
    Dim i As Long, j As Long, k As Long
    Dim Sizes As Variant
    Dim TempArrUp() As Double, TempArrDn() As Double
    Dim CU() As Variant, CD() As Variant, Betas() As Variant
    Dim TargetArrUp As Variant, TargetArrDn As Variant
    
    ' Get sensor-specific data arrays
    Call GetSensorArrays(TestData, sensorType, Sizes, TargetArrUp, TargetArrDn)
    
    ' ISO 16889 Compliant: Calculate rows per 10% time interval
    rowsPerBeta = (ReportData.TerminationTime * 0.1) / ((TestData.CountTime + TestData.HoldTime) / 60)
    
    ' Pre-allocate arrays
    ReDim Betas(1 To 10, 1 To UBound(Sizes))
    ReDim CU(1 To 10, 1 To UBound(Sizes))
    ReDim CD(1 To 10, 1 To UBound(Sizes))
    
    BetaStopRow = ISO16889_SKIP_ROWS
    
    ' Process each 10% time interval
    For i = 1 To 10
        BetaStartRow = BetaStopRow + 1
        BetaStopRow = Round(rowsPerBeta * i, 0)
        
        If BetaStopRow > UBound(TargetArrUp, 1) Then BetaStopRow = UBound(TargetArrUp, 1)
        
        Dim dataWindowSize As Long
        dataWindowSize = BetaStopRow - BetaStartRow + 1
        
        If BetaStopRow <= BetaStartRow Or dataWindowSize < MIN_DATA_WINDOW Then
            ' Leave values blank for insufficient data
            For j = 1 To UBound(Sizes)
                CU(i, j) = Empty
                CD(i, j) = Empty
                Betas(i, j) = Empty
            Next j
            GoTo NextInterval
        End If
        
        ' Calculate averages for each particle size
        For j = 1 To UBound(Sizes)
            Dim windowSize As Long
            windowSize = BetaStopRow - BetaStartRow + 1
            
            ReDim TempArrUp(1 To windowSize)
            ReDim TempArrDn(1 To windowSize)
            
            ' Copy data window
            For k = 1 To windowSize
                Dim dataRow As Long
                dataRow = BetaStartRow + k - 1
                
                If dataRow <= UBound(TargetArrUp, 1) Then
                    TempArrUp(k) = TargetArrUp(dataRow, j)
                    TempArrDn(k) = TargetArrDn(dataRow, j)
                Else
                    TempArrUp(k) = TargetArrUp(UBound(TargetArrUp, 1), j)
                    TempArrDn(k) = TargetArrDn(UBound(TargetArrDn, 1), j)
                End If
            Next k
            
            ' Calculate averages and beta ratios
            CU(i, j) = Application.WorksheetFunction.Average(TempArrUp)
            CD(i, j) = Application.WorksheetFunction.Average(TempArrDn)
            
            If CD(i, j) > 0 Then
                Betas(i, j) = CU(i, j) / CD(i, j)
                If Betas(i, j) > MAX_BETA_VALUE Then Betas(i, j) = MAX_BETA_VALUE
            Else
                Betas(i, j) = MAX_BETA_VALUE
            End If
        Next j
        
NextInterval:
    Next i
    
    ' Store results in ReportData
    Call StoreSensorResults(ReportData, sensorType, CU, CD, Betas)
    
    DevToolsMod.TimerEndCount "ISO16889 " & sensorType & " Data Processing"
End Sub

Private Sub GetSensorArrays(TestData As DataFileClassMod, sensorType As String, ByRef Sizes As Variant, ByRef TargetArrUp As Variant, ByRef TargetArrDn As Variant)
    Select Case sensorType
        Case "LS"
            Sizes = TestData.LS_Sizes
            TargetArrUp = TestData.LSU_CountsData
            TargetArrDn = TestData.LSD_CountsData
        Case "LBE"
            Sizes = TestData.LBE_Sizes
            TargetArrUp = TestData.LBD_CountsData
            TargetArrDn = TestData.LBE_CountsData
        Case Else ' Default to LB
            Sizes = TestData.LB_Sizes
            TargetArrUp = TestData.LBU_CountsData
            TargetArrDn = TestData.LBD_CountsData
    End Select
End Sub

Private Sub StoreSensorResults(ReportData As ISO16889ClassMod, sensorType As String, CU As Variant, CD As Variant, Betas As Variant)
    Select Case sensorType
        Case "LS"
            ReportData.CU_LS = CU
            ReportData.CD_LS = CD
            ReportData.Beta_LS = Betas
        Case "LBE"
            ReportData.CU_LBE = CU
            ReportData.CD_LBE = CD
            ReportData.Beta_LBE = Betas
        Case Else ' LB default
            ReportData.CU_LB = CU
            ReportData.CD_LB = CD
            ReportData.Beta_LB = Betas
    End Select
End Sub

Function CreateISO16889Tables(TestData As DataFileClassMod, ReportData As ISO16889ClassMod, sensorType As String, StartingRow As Long) As Long
    Dim ws As Worksheet
    Set ws = Sheets("ISO16889Data")
    
    Dim Sizes As Variant
    Dim CU As Variant, CD As Variant, Beta As Variant
    Dim LabelPrefix As String
    
    ' Get sensor-specific data and clean arrays
    Call GetSensorDataForTable(TestData, ReportData, sensorType, Sizes, CU, CD, Beta, LabelPrefix)
    
    ' Create the three tables
    StartingRow = CreateSingleTable(ws, StartingRow, LabelPrefix & "U Average Counts", "CU_" & LabelPrefix, ReportData, Sizes, CU)
    StartingRow = CreateSingleTable(ws, StartingRow, LabelPrefix & "D Average Counts", "CD_" & LabelPrefix, ReportData, Sizes, CD)
    StartingRow = CreateSingleTable(ws, StartingRow, LabelPrefix & " Beta Ratios", "Beta_" & LabelPrefix, ReportData, Sizes, Beta)
    
    CreateISO16889Tables = StartingRow
End Function

Private Sub GetSensorDataForTable(TestData As DataFileClassMod, ReportData As ISO16889ClassMod, sensorType As String, ByRef Sizes As Variant, ByRef CU As Variant, ByRef CD As Variant, ByRef Beta As Variant, ByRef LabelPrefix As String)
    Select Case sensorType
        Case "LS"
            Sizes = TestData.LS_Sizes
            CU = ReportData.CU_LS
            CD = ReportData.CD_LS
            Beta = ReportData.Beta_LS
            LabelPrefix = "LS"
        Case "LBE"
            Sizes = TestData.LBE_Sizes
            CU = ReportData.CU_LBE
            CD = ReportData.CD_LBE
            Beta = ReportData.Beta_LBE
            LabelPrefix = "LBE"
        Case Else ' LB default
            Sizes = TestData.LB_Sizes
            CU = ReportData.CU_LB
            CD = ReportData.CD_LB
            Beta = ReportData.Beta_LB
            LabelPrefix = "LB"
    End Select
    
    ' FIX: Use TableMod.CleanArrayForExcel if it exists, otherwise handle Empty values here
    On Error Resume Next
    CU = TableMod.CleanArrayForExcel(CU)
    CD = TableMod.CleanArrayForExcel(CD)
    Beta = TableMod.CleanArrayForExcel(Beta)
    
    ' If TableMod.CleanArrayForExcel doesn't exist, handle it locally
    If Err.Number <> 0 Then
        On Error GoTo 0
        CU = HandleEmptyValues(CU)
        CD = HandleEmptyValues(CD)
        Beta = HandleEmptyValues(Beta)
    End If
    On Error GoTo 0
End Sub

Private Function CleanArrayForExcel(arr As Variant) As Variant
    ' Replace Empty values with "" for Excel compatibility
    If IsEmpty(arr) Then
        CleanArrayForExcel = arr
        Exit Function
    End If
    
    Dim i As Long, j As Long
    Dim cleanedArray As Variant
    cleanedArray = arr
    
    For i = LBound(cleanedArray, 1) To UBound(cleanedArray, 1)
        For j = LBound(cleanedArray, 2) To UBound(cleanedArray, 2)
            If IsEmpty(cleanedArray(i, j)) Then
                cleanedArray(i, j) = ""
            End If
        Next j
    Next i
    
    CleanArrayForExcel = cleanedArray
End Function

Private Function HandleEmptyValues(arr As Variant) As Variant
    If IsEmpty(arr) Then
        HandleEmptyValues = arr
        Exit Function
    End If
    
    Dim i As Long, j As Long
    Dim cleanedArray As Variant
    cleanedArray = arr
    
    On Error Resume Next
    For i = LBound(cleanedArray, 1) To UBound(cleanedArray, 1)
        For j = LBound(cleanedArray, 2) To UBound(cleanedArray, 2)
            If IsEmpty(cleanedArray(i, j)) Then
                cleanedArray(i, j) = ""
            End If
        Next j
    Next i
    On Error GoTo 0
    
    HandleEmptyValues = cleanedArray
End Function

Private Function CreateSingleTable(ws As Worksheet, StartingRow As Long, tableTitle As String, tableName As String, ReportData As ISO16889ClassMod, Sizes As Variant, dataArray As Variant) As Long
    Dim i As Long
    Dim rowCount As Long, colCount As Long
    
    ' Disable updates during table creation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Set title
    With ws.Range("A" & StartingRow)
        .value = tableTitle
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ' Set headers
    ws.Range("A" & StartingRow + 2).value = "Elapsed Time"
    ws.Range("B" & StartingRow + 2).value = "Pressure"
    
    ' Write time and pressure data
    rowCount = UBound(ReportData.C_Times)
    For i = 1 To rowCount
        ws.Cells(StartingRow + 2 + i, 1).value = ReportData.C_Times(i)
        ws.Cells(StartingRow + 2 + i, 2).value = ReportData.C_Pressures(i)
    Next i
    
    ' Write size headers
    colCount = UBound(Sizes)
    For i = 1 To colCount
        ws.Cells(StartingRow + 2, 2 + i).value = Sizes(i)
    Next i
    
    ' Write data array
    ws.Range("C" & StartingRow + 3).Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray
    
    ' Create table
    Call TableMod.CreateTable("ISO16889Data", "A" & StartingRow + 2, tableName, True)
    
    ' Restore application settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    CreateSingleTable = StartingRow + 16
End Function

Sub Format16889DataTables(ByVal wsName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn

    Set ws = ThisWorkbook.Worksheets(wsName)

    For Each tbl In ws.ListObjects
        If tbl.ListColumns.count > 0 Then
            tbl.ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
            
            For Each col In tbl.ListColumns
                If col.index > 1 Then
                    col.DataBodyRange.NumberFormat = "0.00"
                End If
            Next col
        End If
    Next tbl
End Sub

'======================================================================
'================== LEGACY PUBLIC INTERFACE ==========================
'======================================================================

' These maintain backward compatibility - delegate to consolidated functions
Public Function GetISO16889SaveResult(ID As Integer) As String
    GetISO16889SaveResult = ManageISO16889SaveData(ID, "GET")
End Function

Public Sub SetISO16889SaveUserEntry(ID As Integer, SaveValue As String)
    Call ManageISO16889SaveData(ID, "SET_USER", SaveValue)
End Sub

Public Sub SetISO16889DefaultEntry(ID As Integer, SaveValue As String)
    Call ManageISO16889SaveData(ID, "SET_DEFAULT", SaveValue)
End Sub

Public Sub SetISO16889DataEntry(ID As Integer, SaveValue As String)
    Call ManageISO16889SaveData(ID, "SET_DATA", SaveValue)
End Sub

' Chart setup functions - delegate to consolidated function
Public Sub SetISO16889C1DPvMassSI()
    Call SetupISO16889Chart("C1")
End Sub

Public Sub SetISO16889C2SizevBetaSI()
    Call SetupISO16889Chart("C2")
End Sub

Public Sub SetISO16889C3TimevBeta()
    Call SetupISO16889Chart("C3")
End Sub

Public Sub SetISO16889C4PressureSIvBeta()
    Call SetupISO16889Chart("C4")
End Sub

Public Sub SetISO16889C5UpCountsVsTime()
    Call SetupISO16889Chart("C5")
End Sub

Public Sub SetISO16889C6DnCountsVsTime()
    Call SetupISO16889Chart("C6")
End Sub

' Get available options (delegate to class)
Public Function GetAvailableFilterOptions() As String
    If ISO16889ReportData Is Nothing Then
        GetAvailableFilterOptions = "1"
    Else
        GetAvailableFilterOptions = ISO16889ReportData.GetAvailableFilterOptions()
    End If
End Function

Public Function GetAvailableSensorOptions() As String
    If ISO16889ReportData Is Nothing Then
        GetAvailableSensorOptions = ""
    Else
        GetAvailableSensorOptions = ISO16889ReportData.GetAvailableSensorOptions()
    End If
End Function

'======================================================================
'================== WORKSHEET FUNCTIONS (UNCHANGED) ==================
'======================================================================

Function GetISO16889Time(DP As Double, Percentage As Double) As Variant
    GetISO16889Time = ISO16889ReportData.TerminationDP * Percentage
End Function

Function GetISO16889Pressure(DP As Double, Percentage As Double) As Double
    Dim DPData As Variant
    Dim Times As Variant
    Dim timePt As Double
    
    Times = DataFileMod.TestData.Times
    DPData = DataFileMod.TestData.GetAnalogTagData(ISO16889ReportData.TerminationTag)
    timePt = Times(UBound(Times)) * Percentage
    
    GetISO16889Pressure = MathMod.LinInterpolation(timePt, Times, DPData, 1)
End Function

Function GetISO16889Mass(DP As Double, Percentage As Double) As Double
    Dim massData As Variant
    Dim Times As Variant
    Dim timePt As Double
    Dim i As Integer
    Dim InjGrav As Double
    Dim InjFlowAve As Double
    
    InjGrav = ManageISO16889SaveData(5, "GET") / 1000
    Times = DataFileMod.TestData.Times
    InjFlowAve = GetSaveResult(46) / 1000
    
    If IsEmpty(ISO16889ReportData.InjectedMass) Then
        ReDim massData(UBound(Times))
        For i = 1 To UBound(Times)
            massData(i) = Times(i) * 1440 * InjFlowAve * InjGrav
        Next
    Else
        Set massData = ISO16889ReportData.InjectedMass
    End If
    
    timePt = Times(UBound(Times)) * Percentage
    GetISO16889Mass = MathMod.LinInterpolation(timePt, Times, massData, 1)
End Function

Function GetISO16889ElementDP(wkSheet As String) As Variant
    Dim CleanHousingDP As Double
    Dim DPressTag As String
    Dim DPressArry As Variant
    Dim ElementDP As Variant
    Dim i As Integer
    
    DPressTag = ISO16889ReportData.TerminationTag
    
    If DPressTag = "TS_FinalDPress" Then
        CleanHousingDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "CleanHousingDP", 2)
    Else
        CleanHousingDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "CleanHousingDP", 1)
    End If
    
    DPressArry = TableMod.GetArrayFromTable("AnalogData", "AnalogDataTbl", DPressTag)
    
    ReDim ElementDP(UBound(DPressArry))
    For i = 0 To UBound(DPressArry)
        ElementDP(i) = DPressArry(i) - CleanHousingDP
    Next
    
    GetISO16889ElementDP = ElementDP
End Function

'Returns the beta from the selected sensors according to the named range calculation.
Function GetISO16889SizeGivenBeta(Beta As Double) As String
    Dim tempY As Variant
    Dim tempX As Variant
    Dim reshapeArrayY() As Double
    Dim reshapeArrayX() As Double
    Dim i As Integer
    Dim n As Integer
    
    'Get Sizes
    tempY = Application.Transpose(Evaluate("Selected_Sensor_Sizes"))
    'Get Average Betas
    tempX = Evaluate("Selected16889BetasAverages")

    ' Reshape the 2D arrays into 1D 0 indexed arrays to use LinInterpolation function.
    n = UBound(tempY)
    ReDim reshapeArrayX(0 To n - 1)
    ReDim reshapeArrayY(0 To n - 1)

    For i = 0 To n - 1
        reshapeArrayX(i) = tempX(1, i + 3)
        reshapeArrayY(i) = tempY(i + 1)
    Next i

    If reshapeArrayX(0) > Beta Then
        GetISO16889SizeGivenBeta = "<" & tempY(1)
    Else
        GetISO16889SizeGivenBeta = Format(MathMod.LinInterpolation(Beta, reshapeArrayX, reshapeArrayY, 0), "#0.0")
    End If
End Function

'======================================================================
'================== CLEANUP FUNCTIONS (SIMPLIFIED) ==================
'======================================================================

Public Sub CleanupBeforeNewFile()
    DevToolsMod.TimerStartCount
    
    Call DisposeISO16889ClassModule
    Call DisposeDataFileClassModule
    Call ClearFromDataEntries
    Call TableMod.DeleteDataTables("A1")
    Call ClearISO16889Data
    
    DevToolsMod.TimerEndCount "Complete Cleanup"
End Sub

Private Sub DisposeISO16889ClassModule()
    On Error Resume Next
    If Not ISO16889ReportData Is Nothing Then
        Call ISO16889ReportData.InvalidateCache
        
        ' Clear arrays to free memory
        If Not IsEmpty(ISO16889ReportData.C_Times) Then
            Erase ISO16889ReportData.C_Times
        End If
        If Not IsEmpty(ISO16889ReportData.C_Pressures) Then
            Erase ISO16889ReportData.C_Pressures
        End If
        If Not IsEmpty(ISO16889ReportData.C_Masses) Then
            Erase ISO16889ReportData.C_Masses
        End If
        If Not IsEmpty(ISO16889ReportData.CU_LB) Then
            Erase ISO16889ReportData.CU_LB
        End If
        If Not IsEmpty(ISO16889ReportData.CD_LB) Then
            Erase ISO16889ReportData.CD_LB
        End If
        If Not IsEmpty(ISO16889ReportData.Beta_LB) Then
            Erase ISO16889ReportData.Beta_LB
        End If
        If Not IsEmpty(ISO16889ReportData.CU_LS) Then
            Erase ISO16889ReportData.CU_LS
        End If
        If Not IsEmpty(ISO16889ReportData.CD_LS) Then
            Erase ISO16889ReportData.CD_LS
        End If
        If Not IsEmpty(ISO16889ReportData.Beta_LS) Then
            Erase ISO16889ReportData.Beta_LS
        End If
        If Not IsEmpty(ISO16889ReportData.CU_LBE) Then
            Erase ISO16889ReportData.CU_LBE
        End If
        If Not IsEmpty(ISO16889ReportData.CD_LBE) Then
            Erase ISO16889ReportData.CD_LBE
        End If
        If Not IsEmpty(ISO16889ReportData.Beta_LBE) Then
            Erase ISO16889ReportData.Beta_LBE
        End If
        
        Set ISO16889ReportData.WorkbookInstance = Nothing
        Set ISO16889ReportData = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub DisposeDataFileClassModule()
    On Error Resume Next
    If Not DataFileMod.TestData Is Nothing Then
        ' Clear large data arrays to free memory
        If Not IsEmpty(DataFileMod.TestData.analogData) Then
            Erase DataFileMod.TestData.analogData
        End If
        If Not IsEmpty(DataFileMod.TestData.LBU_CountsData) Then
            Erase DataFileMod.TestData.LBU_CountsData
        End If
        If Not IsEmpty(DataFileMod.TestData.LBD_CountsData) Then
            Erase DataFileMod.TestData.LBD_CountsData
        End If
        If Not IsEmpty(DataFileMod.TestData.LSU_CountsData) Then
            Erase DataFileMod.TestData.LSU_CountsData
        End If
        If Not IsEmpty(DataFileMod.TestData.LSD_CountsData) Then
            Erase DataFileMod.TestData.LSD_CountsData
        End If
        If Not IsEmpty(DataFileMod.TestData.LBE_CountsData) Then
            Erase DataFileMod.TestData.LBE_CountsData
        End If
        If Not IsEmpty(DataFileMod.TestData.cycleAnalogData) Then
            Erase DataFileMod.TestData.cycleAnalogData
        End If
        If Not IsEmpty(DataFileMod.TestData.HeaderData) Then
            Erase DataFileMod.TestData.HeaderData
        End If
        
        Set DataFileMod.TestData.WorkbookInstance = Nothing
        Set DataFileMod.TestData = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub ClearFromDataEntries()
    On Error Resume Next
    Call ClearDirectWritesInColumn("SaveDataTable", 6)
    Call ClearDirectWritesInColumn("ISO16889SaveDataTable", 6)
    On Error GoTo 0
End Sub

Private Sub ClearDirectWritesInColumn(tableName As String, columnIndex As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    On Error Resume Next
    Set ws = Sheets("Save_Data")
    Set tbl = ws.ListObjects(tableName)
    If tbl Is Nothing Then Exit Sub
    
    Call SaveDataMod.BeginAutomatedUpdate
    
    For i = 1 To tbl.DataBodyRange.Rows.count
        If Left(tbl.DataBodyRange(i, columnIndex).Formula, 1) <> "=" Then
            If Not IsEmpty(tbl.DataBodyRange(i, columnIndex).value) Then
                tbl.DataBodyRange(i, columnIndex).ClearContents
            End If
        End If
    Next i
    
    Call SaveDataMod.EndAutomatedUpdate
    On Error GoTo 0
End Sub

Private Sub ClearISO16889Data()
    On Error Resume Next
    
    ' ONLY clear the ISO16889Data sheet (analysis tables)
    If Not IsEmpty(Sheets("ISO16889Data").Range("A1")) Then
        Sheets("ISO16889Data").UsedRange.Clear
        Debug.Print "Cleared ISO16889Data analysis sheet"
    End If
    
    ' FIX: ONLY clear user entry columns in chart tables, NOT the chart data
    Call ClearChartTableUserEntriesOnly
    
    Debug.Print "Conservative clear completed - chart data preserved"
    On Error GoTo 0
End Sub

Private Sub ClearChartTableUserEntriesOnly()
    Dim chartSheets As Variant
    Dim tableNames As Variant
    Dim i As Long
    
    ' Define chart sheets and their corresponding table names
    chartSheets = Array(CHART_C1_SHEET, CHART_C2_SHEET, CHART_C3_SHEET, CHART_C4_SHEET, CHART_C5_SHEET, CHART_C6_SHEET)
    tableNames = Array(CHART_C1_TABLE, CHART_C2_TABLE, CHART_C3_TABLE, CHART_C4_TABLE, CHART_C5_TABLE, CHART_C6_TABLE)
    
    For i = LBound(chartSheets) To UBound(chartSheets)
        Call ClearSingleChartTableUserEntry(CStr(chartSheets(i)), CStr(tableNames(i)))
    Next i
End Sub

Private Sub ClearSingleChartTableUserEntry(sheetName As String, tableName As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Multiple safety checks
    Set ws = Sheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Sheet not found: " & sheetName
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(tableName)
    If tbl Is Nothing Then
        Debug.Print "Table not found: " & tableName & " on sheet " & sheetName
        Exit Sub
    End If
    
    ' Check if table has enough columns (need at least 3 for user entry column)
    If tbl.ListColumns.count < 3 Then
        Debug.Print "Table " & tableName & " has insufficient columns (" & tbl.ListColumns.count & ")"
        Exit Sub
    End If
    
    ' Check if table has data body
    If tbl.DataBodyRange Is Nothing Then
        Debug.Print "Table " & tableName & " has no data body"
        Exit Sub
    End If
    
    ' Get the user entry column (assuming it's column 3 based on your table structure)
    Dim userEntryCol As ListColumn
    Set userEntryCol = Nothing
    
    ' Try to find "User Entry" column by name first
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If InStr(1, col.Name, "User", vbTextCompare) > 0 And InStr(1, col.Name, "Entry", vbTextCompare) > 0 Then
            Set userEntryCol = col
            Exit For
        End If
    Next col
    
    ' If not found by name, use column 3 as fallback
    If userEntryCol Is Nothing And tbl.ListColumns.count >= 3 Then
        Set userEntryCol = tbl.ListColumns(3)
    End If
    
    ' Clear only the user entry column
    If Not userEntryCol Is Nothing Then
        userEntryCol.DataBodyRange.ClearContents
        Debug.Print "Cleared user entries from " & tableName & " (column: " & userEntryCol.Name & ")"
    Else
        Debug.Print "Could not find user entry column in " & tableName
    End If
    
    On Error GoTo 0
End Sub

Private Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(wsName).Name = wsName)
    On Error GoTo 0
End Function

'======================================================================
'================== DIAGNOSTIC FUNCTIONS (FOR DEVTOOLS) =============
'======================================================================

Public Sub VerifyCleanup()
    Debug.Print "=== CLEANUP VERIFICATION ==="
    Debug.Print "ISO16889ReportData Is Nothing: " & (ISO16889ReportData Is Nothing)
    Debug.Print "TestData Is Nothing: " & (DataFileMod.TestData Is Nothing)
    
    On Error Resume Next
    Dim count1 As Long, count2 As Long
    count1 = CountNonFormulaEntries("SaveDataTable", 6)
    count2 = CountNonFormulaEntries("ISO16889SaveDataTable", 6)
    Debug.Print "SaveDataTable non-formula entries: " & count1
    Debug.Print "ISO16889SaveDataTable non-formula entries: " & count2
    On Error GoTo 0
    
    Debug.Print "=== END VERIFICATION ==="
End Sub

Private Function CountNonFormulaEntries(tableName As String, columnIndex As Long) As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim count As Long
    
    On Error Resume Next
    Set ws = Sheets("Save_Data")
    Set tbl = ws.ListObjects(tableName)
    
    If tbl Is Nothing Then
        CountNonFormulaEntries = 0
        Exit Function
    End If
    
    count = 0
    For i = 1 To tbl.DataBodyRange.Rows.count
        If Left(tbl.DataBodyRange(i, columnIndex).Formula, 1) <> "=" Then
            If Not IsEmpty(tbl.DataBodyRange(i, columnIndex).value) Then
                count = count + 1
            End If
        End If
    Next i
    
    CountNonFormulaEntries = count
    On Error GoTo 0
End Function

