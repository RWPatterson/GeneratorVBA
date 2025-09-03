Attribute VB_Name = "ISO16889Mod"
Option Explicit

'************************************************************************
'*************  ISO16889 Report Subs and Funcs - UPDATED  **************
'************************************************************************

Public ISO16889ReportData As ISO16889ClassMod

'This instantiates the ISO16889ReportData object, and if there is available data it will parse the data into the class module.
'This also gets called after a file is loaded in, or the forms are started again.
Public Sub SetupISO16889ClassModule()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    ' Always create a fresh instance (disposal handled in CleanupBeforeNewFile)
    Set ISO16889ReportData = New ISO16889ClassMod
    Set ISO16889ReportData.WorkbookInstance = ThisWorkbook
    
    Call DataFileMod.EnsureDataFileReady
    
    If DataFileMod.TestData.DataExist Then
        ' CRITICAL: Validate file compatibility first
        If Not ISO16889ReportData.ValidateFileCompatibility() Then
            ' Exit completely - incompatible file
            GoTo CleanExit
        End If
        
        ' STEP 1: Analyze the data file and determine actual termination values
        ' This populates the class properties AND records baseline values
        If Not InitializeDefaultTermination() Then
            MsgBox "Critical Error: Unable to determine test termination parameters." & vbCrLf & _
                   "This file may be corrupted or incomplete.", vbCritical
            GoTo CleanExit
        End If
        
        ' STEP 2: Apply any valid user overrides (optional - only if they exist and are valid)
        Call ApplyUserInterventions
        
        ' STEP 3: Generate analysis tables based on final settings (actual + any overrides)
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

Private Function InitializeDefaultTermination() As Boolean
    DevToolsMod.TimerStartCount
    
    InitializeDefaultTermination = False ' Assume failure
    
    On Error GoTo InitError
    
    'Test what kind of test is being operated on - use class method
    If Not ISO16889ReportData.EvaluateByTestType() Then
        GoTo InitError
    End If
    
    'Determine based on the test details what sensors to report from by default - use class method
    If Not ISO16889ReportData.EvaluateSelectedSensors() Then
        GoTo InitError
    End If
    
    ' Validate termination parameters were set correctly
    If ISO16889ReportData.TerminationTime <= 0 Then
        MsgBox "Unable to determine test termination time." & vbCrLf & _
               "Check that pressure data and termination setpoints are valid.", vbCritical
        GoTo InitError
    End If
    
    If ISO16889ReportData.TerminationDP <= 0 Then
        MsgBox "Unable to determine test termination pressure." & vbCrLf & _
               "Check that termination DP setpoint is specified in test headers.", vbCritical
        GoTo InitError
    End If
    
    ' DON'T UPDATE CACHE HERE - let GenerateISO16889Analysis handle it after tables are built
    ' Call ISO16889ReportData.UpdateCache  ' <-- REMOVE THIS LINE
    
    InitializeDefaultTermination = True
    DevToolsMod.TimerEndCount "Default Termination Initialization"
    Exit Function
    
InitError:
    DevToolsMod.TimerEndCount "Default Termination Initialization (FAILED)"
    InitializeDefaultTermination = False
End Function

' Main analysis generation - always builds on first run, then uses cache appropriately
Private Sub GenerateISO16889Analysis()
    DevToolsMod.TimerStartCount
    
    ' Check if rebuild is needed using class method
    If Not ISO16889ReportData.IsRebuildRequired() Then
        Debug.Print "Using cached analysis - no rebuild needed"
        DevToolsMod.TimerEndCount "ISO16889 Analysis (cached)"
        Exit Sub
    End If
    
    Debug.Print "Rebuilding ISO16889 analysis..."
    
    ' Calculate clump arrays based on (possibly modified) termination - use class methods
    Call ISO16889ReportData.SetClumpTimes
    Call ISO16889ReportData.SetClumpPressures
    
    ' Clear the 16889 Data Tab
    Debug.Print "Clearing ISO16889Data sheet..."
    Sheets("ISO16889Data").usedRange.Clear
    
    ' Generate all sensor tables
    Debug.Print "Generating ISO16889 tables..."
    Call FillISO16889Tables(DataFileMod.TestData, ISO16889ReportData)
    
    ' Calculate beta sizes after tables are built
    Call ISO16889ReportData.CalculateBetaSizes
    
    ' Record final values to the ISOSaveData table (this updates the "From Data" column)
    Debug.Print "Recording final save data..."
    Call FillISO16889SaveData
    
    ' Update class cache
    Call ISO16889ReportData.UpdateCache
    
    Debug.Print "Analysis rebuild completed"
    DevToolsMod.TimerEndCount "ISO16889 Analysis (rebuilt)"
End Sub


Private Function AnalysisTablesExist() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("ISO16889Data")
    
    On Error Resume Next
    AnalysisTablesExist = (ws.ListObjects.count > 0)
    On Error GoTo 0
    
    If AnalysisTablesExist Then
        ' Verify tables have data rows (not just headers)
        Dim tbl As ListObject
        Dim hasData As Boolean
        
        For Each tbl In ws.ListObjects
            If Not tbl.DataBodyRange Is Nothing Then
                If tbl.DataBodyRange.Rows.count > 0 Then
                    hasData = True
                    Exit For
                End If
            End If
        Next tbl
        
        AnalysisTablesExist = hasData
    End If
End Function

' NEW FUNCTION: Analyze the actual data to determine termination values
Private Function DetermineActualTerminationFromData() As Boolean
    DetermineActualTerminationFromData = False
     
    ' This calls your existing class method that analyzes the pressure data
    ' and determines which filter hit the terminal DP first
    If Not ISO16889ReportData.SetISO16889DiffPressTag() Then
        Debug.Print "ERROR: SetISO16889DiffPressTag failed"
        Exit Function
    End If
    
    ' Log what we determined from the data
    Debug.Print "Actual termination determined from data:"
    Debug.Print "  - Tag: " & ISO16889ReportData.TerminationTag
    Debug.Print "  - DP: " & ISO16889ReportData.TerminationDP
    Debug.Print "  - Time: " & ISO16889ReportData.TerminationTime
    Debug.Print "  - Filter: " & ISO16889ReportData.TerminationFilter
    
    DetermineActualTerminationFromData = True
End Function

' NEW FUNCTION: Record the baseline values in the From Data column
Private Sub RecordBaselineValues()
    ' Use the SaveDataMod to suppress change events
    Call SaveDataMod.BeginAutomatedUpdate
    
    On Error GoTo CleanupEvents
    
    ' Record the actual values determined from data analysis
    ' These go in the "From Data" column and serve as the baseline
    Call SetISO16889DataEntry(1, ISO16889ReportData.TerminationTag)        ' Termination Tag
    Call SetISO16889DataEntry(2, ISO16889ReportData.TerminationDP)         ' Termination DP
    Call SetISO16889DataEntry(3, ISO16889ReportData.TerminationTime)       ' Termination Time
    Call SetISO16889DataEntry(4, ISO16889ReportData.RetainedMassValid)     ' RetainedMassValid
    Call SetISO16889DataEntry(7, ISO16889ReportData.TerminationFilter)     ' Selected Filter
    Call SetISO16889DataEntry(8, ISO16889ReportData.TerminationSizePhrase) ' Selected Sensor
    
    
    
    Debug.Print "Baseline values recorded in From Data column"
    
CleanupEvents:
    Call SaveDataMod.EndAutomatedUpdate
End Sub

' Apply user interventions from SaveData table
Private Sub ApplyUserInterventions()
    DevToolsMod.TimerStartCount
    
    Dim userFilterChoice As String
    Dim userDPOverride As String
    Dim userSensorChoice As String
    Dim rebuildRequired As Boolean
    
    rebuildRequired = False
    
    ' Check for user filter selection override
    userFilterChoice = GetISO16889SaveResult(7)
    If IsNumeric(userFilterChoice) And userFilterChoice <> "" Then
        Dim newFilter As Integer
        newFilter = CInt(userFilterChoice)
        
        If newFilter <> ISO16889ReportData.TerminationFilter And ISO16889ReportData.IsValidFilterChoice(newFilter) Then
            Call ISO16889ReportData.ApplyFilterOverride(newFilter)
            rebuildRequired = True
        End If
    End If
    
    ' Check for user DP override (must be <= actual termination)
    userDPOverride = GetISO16889SaveResult(2) ' Assuming ID 2 is user DP override
    If IsNumeric(userDPOverride) And userDPOverride <> "" Then
        Dim newDP As Double
        newDP = CDbl(userDPOverride)
        
        If newDP <> ISO16889ReportData.TerminationDP And ISO16889ReportData.IsValidDPOverride(newDP) Then
            ISO16889ReportData.TerminationDP = newDP
            Call ISO16889ReportData.SetTerminationTime ' Recalculate termination time
            rebuildRequired = True
        End If
    End If
    
    ' Check for user sensor selection (display only - no rebuild needed)
    userSensorChoice = GetISO16889SaveResult(8)
    If userSensorChoice <> "" And userSensorChoice <> ISO16889ReportData.TerminationSizePhrase Then
        If ISO16889ReportData.IsValidSensorChoice(userSensorChoice) Then
            ISO16889ReportData.TerminationSizePhrase = userSensorChoice
        End If
    End If
    
    ' Invalidate cache if rebuild will be required
    If rebuildRequired Then
        Call ISO16889ReportData.InvalidateCache
    End If
    
    DevToolsMod.TimerEndCount "User Interventions Applied"
End Sub

' Public function to force rebuild (for dashboard button clicks)
Public Sub ForceRebuildAnalysis()
    Call ISO16889ReportData.InvalidateCache
    Call SetupISO16889ClassModule
End Sub

' Get available options for dashboard (delegate to class)
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
'================ EXCEL INTEGRATION FUNCTIONS =======================
'======================================================================

Sub FillISO16889SaveData()
    ' CRITICAL: Suppress change events during automated updates using SaveDataMod
    Call SaveDataMod.BeginAutomatedUpdate
    
    On Error GoTo CleanupEvents
    
    'Set Termination Tag
    Call SetISO16889DataEntry(1, ISO16889ReportData.TerminationTag)
    
    'Set Termination DP
    Call SetISO16889DataEntry(2, ISO16889ReportData.TerminationDP)
    
    'Set Termination Time
    Call SetISO16889DataEntry(3, ISO16889ReportData.TerminationTime)
    
    'Set RetainedMassValid
    Call SetISO16889DataEntry(4, ISO16889ReportData.RetainedMassValid)
    
    'Set SelectedFilter
    Call SetISO16889DataEntry(7, ISO16889ReportData.TerminationFilter)
    
    'Set SelectedSizePhrase
    Call SetISO16889DataEntry(8, ISO16889ReportData.TerminationSizePhrase)
    
    ' Set Beta Size Values
    Call SetISO16889DataEntry(9, ISO16889ReportData.ISO16889SizeAtBeta2)
    Call SetISO16889DataEntry(10, ISO16889ReportData.ISO16889SizeAtBeta10)
    Call SetISO16889DataEntry(11, ISO16889ReportData.ISO16889SizeAtBeta75)
    Call SetISO16889DataEntry(12, ISO16889ReportData.ISO16889SizeAtBeta100)
    Call SetISO16889DataEntry(13, ISO16889ReportData.ISO16889SizeAtBeta200)
    Call SetISO16889DataEntry(14, ISO16889ReportData.ISO16889SizeAtBeta1000)
    
CleanupEvents:
    ' CRITICAL: Re-enable change events using SaveDataMod
    Call SaveDataMod.EndAutomatedUpdate
End Sub

' Enhanced with performance optimization
Sub FillISO16889Tables(TestData As DataFileClassMod, ReportData As ISO16889ClassMod)
    DevToolsMod.TimerStartCount
    
    Dim currentRow As Long
    currentRow = ReportFillMod.GetLastRow("ISO16889Data")

    ' Process all available sensor types
    If Not IsEmpty(TestData.LBU_CountsData) Then
        Call GetISO16889TableData(TestData, ReportData, "LB")
        currentRow = CreateISO16889Tables(TestData, ReportData, "LB", currentRow)
    End If
    
    If Not IsEmpty(TestData.LSU_CountsData) Then
        Call GetISO16889TableData(TestData, ReportData, "LS")
        currentRow = CreateISO16889Tables(TestData, ReportData, "LS", currentRow)
    End If
    
    If Not IsEmpty(TestData.LBE_CountsData) Then
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
    Dim minRequiredRows As Long
    
    ' Select sizes and count data arrays based on sensor type
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
    
    ' ISO 16889 Compliant: Calculate rows per 10% time interval
    rowsPerBeta = (ReportData.TerminationTime * 0.1) / ((TestData.CountTime + TestData.HoldTime) / 60)
    
    ' Minimum required rows for valid calculation (e.g., need at least 2 data points)
    minRequiredRows = 2
    
    ' ISO 16889: Skip first 3 records per standard
    BetaStopRow = 3
    
    ' Pre-allocate arrays for better performance
    ReDim Betas(1 To 10, 1 To UBound(Sizes))
    ReDim CU(1 To 10, 1 To UBound(Sizes))
    ReDim CD(1 To 10, 1 To UBound(Sizes))
    
    ' Process each 10% time interval (10%, 20%, ... 100% of termination time)
    For i = 1 To 10
        BetaStartRow = BetaStopRow + 1
        BetaStopRow = Round(rowsPerBeta * i, 0)
        
        ' Ensure we don't exceed available data
        If BetaStopRow > UBound(TargetArrUp, 1) Then
            BetaStopRow = UBound(TargetArrUp, 1)
        End If
        
        ' CRITICAL FIX: Check if we have sufficient data for meaningful calculation
        Dim dataWindowSize As Long
        dataWindowSize = BetaStopRow - BetaStartRow + 1
        
        If BetaStopRow <= BetaStartRow Or dataWindowSize < minRequiredRows Then
            ' FIXED: Leave values blank (Empty) instead of setting to 0
            For j = 1 To UBound(Sizes)
                CU(i, j) = Empty      ' Was: CU(i, j) = 0
                CD(i, j) = Empty      ' Was: CD(i, j) = 0
                Betas(i, j) = Empty   ' Was: Betas(i, j) = 100000
            Next j
            
            Debug.Print "Interval " & i & " (time " & Format(ReportData.TerminationTime * i / 10, "0.0") & " min): " & _
                       "Insufficient data (window size: " & dataWindowSize & "). Values left blank."
            GoTo NextInterval
        End If
        
        ' Log data window information for debugging
        Debug.Print "Interval " & i & " (time " & Format(ReportData.TerminationTime * i / 10, "0.0") & " min): " & _
                   "Rows " & BetaStartRow & " to " & BetaStopRow & " (window size: " & dataWindowSize & ")"
        
        ' Process each particle size bin
        For j = 1 To UBound(Sizes)
            Dim windowSize As Long
            windowSize = BetaStopRow - BetaStartRow + 1
            
            ' Pre-allocate temp arrays for this window
            ReDim TempArrUp(1 To windowSize)
            ReDim TempArrDn(1 To windowSize)
            
            ' Copy data for this time window
            For k = 1 To windowSize
                Dim dataRow As Long
                dataRow = BetaStartRow + k - 1
                
                If dataRow <= UBound(TargetArrUp, 1) Then
                    TempArrUp(k) = TargetArrUp(dataRow, j)
                    TempArrDn(k) = TargetArrDn(dataRow, j)
                Else
                    ' Use last available data point
                    TempArrUp(k) = TargetArrUp(UBound(TargetArrUp, 1), j)
                    TempArrDn(k) = TargetArrDn(UBound(TargetArrDn, 1), j)
                End If
            Next k
            
            ' Calculate averages for this time interval
            CU(i, j) = Application.WorksheetFunction.Average(TempArrUp)
            CD(i, j) = Application.WorksheetFunction.Average(TempArrDn)
            
            ' Calculate beta ratio with proper handling of zero downstream counts
            If CD(i, j) > 0 Then
                Betas(i, j) = CU(i, j) / CD(i, j)
            Else
                Betas(i, j) = 100000 ' ISO 16889: Maximum beta when no downstream counts
            End If
            
            ' Cap maximum beta value per ISO 16889
            If IsNumeric(Betas(i, j)) And Betas(i, j) > 100000 Then
                Betas(i, j) = 100000
            End If
        Next j
        
NextInterval:
    Next i
    
    ' Store results in ReportData properties
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
    
    DevToolsMod.TimerEndCount "ISO16889 " & sensorType & " Data Processing"
End Sub

Function CreateISO16889Tables(TestData As DataFileClassMod, ReportData As ISO16889ClassMod, sensorType As String, StartingRow As Long) As Long
    Dim ws As Worksheet
    Set ws = Sheets("ISO16889Data")
    
    Dim Sizes As Variant
    Dim CU As Variant, CD As Variant, Beta As Variant
    Dim LabelPrefix As String
    Dim colCount As Long
    Dim arrTimes() As Variant, arrPressures() As Variant
    Dim i As Long, j As Long, rowCount As Long
    
    ' Select data arrays and label prefix for sensor type
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
    
    ' ENHANCED: Clean arrays to handle Empty values properly before Excel write
    CU = TableMod.CleanArrayForExcel(CU)
    CD = TableMod.CleanArrayForExcel(CD)
    Beta = TableMod.CleanArrayForExcel(Beta)
    
    ' Pre-convert 1D arrays C_Times and C_Pressures to 2D variant arrays for bulk write
    rowCount = UBound(ReportData.C_Times)
    ReDim arrTimes(1 To rowCount, 1 To 1)
    ReDim arrPressures(1 To rowCount, 1 To 1)
    For i = 1 To rowCount
        arrTimes(i, 1) = ReportData.C_Times(i)
        arrPressures(i, 1) = ReportData.C_Pressures(i)
    Next i
    
    ' Sizes array 1D to 2D for row write
    colCount = UBound(Sizes)
    Dim arrSizes() As Variant
    ReDim arrSizes(1 To 1, 1 To colCount)
    For i = 1 To colCount
        arrSizes(1, i) = Sizes(i)
    Next i
    
    ' Disable updates during dump
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' -- CU Table --
    With ws.Range("A" & StartingRow)
        .Value = LabelPrefix & "U Average Counts"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ws.Range("A" & StartingRow + 2).Value = "Elapsed Time"
    ws.Range("B" & StartingRow + 2).Value = "Pressure"
    
    ws.Range("A" & StartingRow + 3).Resize(rowCount, 1).Value = arrTimes
    ws.Range("B" & StartingRow + 3).Resize(rowCount, 1).Value = arrPressures
    ws.Range("C" & StartingRow + 2).Resize(1, colCount).Value = arrSizes
    ws.Range("C" & StartingRow + 3).Resize(UBound(CU, 1), UBound(CU, 2)).Value = CU
    
    Call TableMod.CreateTable("ISO16889Data", "A" & StartingRow + 2, "CU_" & LabelPrefix, True)
    
    StartingRow = StartingRow + 16
    
    ' -- CD Table --
    With ws.Range("A" & StartingRow)
        .Value = LabelPrefix & "D Average Counts"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ws.Range("A" & StartingRow + 2).Value = "Elapsed Time"
    ws.Range("B" & StartingRow + 2).Value = "Pressure"
    
    ws.Range("A" & StartingRow + 3).Resize(rowCount, 1).Value = arrTimes
    ws.Range("B" & StartingRow + 3).Resize(rowCount, 1).Value = arrPressures
    ws.Range("C" & StartingRow + 2).Resize(1, colCount).Value = arrSizes
    ws.Range("C" & StartingRow + 3).Resize(UBound(CD, 1), UBound(CD, 2)).Value = CD
    
    Call TableMod.CreateTable("ISO16889Data", "A" & StartingRow + 2, "CD_" & LabelPrefix, True)
    
    StartingRow = StartingRow + 16
    
    ' -- Beta Table --
    With ws.Range("A" & StartingRow)
        .Value = LabelPrefix & " Beta Ratios"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ws.Range("A" & StartingRow + 2).Value = "Elapsed Time"
    ws.Range("B" & StartingRow + 2).Value = "Pressure"
    
    ws.Range("A" & StartingRow + 3).Resize(rowCount, 1).Value = arrTimes
    ws.Range("B" & StartingRow + 3).Resize(rowCount, 1).Value = arrPressures
    ws.Range("C" & StartingRow + 2).Resize(1, colCount).Value = arrSizes
    ws.Range("C" & StartingRow + 3).Resize(UBound(Beta, 1), UBound(Beta, 2)).Value = Beta
    
    Call TableMod.CreateTable("ISO16889Data", "A" & StartingRow + 2, "Beta_" & LabelPrefix, True)
    
    ' Restore application settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Return next free row after tables
    CreateISO16889Tables = StartingRow + 16
End Function

'======================================================================
'============ WORKSHEET/EXCEL SPECIFIC FUNCTIONS ====================
'======================================================================

'Get ISO 16889 Pressure
'Usage: =GetISO16889Time(TerminalDP,0.7)
Function GetISO16889Time(DP As Double, Percentage As Double) As Variant
    GetISO16889Time = ISO16889ReportData.TerminationDP * Percentage
End Function

'Get ISO 16889 Pressure
'Usage: =GetISO16889Pressure(TerminalDP,0.7)
Function GetISO16889Pressure(DP As Double, Percentage As Double) As Double
    Dim DPData As Variant
    Dim Times As Variant
    Dim t As Variant
    Dim firstTime As Variant
    Dim timePt As Double
    Dim i As Integer
    
    Times = DataFileMod.TestData.Times
    firstTime = Times(1)
    
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
    
    InjGrav = GetISO16889SaveResult(5) / 1000 'grams per ml
    Times = DataFileMod.TestData.Times
    InjFlowAve = GetSaveResult(46) / 1000
    
    If IsEmpty(ISO16889ReportData.InjectedMass) Then
        ReDim massData(UBound(Times))
                
        For i = 1 To UBound(Times)
            massData(i) = Times(i) * 1440 * InjFlowAve * InjGrav 'min * ml/min * gram/ml = grams
        Next
    Else
        Set massData = ISO16889ReportData.InjectedMass
    End If
    
    timePt = Times(UBound(Times)) * Percentage
    
    GetISO16889Mass = MathMod.LinInterpolation(timePt, Times, massData, 1)
End Function

'Returns an array of the selected termination
Function GetISO16889ElementDP(wkSheet As String) As Variant
    Dim CleanHousingDP As Double
    Dim DPressTag As String
    Dim DPressArry As Variant
    Dim ElementDP As Variant
    Dim i As Integer
    
    DPressTag = ISO16889ReportData.TerminationTag
    
    'if the differential pressure tag is set to FinalDPress then the second element in the CleanHousingDP
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


'======================================================================
'================ SAVEDATA TABLE FUNCTIONS ============================
'======================================================================

'This sub returns the value of a field from the Save_Data table.
Public Function GetISO16889SaveResult(ID As Integer) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    
    GetISO16889SaveResult = tbl.DataBodyRange(ID, 3).Value
    Exit Function

ErrorHandler:
    GetISO16889SaveResult = "ID Not Found"
End Function

Public Sub SetISO16889SaveUserEntry(ID As Integer, SaveValue As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    
    tbl.DataBodyRange(ID, 4).Value = SaveValue
    Exit Sub

ErrorHandler:
    ' Silent error handling
End Sub

Public Sub SetISO16889DefaultEntry(ID As Integer, SaveValue As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    
    tbl.DataBodyRange(ID, 5).Value = SaveValue
    Exit Sub

ErrorHandler:
    ' Silent error handling
End Sub

Public Sub SetISO16889DataEntry(ID As Integer, SaveValue As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    
    Debug.Print "Setting ID " & ID & " to value: " & SaveValue
    tbl.DataBodyRange(ID, 6).Value = SaveValue  ' Column 6 = "From Data"
    
    Exit Sub

ErrorHandler:
    Debug.Print "Error in SetISO16889DataEntry ID " & ID & ": " & Err.Description
End Sub

'************************************************************************
'****************  Save Data Table Management  *************************
'************************************************************************

Private Sub ClearISO16889Data()
    On Error GoTo ErrorHandler
    
    ' Clear the ISO16889Data sheet completely
    If Not IsEmpty(Sheets("ISO16889Data").Range("A1")) Then
        Sheets("ISO16889Data").usedRange.Clear
        Debug.Print "Cleared ISO16889Data sheet"
    End If
    
    ' Clear only specific table columns with maximum safety
    Call ClearChartTableUserEntries
    
    Debug.Print "Conservative clear completed"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ClearISO16889Data_Conservative: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub ClearChartTableUserEntries()
    Dim sheetTablePairs As Variant
    
    ' Define sheet/table pairs explicitly to avoid any confusion
    sheetTablePairs = Array( _
        Array("C1_DP_v_Mass", "ISO16889C1SITable"), Array("C2_Beta_v_Size", "ISO16889C2Table"), Array("C3_Beta_v_Time", "ISO16889C3Table"), Array("C4_Beta_v_Press", "ISO16889C4SITable"), Array("C5_Up_Counts", "ISO16889C5Table"), Array("C6_Down_Counts", "ISO16889C6Table"))
    
    Dim i As Long
    For i = LBound(sheetTablePairs) To UBound(sheetTablePairs)
        Call ClearSingleTableUserEntry((sheetTablePairs(i)(0)), (sheetTablePairs(i)(1)))
    Next i
End Sub

Private Sub ClearSingleTableUserEntry(sheetName As String, tableName As String)
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
    
    If tbl.ListColumns.count < 3 Then
        Debug.Print "Table " & tableName & " has insufficient columns"
        Exit Sub
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        Debug.Print "Table " & tableName & " has no data body"
        Exit Sub
    End If
    
    ' Finally clear the user entry column (column 3)
    tbl.ListColumns(3).DataBodyRange.ClearContents
    Debug.Print "Cleared user entries from " & tableName
    
    On Error GoTo 0
End Sub

Private Function IsISO16889ModuleDataRow(rowID As Long) As Boolean
    ' Only clear rows that ISO16889 modules directly populate
    Select Case rowID
        Case 1, 2, 3, 4, 7, 8, 9, 10, 11, 12, 13, 14
            IsISO16889ModuleDataRow = True
        Case Else
            IsISO16889ModuleDataRow = False  ' Formula-calculated or other data
    End Select
End Function



'======================================================================
'================ CHART SETUP FUNCTIONS ==============================
'======================================================================

Sub Format16889DataTables(ByVal wsName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn

    Set ws = ThisWorkbook.Worksheets(wsName)

    For Each tbl In ws.ListObjects
        tbl.ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
        
        For Each col In tbl.ListColumns
            If col.index > 1 Then
                col.DataBodyRange.NumberFormat = "0.00"
            End If
        Next col
    Next tbl
End Sub

Sub SetISO16889C1DPvMassSI()
    Call FormatChart("C1_DP_v_Mass", "ISO16889Chart1SI", "ISO16889C1SITable")
    Call SetChartSeriesByRange("C1_DP_v_Mass", "ISO16889Chart1SI", "V3")
End Sub

Sub SetISO16889C2SizevBetaSI()
    Call FormatChart("C2_Beta_v_Size", "ISO16889C2Chart", "ISO16889C2Table")
    Call SetChartSeriesByRange("C2_Beta_v_Size", "ISO16889C2Chart", "V3")
End Sub

Sub SetISO16889C3TimevBeta()
    Call FormatChart("C3_Beta_v_Time", "ISO16889C3Chart", "ISO16889C3Table")
    Call SetChartSeriesByRange("C3_Beta_v_Time", "ISO16889C3Chart", "V3")
End Sub

Sub SetISO16889C4PressureSIvBeta()
    Call FormatChart("C4_Beta_v_Press", "ISO16889C4Chart", "ISO16889C4SITable")
    Call SetChartSeriesByRange("C4_Beta_v_Press", "ISO16889C4Chart", "V3")
End Sub

Sub SetISO16889C5UpCountsVsTime()
    Call FormatChart("C5_Up_Counts", "ISO16889C5UpCountsVsTime", "ISO16889C5Table")
    Call SetChartSeriesByRange("C5_Up_Counts", "ISO16889C5UpCountsVsTime", "V3")
End Sub

Sub SetISO16889C6DnCountsVsTime()
    Call FormatChart("C6_Down_Counts", "ISO16889C6DnCountsVsTimeChart", "ISO16889C6Table")
    Call SetChartSeriesByRange("C6_Down_Counts", "ISO16889C6DnCountsVsTimeChart", "V3")
End Sub

'======================================================================
'================ CLEANUP AND DISPOSAL FUNCTIONS ====================
'======================================================================

' Call this function before loading any new data file
Public Sub CleanupBeforeNewFile()
    DevToolsMod.TimerStartCount
    
    ' 1. Properly dispose class modules
    Call DisposeISO16889ClassModule
    Call DisposeDataFileClassModule
    
    ' 2. Selectively clear From Data column entries (preserve formulas)
    Call ClearFromDataEntries
    
    ' 3. Clear data sheets (existing functionality)
    Call TableMod.DeleteDataTables("A1")
    
    ' 4. Clear ISO 16889 specific data
    Call ClearISO16889Data
    
    DevToolsMod.TimerEndCount "Complete Cleanup"
End Sub

' Properly dispose of ISO16889 class module
Private Sub DisposeISO16889ClassModule()
    On Error Resume Next
    
    If Not ISO16889ReportData Is Nothing Then
        ' Clear any cached data first
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
        
        ' Clear object references
        Set ISO16889ReportData.WorkbookInstance = Nothing
        Set ISO16889ReportData = Nothing
    End If
    
    On Error GoTo 0
End Sub

' Dispose of DataFile class module
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
        
        ' Clear object reference
        Set DataFileMod.TestData.WorkbookInstance = Nothing
        Set DataFileMod.TestData = Nothing
    End If
    
    On Error GoTo 0
End Sub

' Clear From Data column entries but preserve formulas
Private Sub ClearFromDataEntries()
    On Error Resume Next
    
    ' Clear SaveDataTable "From Data" entries (column 6)
    Call ClearDirectWritesInColumn("SaveDataTable", 6)
    
    ' Clear ISO16889SaveDataTable "From Data" entries (column 6)
    Call ClearDirectWritesInColumn("ISO16889SaveDataTable", 6)
    
    On Error GoTo 0
End Sub

' Helper function to clear only direct writes, preserve formulas
Private Sub ClearDirectWritesInColumn(tableName As String, columnIndex As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim cellValue As Variant
    Dim hasFormula As Boolean
    
    On Error Resume Next
    Set ws = Sheets("Save_Data")
    Set tbl = ws.ListObjects(tableName)
    
    If tbl Is Nothing Then Exit Sub
    
    ' Suppress change events during cleanup using SaveDataMod
    Call SaveDataMod.BeginAutomatedUpdate
    
    For i = 1 To tbl.DataBodyRange.Rows.count
        ' Check if cell has a formula
        hasFormula = (Left(tbl.DataBodyRange(i, columnIndex).Formula, 1) = "=")
        
        ' Only clear cells that don't have formulas (direct writes from code)
        If Not hasFormula Then
            cellValue = tbl.DataBodyRange(i, columnIndex).Value
            
            ' Only clear if there's actually a value (not already empty)
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                tbl.DataBodyRange(i, columnIndex).ClearContents
            End If
        End If
    Next i
    
    ' Re-enable change events using SaveDataMod
    Call SaveDataMod.EndAutomatedUpdate
    On Error GoTo 0
End Sub



' Helper function to check if worksheet exists
Private Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(wsName).Name = wsName)
    On Error GoTo 0
End Function

' Debug function to verify cleanup worked
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
            If Not IsEmpty(tbl.DataBodyRange(i, columnIndex).Value) Then
                count = count + 1
            End If
        End If
    Next i
    
    CountNonFormulaEntries = count
    On Error GoTo 0
End Function
