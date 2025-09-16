Attribute VB_Name = "Main"
Option Explicit

'*************************************************************************************
'******* MAIN CONTROLLER - Performance Optimized for Report Generation ***************
'*************************************************************************************

' Key Performance Principles:
' 1. Validate once, cache results
' 2. Only rebuild analysis when parameters actually change
' 3. Batch operations where possible
' 4. Minimize Excel recalculations

Private Type SystemState
    DataFileLoaded As Boolean
    AnalysisBuilt As Boolean
    chartsInitialized As Boolean
    LastAnalysisHash As String  ' Cache hash to detect changes
    IsProcessing As Boolean     ' Prevent recursive calls
End Type

Private analysisHash As String
Private chartsInitialized As Boolean
Private sysState As SystemState

'======================================================================
'============== MAIN ENTRY POINTS ===================================
'======================================================================

' Primary initialization - called from Workbook_Open
Public Sub GenerateReport()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    Debug.Print "=== GenerateReport Started ==="
    
    On Error GoTo CleanExit
    
    ' Prevent recursive calls
    If sysState.IsProcessing Then
        Debug.Print "GenerateReport already in progress, exiting"
        GoTo CleanExit
    End If
    sysState.IsProcessing = True
    
    ' STEP 1: Quick validation - is there raw data waiting?
    If HasUnprocessedRawData() Then
        Debug.Print "Found unprocessed raw data, processing..."
        Call ProcessDataFileOnce
        sysState.DataFileLoaded = True
    End If
    
    ' STEP 2: Ensure TestData object exists (lightweight check)
    If Not DataFileMod.EnsureTestDataReady() Then
        Debug.Print "No valid test data available"
        Call UpdateDashboard  ' Show empty state
        GoTo CleanExit
    End If
    
    ' STEP 3: Only build analysis if we have processed data and it's needed
    If sysState.DataFileLoaded And DataFileMod.TestData.DataExist Then
        Call BuildAnalysisIfNeeded
    End If
    
    ' STEP 4: Update dashboard (lightweight operation)
    Call UpdateDashboard
    
    Debug.Print "=== GenerateReport Completed ==="
    
CleanExit:
    sysState.IsProcessing = False
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Main.GenerateReport"
End Sub

' Lightweight function to load new data file
Public Function LoadNewDataFile() As Boolean
    DevToolsMod.TimerStartCount
    
    Debug.Print "=== LoadNewDataFile Started ==="
    LoadNewDataFile = False
    
    On Error GoTo LoadError
    
    ' Clean up previous data first
    Call CleanupBeforeNewFile
    
    ' Open new file using existing function
    If File_Subs.OpenDataFile() Then
        ' Process the file immediately
        Call ProcessDataFileOnce
        
        ' Mark system state
        sysState.DataFileLoaded = True
        sysState.AnalysisBuilt = False
        sysState.chartsInitialized = False
        
        LoadNewDataFile = True
        Debug.Print "New data file loaded successfully"
    Else
        Debug.Print "Failed to open new data file"
    End If
    
    ' Always update dashboard to reflect current state
    Call UpdateDashboard
    
    DevToolsMod.TimerEndCount "Main.LoadNewDataFile"
    Exit Function
    
LoadError:
    Debug.Print "Error in LoadNewDataFile: " & Err.Description
    LoadNewDataFile = False
    DevToolsMod.TimerEndCount "Main.LoadNewDataFile (Error)"
End Function

'======================================================================
'============== DASHBOARD MANAGEMENT =================================
'======================================================================

' Lightweight dashboard update - no heavy processing
Public Sub UpdateDashboard()
    DevToolsMod.TimerStartCount
    
    On Error Resume Next  ' Dashboard updates should never crash the system
    
    ' Quick state assessment without expensive validation
    Dim hasData As Boolean
    Dim hasAnalysis As Boolean
    
    hasData = (Not DataFileMod.TestData Is Nothing) And DataFileMod.TestData.DataExist
    hasAnalysis = hasData And (Not ISO16889Mod.ISO16889ReportData Is Nothing)
    
    ' Update dashboard display elements
    Call UpdateDashboardDisplay(hasData, hasAnalysis)
    
    ' Only trigger heavy operations if explicitly needed
    If hasData And Not sysState.AnalysisBuilt Then
        ' User has data but no analysis - offer to build it
        Call SetDashboardMessage("Data loaded. Click 'Build Analysis' to generate report.")
    ElseIf Not hasData Then
        Call SetDashboardMessage("No data loaded. Click 'Load File' to begin.")
    Else
        Call SetDashboardMessage("Ready")
    End If
    
    DevToolsMod.TimerEndCount "Main.UpdateDashboard"
End Sub

' Update dashboard visual elements without triggering calculations
Private Sub UpdateDashboardDisplay(hasData As Boolean, hasAnalysis As Boolean)
    Dim ws As Worksheet
    Dim pc As String
    Dim filterSel As Variant
    Dim unitsSel As String
    Dim pressurePhrase As String

    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' === Buttons that require data ===
    HandleButton ws, "BtnModifyGravs", hasData, "EditGravimetrics"
    HandleButton ws, "BtnModifyGraphs", hasData, "ShowChartForm"
    HandleButton ws, "BtnPrintReport", hasData, "PrintSelectedSheets"
    
    ' === Always visible buttons ===
    HandleButton ws, "BtnCreateReport", True, "CreateReport"
    HandleButton ws, "BtnModifyLogo", True, "ModifyLogoMacro"

    ' === Macro changes & file name display ===
    If hasData Then
        HandleButton ws, "BtnModifyTestInfo", True, "MacroModifyTestInfo_Normal"
        HandleButton ws, "BtnSaveReport", True, "SaveAsReport"
        ws.Shapes("BtnSaveReport").TextFrame.Characters.Text = "Save Report"
        ws.Shapes("BoxFileName").TextFrame.Characters.Text = "File Name: " & Range("RD_FileName").Value
    Else
        HandleButton ws, "BtnModifyTestInfo", True, "MacroModifyTestInfo_CustomDefaults"
        HandleButton ws, "BtnSaveReport", True, "SaveAsTemplate"
        ws.Shapes("BtnSaveReport").TextFrame.Characters.Text = "Save Template"
        ws.Shapes("BoxFileName").TextFrame.Characters.Text = "File Name: "
    End If

    ' === TOGGLE BUTTONS ===
    pc = ""
    If hasData Then pc = ISO16889Mod.GetISO16889SaveResult(8)
    
    With ws.Shapes("BtnToggleParticleCounter")
        If hasData And pc <> "" Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .OnAction = "ToggleParticleCounter"
            .TextFrame.Characters.Text = "Counter: " & pc
        ElseIf hasData And pc = "" Then
            .Fill.ForeColor.RGB = RGB(217, 217, 217)
            .OnAction = ""
            .TextFrame.Characters.Text = "Single Set"
        Else
            .Fill.ForeColor.RGB = RGB(191, 191, 191)
            .OnAction = ""
            .TextFrame.Characters.Text = "Counter: --"
        End If
    End With

    ' Filter Pressure Toggle
    filterSel = ""
    If hasData Then filterSel = ISO16889Mod.GetISO16889SaveResult(7)
    pressurePhrase = CStr(filterSel)
    
    With ws.Shapes("BtnToggleFilterPressure")
        If hasData And pressurePhrase <> "TS_DPress" Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .OnAction = "ToggleFilterPressure"
            .TextFrame.Characters.Text = "Filter: " & filterSel
        ElseIf hasData And pressurePhrase = "TS_DPress" Then
            .Fill.ForeColor.RGB = RGB(217, 217, 217)
            .OnAction = ""
            .TextFrame.Characters.Text = "Filter 1 only"
        Else
            .Fill.ForeColor.RGB = RGB(191, 191, 191)
            .OnAction = ""
            .TextFrame.Characters.Text = "Filter: --"
        End If
    End With

    ' Report Units Toggle
    unitsSel = ""
    If hasData Then unitsSel = ReportFillMod.GetSaveResult(30)
    
    With ws.Shapes("BtnToggleReportUnits")
        If hasData Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .OnAction = "ToggleReportUnits"
            .TextFrame.Characters.Text = "Units: " & unitsSel
        Else
            .Fill.ForeColor.RGB = RGB(191, 191, 191)
            .OnAction = ""
            .TextFrame.Characters.Text = "Units: --"
        End If
    End With
End Sub

Private Sub SetDashboardMessage(message As String)
    On Error Resume Next
    ' Update a status label or cell on the dashboard
    Range("DashboardStatus").Value = message
End Sub

Private Sub EnableDashboardButton(buttonName As String, enabled As Boolean)
    On Error Resume Next
    ' Enable/disable buttons on dashboard
    ' Implementation depends on how buttons are created (shapes, ActiveX, etc.)
End Sub

'======================================================================
'============== INTELLIGENT PROCESSING ==============================
'======================================================================

' Check for unprocessed raw data without expensive validation
Private Function HasUnprocessedRawData() As Boolean
    HasUnprocessedRawData = False
    
    On Error Resume Next
    
    ' Quick check - is there a HEADER in RawData but no processed data arrays?
    If Sheets("RawData").Cells(1, 1).Value = "HEADER" Then
        ' Raw data exists, check if TestData object has been populated
        If DataFileMod.TestData Is Nothing Then
            HasUnprocessedRawData = True
        ElseIf Not DataFileMod.TestData.DataExist Then
            HasUnprocessedRawData = True
        ElseIf DataFileMod.TestData.FileName = "" Then
            ' Object exists but is empty
            HasUnprocessedRawData = True
        End If
    End If
    
    Debug.Print "HasUnprocessedRawData: " & HasUnprocessedRawData
End Function

' Process data file exactly once - no redundant validation
Private Sub ProcessDataFileOnce()
    DevToolsMod.TimerStartCount
    
    Debug.Print "ProcessDataFileOnce: Starting data processing"
    
    ' Process data using existing optimized function
    Call DataFileMod.ProcessDataFile
    
    ' Verify processing succeeded
    If DataFileMod.TestData.DataExist Then
        Debug.Print "Data processing completed successfully"
        Debug.Print "  File: " & DataFileMod.TestData.FileName
        Debug.Print "  Type: " & DataFileMod.TestData.testType
        Debug.Print "  Rows: " & DataFileMod.TestData.DataRowCount
    Else
        Debug.Print "Data processing failed"
    End If
    
    DevToolsMod.TimerEndCount "ProcessDataFileOnce"
End Sub

' Build analysis only if needed and cache the result
Private Sub BuildAnalysisIfNeeded()
    DevToolsMod.TimerStartCount
    
    ' Check if analysis is already built and current
    If sysState.AnalysisBuilt And AnalysisIsCurrent() Then
        Debug.Print "Analysis is current, skipping rebuild"
        DevToolsMod.TimerEndCount "BuildAnalysisIfNeeded (cached)"
        Exit Sub
    End If
    
    Debug.Print "Building ISO 16889 analysis..."
    
    ' Build analysis using existing function
    Call ISO16889Mod.SetupISO16889ClassModule
    
    ' Update system state
    sysState.AnalysisBuilt = True
    sysState.LastAnalysisHash = GenerateAnalysisHash()
    
    ' Initialize charts after analysis is built
    Call InitializeChartsIfNeeded
    
    DevToolsMod.TimerEndCount "BuildAnalysisIfNeeded (rebuilt)"
End Sub

' Check if current analysis matches the data state
Private Function AnalysisIsCurrent() As Boolean
    AnalysisIsCurrent = False
    
    ' Quick checks first
    If ISO16889Mod.ISO16889ReportData Is Nothing Then Exit Function
    If sysState.LastAnalysisHash = "" Then Exit Function
    
    ' Compare current state hash with cached hash
    Dim currentHash As String
    currentHash = GenerateAnalysisHash()
    
    AnalysisIsCurrent = (currentHash = sysState.LastAnalysisHash)
    
    Debug.Print "AnalysisIsCurrent: " & AnalysisIsCurrent
    If Not AnalysisIsCurrent Then
        Debug.Print "  Cached hash: " & sysState.LastAnalysisHash
        Debug.Print "  Current hash: " & currentHash
    End If
End Function

' Generate a simple hash of key analysis parameters
Private Function GenerateAnalysisHash() As String
    On Error Resume Next
    
    Dim hash As String
    
    ' Include key parameters that would trigger analysis rebuild
    If Not DataFileMod.TestData Is Nothing Then
        hash = hash & "|File:" & DataFileMod.TestData.FileName
        hash = hash & "|Rows:" & DataFileMod.TestData.DataRowCount
    End If
    
    ' Include key ISO 16889 parameters
    hash = hash & "|Filter:" & ISO16889Mod.GetISO16889SaveResult(7)
    hash = hash & "|Sensor:" & ISO16889Mod.GetISO16889SaveResult(8)
    hash = hash & "|DP:" & ISO16889Mod.GetISO16889SaveResult(2)
    
    GenerateAnalysisHash = hash
End Function

'======================================================================
'============== CHART MANAGEMENT ====================================
'======================================================================

' Initialize charts only once after analysis is built
Private Sub InitializeChartsIfNeeded()
    If sysState.chartsInitialized Then
        Debug.Print "Charts already initialized"
        Exit Sub
    End If
    
    DevToolsMod.TimerStartCount
    
    ' Batch update all charts at once
    Call UpdateAllCharts
    
    sysState.chartsInitialized = True
    DevToolsMod.TimerEndCount "InitializeChartsIfNeeded"
End Sub

' Update all charts in batch for better performance
Public Sub UpdateAllCharts()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    Debug.Print "Updating all charts..."
    
    On Error Resume Next
    
    ' Update all 6 charts in sequence
    Call ISO16889Mod.SetISO16889C1DPvMassSI
    Call ISO16889Mod.SetISO16889C2SizevBetaSI
    Call ISO16889Mod.SetISO16889C3TimevBeta
    Call ISO16889Mod.SetISO16889C4PressureSIvBeta
    Call ISO16889Mod.SetISO16889C5UpCountsVsTime
    Call ISO16889Mod.SetISO16889C6DnCountsVsTime
    
    Debug.Print "All charts updated"
    
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "UpdateAllCharts"
End Sub

' Force chart refresh (for manual button clicks)
Public Sub ForceChartRefresh()
    sysState.chartsInitialized = False
    Call InitializeChartsIfNeeded
End Sub

'======================================================================
'============== CLEANUP MANAGEMENT ==================================
'======================================================================

' Clean up before loading new file
Private Sub CleanupBeforeNewFile()
    DevToolsMod.TimerStartCount
    
    Debug.Print "Cleaning up before new file..."
    
    ' Reset system state
    sysState.DataFileLoaded = False
    sysState.AnalysisBuilt = False
    sysState.chartsInitialized = False
    sysState.LastAnalysisHash = ""
    
    ' Clean up data using existing function
    Call ISO16889Mod.CleanupBeforeNewFile
    
    DevToolsMod.TimerEndCount "CleanupBeforeNewFile"
End Sub

'======================================================================
'============== PUBLIC INTERFACE FOR BUTTONS/FORMS ==================
'======================================================================

' Force rebuild analysis (for dashboard buttons)
Public Sub ForceRebuildAnalysis()
    sysState.AnalysisBuilt = False
    sysState.LastAnalysisHash = ""
    Call BuildAnalysisIfNeeded
    Call UpdateDashboard
End Sub

' Quick status check for external calls
Public Function GetSystemStatus() As String
    Dim status As String
    
    status = "System Status:" & vbCrLf
    status = status & "  Data Loaded: " & sysState.DataFileLoaded & vbCrLf
    status = status & "  Analysis Built: " & sysState.AnalysisBuilt & vbCrLf
    status = status & "  Charts Initialized: " & sysState.chartsInitialized & vbCrLf
    
    If Not DataFileMod.TestData Is Nothing Then
        status = status & "  File: " & DataFileMod.TestData.FileName & vbCrLf
        status = status & "  Rows: " & DataFileMod.TestData.DataRowCount & vbCrLf
    End If
    
    GetSystemStatus = status
End Function

' Reset system state (for troubleshooting)
Public Sub ResetSystemState()
    sysState.DataFileLoaded = False
    sysState.AnalysisBuilt = False
    sysState.chartsInitialized = False
    sysState.LastAnalysisHash = ""
    sysState.IsProcessing = False
    
    Debug.Print "System state reset"
End Sub

Public Sub CreateReport()
    If File_Subs.OpenDataFile() Then
        ' Clean up any previous data first
        Call CleanupBeforeNewFile
        
        ' Process the newly loaded file
        Call ProcessDataFileOnce
        
        ' Verify data was loaded
        If DataFileMod.TestData.DataExist Then
            ' Mark that we have new data
            sysState.DataFileLoaded = True
            sysState.AnalysisBuilt = False
            sysState.chartsInitialized = False
            
            ' Build analysis and initialize charts
            Call BuildAnalysisIfNeeded
            
            ' Update dashboard to reflect new state
            Call UpdateDashboard
            
            Debug.Print "CreateReport completed successfully"
        Else
            MsgBox "Data file processing failed. Please check the file format.", vbExclamation
        End If
    End If
End Sub

Private Sub HandleButton(ws As Worksheet, btnName As String, _
                         enableButton As Boolean, Optional macroName As String = "")
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(btnName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    ' Buttons always visible
    shp.Visible = True

    ' Change fill color to indicate enabled/disabled
    If enableButton Then
        shp.Fill.ForeColor.RGB = RGB(68, 114, 196)  ' Blue for active
        If macroName <> "" Then shp.OnAction = macroName
    Else
        shp.Fill.ForeColor.RGB = RGB(191, 191, 191) ' Gray for inactive
        shp.OnAction = "" ' Clear action so it's not clickable
    End If
End Sub

'======================================================================
'============== TOGGLE BUTTONS ======================================
'======================================================================

Public Sub ToggleParticleCounter()
    Dim currentCounter As String
    Dim altCounter As String
    
    Debug.Print "=== ToggleParticleCounter Start ==="
    
    currentCounter = ISO16889Mod.GetISO16889SaveResult(8) ' Index 8 = particle counter phrase
    Debug.Print "Current counter from SaveData: '" & currentCounter & "'"
    
    If currentCounter = "" Then
        MsgBox "Only one particle counter dataset available."
        Exit Sub
    End If
    
    ' ENSURE DATA EXISTS WITH DIAGNOSTICS
    Debug.Print "Calling EnsureDataFileReady..."
    Call DataFileMod.EnsureTestDataReady
    
    ' VERIFY OBJECT STATE AFTER ENSURE
    Debug.Print "After EnsureDataFileReady:"
    Debug.Print "  DataExist: " & DataFileMod.TestData.DataExist
    
    If Not DataFileMod.TestData.DataExist Then
        MsgBox "Cannot toggle particle counter - no test data loaded.", vbExclamation
        Exit Sub
    End If
    
    Select Case currentCounter
        Case "LB"
            Debug.Print "Current is LB - checking for alternates..."
            If ReportFillMod.hasData(DataFileMod.TestData.LS_Sizes) Then
                altCounter = "LS"
                Debug.Print "  Found LS data"
            ElseIf ReportFillMod.hasData(DataFileMod.TestData.LBE_Sizes) Then
                altCounter = "LBE"
                Debug.Print "  Found LBE data"
            Else
                Debug.Print "  No alternate data found"
                MsgBox "No alternate particle counter data available."
                Exit Sub
            End If
        Case "LS", "LBE"
            Debug.Print "Current is " & currentCounter & " - switching to LB"
            altCounter = "LB"
        Case Else
            MsgBox "Unexpected particle counter: " & currentCounter
            Exit Sub
    End Select
    
    Debug.Print "Switching from " & currentCounter & " to " & altCounter
    
    ' Save alternate choice
    Call ISO16889Mod.SetISO16889SaveUserEntry(8, altCounter)
    
    ' Update dashboard
    Call UpdateDashboard
    
    Debug.Print "=== ToggleParticleCounter Complete ==="
End Sub

Public Sub ToggleFilterPressure()
    Dim currentFilter As Long
    Dim pressurePhrase As String

    currentFilter = CLng(ISO16889Mod.GetISO16889SaveResult(7))
    pressurePhrase = CStr(ISO16889Mod.GetISO16889SaveResult(7))
    
    ' If the tag is "TS_DPress", only one dataset is available
    If pressurePhrase = "TS_DPress" Then
        MsgBox "Only one pressure dataset available."
        Exit Sub
    End If
    
    ' Toggle
    If currentFilter = 1 Then
        ISO16889Mod.SetISO16889SaveUserEntry 7, 2
    Else
        ISO16889Mod.SetISO16889SaveUserEntry 7, 1
    End If
    
    UpdateDashboard
End Sub

Public Sub ToggleReportUnits()
    Dim currentUnits As String
    
    currentUnits = ReportFillMod.GetSaveResult(30)
    
    If UCase(currentUnits) = "SI" Then
        ReportFillMod.SetSaveUserEntry 30, "ENG"
    Else
        ReportFillMod.SetSaveUserEntry 30, "SI"
    End If
    
    UpdateDashboard
End Sub

'======================================================================
'============== FORM INTEGRATION ====================================
'======================================================================

' Function for forms to check if data is available
Public Function hasProcessedData() As Boolean
    hasProcessedData = (DataFileMod.EnsureTestDataReady() And DataFileMod.TestData.DataExist)
End Function

' Function for forms to get data status text
Public Function GetDataStatusText() As String
    If hasProcessedData() Then
        GetDataStatusText = "File: " & DataFileMod.TestData.FileName & " (" & DataFileMod.TestData.testType & ")"
    Else
        GetDataStatusText = "No data file loaded"
    End If
End Function

Public Function GetWorkbookStatus() As String
    Dim status As String
    
    status = "=== WORKBOOK STATUS ===" & vbCrLf
    status = status & "TestData: " & DataFileMod.GetTestDataStatus() & vbCrLf
    
    If Not ISO16889Mod.ISO16889ReportData Is Nothing Then
        status = status & "ISO16889: Object exists" & vbCrLf
        status = status & "  - Termination DP: " & ISO16889Mod.ISO16889ReportData.TerminationDP & vbCrLf
        status = status & "  - Termination Time: " & ISO16889Mod.ISO16889ReportData.TerminationTime & vbCrLf
    Else
        status = status & "ISO16889: No object" & vbCrLf
    End If
    
    GetWorkbookStatus = status
End Function

Public Sub RunDiagnostics()
    Debug.Print GetWorkbookStatus()
    
    ' Test data validation
    If DataFileMod.ValidateTestDataIntegrity() Then
        Debug.Print "Data integrity: PASSED"
    Else
        Debug.Print "Data integrity: FAILED"
    End If
End Sub

Private Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(wsName).Name = wsName)
    On Error GoTo 0
End Function

' Save file function (can be called from forms)
Public Function SaveFile() As Boolean
    SaveFile = False
    
    On Error GoTo SaveFailed
    
    If hasProcessedData() Then
        Call File_Subs.SaveAsReport
        SaveFile = True
    Else
        Call File_Subs.SaveAsTemplate
        SaveFile = True
    End If
    
    Exit Function
    
SaveFailed:
    MsgBox "Save operation failed: " & Err.Description, vbCritical
    SaveFile = False
End Function
