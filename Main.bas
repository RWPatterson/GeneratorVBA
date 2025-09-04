Attribute VB_Name = "Main"
Option Explicit

Private Sub CreateReport()
    If OpenDataFile() Then
        
        ' Perform all setup and dashboard update logic
        GenerateReport
        
    End If
End Sub


Public Sub GenerateReport()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    Debug.Print "=== WORKBOOK INITIALIZATION STARTED ==="
    
    ' STEP 1: Ensure basic object structure exists
    Call EnsureBasicObjects
    
    ' STEP 2: Check if we have raw data that needs processing
    If DataFileMod.ShouldProcessRawData() Then
        Debug.Print "Raw data detected - processing file..."
        Call ProcessDataFileAndStandards
    Else
        Debug.Print "No unprocessed raw data found"
        ' Objects exist but no data to process - normal for empty template
    End If
    
    ' STEP 3: Update any dashboard or UI elements
    Call UpdateWorkbookUI
    
    Debug.Print "=== WORKBOOK INITIALIZATION COMPLETED ==="
    
CleanExit:
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Complete Workbook Initialization"
    If Err.Number <> 0 Then
        Debug.Print "GenerateReport Error: " & Err.Description
        MsgBox "Workbook initialization encountered an error: " & Err.Description, vbExclamation
    End If
End Sub

' Step 1: Ensure basic object structure exists (recovery function)
Private Sub EnsureBasicObjects()
    DevToolsMod.TimerStartCount
    
    ' This just ensures TestData object exists and is valid
    ' It does NOT attempt to process any data
    If Not DataFileMod.EnsureTestDataReady() Then
        Debug.Print "Error: Could not establish TestData object"
    Else
        Debug.Print "TestData object ready - Status: " & IIf(DataFileMod.TestData.DataExist, "Contains Data", "Empty")
        
        ' ADDED: Additional debugging
        Debug.Print "TestData details:"
        Debug.Print "  - FileName: '" & DataFileMod.TestData.FileName & "'"
        Debug.Print "  - TestType: '" & DataFileMod.TestData.testType & "'"
        Debug.Print "  - DataRowCount: " & DataFileMod.TestData.DataRowCount
        Debug.Print "  - Object is valid and ready for use"
    End If
    
    DevToolsMod.TimerEndCount "Basic Objects Check"
End Sub

' Step 2: Process data file and set up standards (only when needed)
Private Sub ProcessDataFileAndStandards()
    DevToolsMod.TimerStartCount
    
    ' Process the raw data file
    Debug.Print "Processing data file..."
    Call DataFileMod.ProcessDataFile
    
    ' Verify processing was successful
    If Not DataFileMod.TestData.DataExist Then
        Debug.Print "Warning: Data processing failed"
        DevToolsMod.TimerEndCount "Data Processing (FAILED)"
        Exit Sub
    End If
    
    Debug.Print "Data processing completed successfully"
    Debug.Print "File: " & DataFileMod.TestData.FileName & " (" & DataFileMod.TestData.testType & ")"
    
    ' Set up ISO 16889 analysis (or other standards)
    Debug.Print "Setting up ISO 16889 analysis..."
    Call ISO16889Mod.SetupISO16889ClassModule
    
    DevToolsMod.TimerEndCount "Data Processing and Standards Setup"
End Sub

' Step 3: Update UI and dashboard elements
Private Sub UpdateWorkbookUI()
    DevToolsMod.TimerStartCount
    
    ' Update any dashboard elements, charts, or UI components
    ' This is where you'd put any chart updates, dashboard refreshes, etc.
    
    ' Example: Update charts if data exists
    If DataFileMod.EnsureTestDataReady() And DataFileMod.TestData.DataExist Then
        ' Only update charts if we have processed data
        Call Charts.UpdateCharts
        Debug.Print "Charts updated"
    End If
    
    ' Set active sheet to Dashboard or Home
    On Error Resume Next
    If WorksheetExists("Dashboard") Then
        Sheets("Dashboard").Select
    ElseIf WorksheetExists("Home") Then
        Sheets("Home").Select
    End If
    On Error GoTo 0
    
    DevToolsMod.TimerEndCount "UI Updates"
End Sub

'======================================================================
'================ FILE OPERATIONS ===================================
'======================================================================

' Main function for opening new data files (called from forms)
Public Function LoadNewDataFile() As Boolean
    DevToolsMod.TimerStartCount
    LoadNewDataFile = False
    
    On Error GoTo LoadFailed
    
    Debug.Print "=== LOADING NEW DATA FILE ==="
    
    ' Step 1: Open and load the file
    If Not File_Subs.OpenDataFile() Then
        Debug.Print "File opening was cancelled or failed"
        GoTo LoadFailed
    End If
    
    ' Step 2: Process the newly loaded data
    Call ProcessDataFileAndStandards
    
    ' Step 3: Verify success
    If Not DataFileMod.TestData.DataExist Then
        MsgBox "Data file loaded but processing failed. Please check the file format.", vbExclamation
        GoTo LoadFailed
    End If
    
    ' Step 4: Update UI
    Call UpdateWorkbookUI
    
    LoadNewDataFile = True
    Debug.Print "=== NEW DATA FILE LOADED SUCCESSFULLY ==="
    DevToolsMod.TimerEndCount "New Data File Loading"
    Exit Function
    
LoadFailed:
    Debug.Print "=== DATA FILE LOADING FAILED ==="
    DevToolsMod.TimerEndCount "New Data File Loading (FAILED)"
    LoadNewDataFile = False
End Function

' Force refresh of all analysis (for dashboard "Rebuild" buttons)
Public Sub ForceCompleteRefresh()
    DevToolsMod.TimerStartCount
    
    Debug.Print "=== FORCING COMPLETE REFRESH ==="
    
    ' Clear any cached analysis
    If Not ISO16889Mod.ISO16889ReportData Is Nothing Then
        Call ISO16889Mod.ISO16889ReportData.InvalidateCache
    End If
    
    ' Rebuild everything if we have data
    If DataFileMod.EnsureTestDataReady() And DataFileMod.TestData.DataExist Then
        Call ISO16889Mod.SetupISO16889ClassModule
        Call UpdateWorkbookUI
        Debug.Print "Complete refresh completed"
    Else
        Debug.Print "No data available for refresh"
    End If
    
    DevToolsMod.TimerEndCount "Complete Refresh"
End Sub

'Todo: Update can't update the name of the data file in the text box on first launch.
Sub UpdateDashboard()
    Dim ws As Worksheet
    Dim dataExists As Boolean
    Dim pc As String
    Dim filterSel As Variant
    Dim unitsSel As String
    Dim pressurePhrase As String

    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Check if data exists
    Call DataFileMod.EnsureTestDataReady

    dataExists = DataFileMod.TestData.DataExist

    ' === Buttons that require data ===
    HandleButton ws, "BtnModifyGravs", dataExists, "EditGravimetrics"
    HandleButton ws, "BtnModifyGraphs", dataExists, "ShowChartForm"
    HandleButton ws, "BtnPrintReport", dataExists, "PrintSelectedSheets"
    
    ' === Always visible buttons ===
    HandleButton ws, "BtnCreateReport", True, "CreateReport"
    HandleButton ws, "BtnModifyLogo", True, "ModifyLogoMacro"

    ' === Macro changes & file name display ===
    If dataExists Then
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

    ' === TOGGLE BUTTONS — only active when dataExists ===
    ' --- Particle Counter Toggle ---
    pc = GetISO16889SaveResult(8)
    With ws.Shapes("BtnToggleParticleCounter")
        If dataExists And pc <> "" Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196) ' Active blue
            .OnAction = "ToggleParticleCounter"
            .TextFrame.Characters.Text = "Counter: " & pc
        ElseIf dataExists And pc = "" Then
            ' Only one dataset — disable
            .Fill.ForeColor.RGB = RGB(217, 217, 217)
            .OnAction = ""
            .TextFrame.Characters.Text = "Single Set"
        Else
            ' Data not loaded — disable
            .Fill.ForeColor.RGB = RGB(191, 191, 191)
            .OnAction = ""
            .TextFrame.Characters.Text = "Counter: --"
        End If
    End With

    ' --- Filter Pressure Toggle ---
    filterSel = GetISO16889SaveResult(7)
    pressurePhrase = CStr(GetISO16889SaveResult(7))
    With ws.Shapes("BtnToggleFilterPressure")
        If dataExists And pressurePhrase <> "TS_DPress" Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .OnAction = "ToggleFilterPressure"
            .TextFrame.Characters.Text = "Filter: " & filterSel
        ElseIf dataExists And pressurePhrase = "TS_DPress" Then
            .Fill.ForeColor.RGB = RGB(217, 217, 217)
            .OnAction = ""
            .TextFrame.Characters.Text = "Filter 1 only"
        Else
            .Fill.ForeColor.RGB = RGB(191, 191, 191)
            .OnAction = ""
            .TextFrame.Characters.Text = "Filter: --"
        End If
    End With

    ' --- Report Units Toggle ---
    unitsSel = GetSaveResult(30)
    With ws.Shapes("BtnToggleReportUnits")
        If dataExists Then
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

Private Sub ProcessCurrentStandard()
    Dim currentStandard As String
    currentStandard = DetermineStandardFromFile()
    
    Select Case currentStandard
        Case "ISO16889"
            Call ISO16889Mod.SetupISO16889ClassModule
        Case "ISO23369"
            ' Call ISO23369Mod.SetupISO23369ClassModule  ' Future
        Case Else
            Call ISO16889Mod.SetupISO16889ClassModule  ' Default fallback
    End Select
End Sub

Private Function DetermineStandardFromFile() As String
    If DataFileMod.TestData Is Nothing Then
        DetermineStandardFromFile = "ISO16889"
        Exit Function
    End If
    
    Select Case DataFileMod.TestData.testType
        Case "Single-Pass", "Multipass", "Multipass Series"
            DetermineStandardFromFile = "ISO16889"
        Case "Cyclic Multipass", "Cyclic Series Multipass"
            DetermineStandardFromFile = "ISO23369"
        Case Else
            DetermineStandardFromFile = "ISO16889"
    End Select
End Function


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
        shp.Fill.ForeColor.RGB = RGB(68, 114, 196)  '  for active
        If macroName <> "" Then shp.OnAction = macroName
    Else
        shp.Fill.ForeColor.RGB = RGB(191, 191, 191) ' Gray for inactive
        shp.OnAction = "" ' Clear action so it's not clickable
    End If
End Sub

'Toggle Buttons logic


Public Sub ToggleParticleCounter()
    Dim currentCounter As String
    Dim altCounter As String
    
    Debug.Print "=== ToggleParticleCounter Start ==="
    
    currentCounter = GetISO16889SaveResult(8) ' Index 8 = particle counter phrase
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
            If hasData(DataFileMod.TestData.LS_Sizes) Then
                altCounter = "LS"
                Debug.Print "  Found LS data"
            ElseIf hasData(DataFileMod.TestData.LBE_Sizes) Then
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
    Call SetISO16889SaveUserEntry(8, altCounter)
    
    ' Update dashboard
    Call UpdateDashboard
    
    Debug.Print "=== ToggleParticleCounter Complete ==="
End Sub

Public Sub ToggleFilterPressure()
    Dim currentFilter As Long
    Dim pressurePhrase As String  ' must be String if comparing to text

    currentFilter = CLng(GetISO16889SaveResult(7))
    pressurePhrase = CStr(GetISO16889SaveResult(7))
    
    ' If the tag is "TS_DPress", only one dataset is available
    If pressurePhrase = "TS_DPress" Then
        MsgBox "Only one pressure dataset available."
        Exit Sub
    End If
    
    ' Toggle
    If currentFilter = 1 Then
        SetISO16889SaveUserEntry 7, 2
    Else
        SetISO16889SaveUserEntry 7, 1
    End If
    
    UpdateDashboard
End Sub


Public Sub ToggleReportUnits()
    Dim currentUnits As String
    
    currentUnits = GetSaveResult(30)
    
    If UCase(currentUnits) = "SI" Then
        SetSaveUserEntry 30, "ENG"
    Else
        SetSaveUserEntry 30, "SI"
    End If
    
    UpdateDashboard
End Sub

'======================================================================
'================ FORM INTEGRATION ==================================
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

