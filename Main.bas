Attribute VB_Name = "Main"
Option Explicit

Private Sub CreateReport()
    If OpenDataFile() Then
        
        ' Perform all setup and dashboard update logic
        GenerateReport
        
    End If
End Sub


Public Sub GenerateReport()
'This code is called by the workbook open function in the ThisWorkbook excel object.

Application.ScreenUpdating = False

    'create DataFile class mod and parse data if present
        DataFileMod.ProcessDataFile
        
    'perform standard specific operations
        Call ProcessCurrentStandard
    
   ' Update the Dashboard buttons/appearance based on data presence
        UpdateDashboard
    
    'Show Dashboard
    Sheets("Dashboard").Select
    
Application.ScreenUpdating = True

End Sub



Sub UpdateDashboard()
    Dim ws As Worksheet
    Dim dataExists As Boolean
    Dim pc As String
    Dim filterSel As Variant
    Dim unitsSel As String
    Dim pressurePhrase As String

    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Check if data exists
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
    
    currentCounter = GetISO16889SaveResult(8) ' Index 8 = particle counter phrase
    
    If currentCounter = "" Then
        MsgBox "Only one particle counter dataset available."
        Exit Sub
    End If
    
    Select Case currentCounter
        Case "LB"
            If hasData(TestData.LS_Sizes) Then
                altCounter = "LS"
            ElseIf hasData(TestData.LBE_Sizes) Then
                altCounter = "LBE"
            Else
                MsgBox "No alternate particle counter data available."
                Exit Sub
            End If
        Case "LS", "LBE"
            altCounter = "LB"
        Case Else
            MsgBox "Unexpected particle counter: " & currentCounter
            Exit Sub
    End Select
    
    ' Save alternate choice
    SetISO16889SaveUserEntry 8, altCounter
    
    ' Update dashboard
    UpdateDashboard
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



