VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' In the Save_Data worksheet module (not a regular module)
' Enhanced to prevent triggering during automated updates

Private isAutomatedUpdate As Boolean ' Flag to suppress events during automation

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim saveDataTable As ListObject
    Dim iso16889Table As ListObject
    Dim changedTable As ListObject
    
    ' CRITICAL: Exit immediately if this is an automated update
    If isAutomatedUpdate Then Exit Sub
    
    ' CRITICAL: Only trigger for specific columns that users should edit
    If Not IsUserEditableRange(Target) Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Prevent recursive calls
    Application.EnableEvents = False
    
    ' Check which table was modified
    Set changedTable = Nothing
    
    ' Check SaveDataTable
    On Error Resume Next
    Set saveDataTable = Me.ListObjects("SaveDataTable")
    If Not saveDataTable Is Nothing Then
        If Not Intersect(Target, saveDataTable.Range) Is Nothing Then
            Set changedTable = saveDataTable
        End If
    End If
    
    ' Check ISO16889SaveDataTable
    Set iso16889Table = Me.ListObjects("ISO16889SaveDataTable")
    If Not iso16889Table Is Nothing And changedTable Is Nothing Then
        If Not Intersect(Target, iso16889Table.Range) Is Nothing Then
            Set changedTable = iso16889Table
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' If a SaveData table was changed, handle the update
    If Not changedTable Is Nothing Then
        Call HandleSaveDataChange(Target, changedTable)
    End If
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Debug.Print "Error in Worksheet_Change: " & Err.Description
End Sub

' CRITICAL: Only allow changes to user-editable columns
Private Function IsUserEditableRange(Target As Range) As Boolean
    Dim saveDataTable As ListObject
    Dim iso16889Table As ListObject
    Dim userEntryCol As Long, customDefaultCol As Long, reportValueCol As Long
    
    IsUserEditableRange = False
    
    On Error Resume Next
    
    ' Check SaveDataTable columns
    Set saveDataTable = Me.ListObjects("SaveDataTable")
    If Not saveDataTable Is Nothing Then
        If Not Intersect(Target, saveDataTable.Range) Is Nothing Then
            ' Get column numbers for user-editable columns
            userEntryCol = saveDataTable.ListColumns("User Entry").index
            customDefaultCol = saveDataTable.ListColumns("Custom Default").index
            reportValueCol = saveDataTable.ListColumns("Report Value").index
            
            ' Check if target intersects these specific columns
            If Not Intersect(Target, saveDataTable.ListColumns(userEntryCol).Range) Is Nothing Or _
               Not Intersect(Target, saveDataTable.ListColumns(customDefaultCol).Range) Is Nothing Or _
               Not Intersect(Target, saveDataTable.ListColumns(reportValueCol).Range) Is Nothing Then
                IsUserEditableRange = True
                Exit Function
            End If
        End If
    End If
    
    ' Check ISO16889SaveDataTable columns
    Set iso16889Table = Me.ListObjects("ISO16889SaveDataTable")
    If Not iso16889Table Is Nothing Then
        If Not Intersect(Target, iso16889Table.Range) Is Nothing Then
            userEntryCol = iso16889Table.ListColumns("User Entry").index
            customDefaultCol = iso16889Table.ListColumns("Custom Default").index
            reportValueCol = iso16889Table.ListColumns("Report Value").index
            
            If Not Intersect(Target, iso16889Table.ListColumns(userEntryCol).Range) Is Nothing Or _
               Not Intersect(Target, iso16889Table.ListColumns(customDefaultCol).Range) Is Nothing Or _
               Not Intersect(Target, iso16889Table.ListColumns(reportValueCol).Range) Is Nothing Then
                IsUserEditableRange = True
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

' PUBLIC: Call this before automated updates to suppress change events
Public Sub BeginAutomatedUpdate()
    isAutomatedUpdate = True
End Sub

' PUBLIC: Call this after automated updates to re-enable change events
Public Sub EndAutomatedUpdate()
    isAutomatedUpdate = False
End Sub

' Handle changes to SaveData tables
Private Sub HandleSaveDataChange(ByVal Target As Range, ByVal changedTable As ListObject)
    Dim changedRow As Long
    Dim changedCol As Long
    Dim changedID As Long
    Dim columnHeader As String
    Dim newValue As Variant
    Dim isISO16889Table As Boolean
    
    ' Determine which table and get basic info
    isISO16889Table = (changedTable.Name = "ISO16889SaveDataTable")
    
    ' Get row and column within the table
    changedRow = Target.Row - changedTable.HeaderRowRange.Row
    changedCol = Target.Column - changedTable.Range.Column + 1
    
    ' Skip header row changes
    If changedRow <= 0 Then Exit Sub
    
    ' Get the ID and column header
    changedID = changedTable.DataBodyRange(changedRow, 1).Value ' Assuming ID is column 1
    columnHeader = changedTable.HeaderRowRange(changedCol).Value
    newValue = Target.Value
    
    ' Handle different column types
    Select Case columnHeader
        Case "User Entry"
            Call HandleUserEntryChange(changedID, newValue, isISO16889Table)
        Case "Custom Default"
            Call HandleCustomDefaultChange(changedID, newValue, isISO16889Table)
        Case "Report Value"
            ' User edited Report Value directly - this overrides the formula
            Call HandleReportValueDirectEdit(changedID, newValue, isISO16889Table)
        Case Else
            ' Other columns changed - may need to recalculate
    End Select
End Sub

' Handle User Entry changes
Private Sub HandleUserEntryChange(ID As Long, newValue As Variant, isISO16889 As Boolean)
    ' User Entry changes should trigger automatic Report Value recalculation
    ' The Excel formula will handle this automatically, but we may need to trigger actions
    
    If isISO16889 Then
        ' Check if this change affects termination parameters
        Select Case ID
            Case 2 ' Termination DP override
                If IsNumeric(newValue) And newValue <> "" Then
                    Call ValidateAndApplyDPOverride(CDbl(newValue))
                End If
            Case 7 ' Filter selection
                If IsNumeric(newValue) And newValue <> "" Then
                    Call ValidateAndApplyFilterOverride(CInt(newValue))
                End If
            Case 8 ' Sensor selection
                If newValue <> "" Then
                    Call ValidateAndApplySensorOverride(CStr(newValue))
                End If
        End Select
    End If
    
    ' Log the change for audit trail
    Call LogDataChange("User Entry", ID, newValue, isISO16889)
End Sub

' Handle Custom Default changes
Private Sub HandleCustomDefaultChange(ID As Long, newValue As Variant, isISO16889 As Boolean)
    ' Custom Default changes affect template behavior
    Call LogDataChange("Custom Default", ID, newValue, isISO16889)
    
    ' Could trigger template save prompt
    Call MarkTemplateAsModified
End Sub

' Handle direct Report Value edits (bypassing formula)
Private Sub HandleReportValueDirectEdit(ID As Long, newValue As Variant, isISO16889 As Boolean)
    ' User edited Report Value directly - treat as User Entry override
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = Me
    
    If isISO16889 Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    ' Move the direct edit to User Entry column and restore formula
    tbl.DataBodyRange(ID, 4).Value = newValue ' User Entry column
    
    ' Restore the formula in Report Value column
    Dim formulaText As String
    formulaText = "=IF(NOT(ISBLANK([@[User Entry]])), [@[User Entry]], IF(NOT(ISBLANK([@[Custom Default]])), [@[Custom Default]], IF(NOT(ISBLANK([@[From Data]])), [@[From Data]], 0)))"
    
    tbl.DataBodyRange(ID, 3).Formula = formulaText ' Report Value column
    
    ' Handle the change as if it were a User Entry change
    Call HandleUserEntryChange(ID, newValue, isISO16889)
    
    ' Notify user of the conversion
    MsgBox "Direct edit moved to User Entry column to maintain data integrity.", vbInformation
End Sub

' Validation functions with informative messages
Private Sub ValidateAndApplyDPOverride(newDP As Double)
    If ISO16889Mod.ISO16889ReportData Is Nothing Then Exit Sub
    
    Dim actualTerminationDP As Double
    actualTerminationDP = ISO16889Mod.ISO16889ReportData.GetActualTerminationDP()
    
    If ISO16889Mod.ISO16889ReportData.IsValidDPOverride(newDP) Then
        ' Mark for rebuild
        Call ISO16889Mod.ISO16889ReportData.InvalidateCache
        ' Could trigger automatic rebuild or show rebuild prompt
        Call PromptForISO16889Rebuild("DP override")
    Else
        MsgBox "Invalid DP override: " & newDP & vbCrLf & vbCrLf & _
               "Maximum allowed value: " & actualTerminationDP & vbCrLf & _
               "You can only trim data to a lower termination point, not extrapolate beyond the actual test termination.", _
               vbExclamation, "DP Override Validation"
    End If
End Sub

Private Sub ValidateAndApplyFilterOverride(newFilter As Integer)
    If ISO16889Mod.ISO16889ReportData Is Nothing Then Exit Sub
    
    If ISO16889Mod.ISO16889ReportData.IsValidFilterChoice(newFilter) Then
        Call ISO16889Mod.ISO16889ReportData.InvalidateCache
        Call PromptForISO16889Rebuild("filter selection")
    Else
        Dim availableFilters As String
        availableFilters = ISO16889Mod.ISO16889ReportData.GetAvailableFilterOptions()
        
        MsgBox "Invalid filter selection: Filter " & newFilter & vbCrLf & vbCrLf & _
               "Available filters: " & availableFilters & vbCrLf & _
               "Filter " & newFilter & " data is not available in this test file.", _
               vbExclamation, "Filter Selection Validation"
    End If
End Sub

Private Sub ValidateAndApplySensorOverride(newSensor As String)
    If ISO16889Mod.ISO16889ReportData Is Nothing Then Exit Sub
    
    If ISO16889Mod.ISO16889ReportData.IsValidSensorChoice(newSensor) Then
        ' Sensor changes don't require rebuild, just display update
        ' DON'T update the class directly - let the user intervention system handle it
        Call PromptForSensorChange(newSensor)
    Else
        Dim availableSensors As String
        availableSensors = ISO16889Mod.ISO16889ReportData.GetAvailableSensorOptions()
        
        MsgBox "Invalid sensor selection: " & newSensor & vbCrLf & vbCrLf & _
               "Available sensors: " & availableSensors & vbCrLf & _
               newSensor & " sensor data is not available in this test file.", _
               vbExclamation, "Sensor Selection Validation"
    End If
End Sub

' Prompt for sensor change (no rebuild needed)
Private Sub PromptForSensorChange(newSensor As String)
    Dim result As VbMsgBoxResult
    
    result = MsgBox("Change particle counter sensor to " & newSensor & "?" & vbCrLf & _
                   "This will update report displays but does not require rebuilding analysis.", _
                   vbYesNo + vbQuestion, "Sensor Change")
    
    If result = vbYes Then
        ' Update dashboard if available
        Call UpdateDashboard ' If you have dashboard update function
        MsgBox "Sensor selection updated. Report displays will now use " & newSensor & " data.", vbInformation
    End If
End Sub

' Prompt user for rebuild when critical parameters change
Private Sub PromptForISO16889Rebuild(changeType As String)
    Dim result As VbMsgBoxResult
    
    result = MsgBox("Your " & changeType & " change requires rebuilding ISO 16889 analysis." & vbCrLf & _
                   "Rebuild now?", vbYesNo + vbQuestion, "Rebuild Required")
    
    If result = vbYes Then
        Call ISO16889Mod.ForceRebuildAnalysis
    End If
End Sub

' Audit trail logging
Private Sub LogDataChange(changeType As String, ID As Long, newValue As Variant, isISO16889 As Boolean)
    Dim logEntry As String
    Dim tableType As String
    
    tableType = IIf(isISO16889, "ISO16889", "SaveData")
    
    logEntry = Now & " - " & tableType & " ID " & ID & " - " & changeType & " changed to: " & newValue
    
    ' Could log to hidden sheet, file, or debug window
    Debug.Print logEntry
End Sub

' Mark template as modified
Private Sub MarkTemplateAsModified()
    ' Could set a flag or update a timestamp
    ' ThisWorkbook.CustomDocumentProperties("TemplateModified") = True
End Sub

