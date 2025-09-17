Attribute VB_Name = "SaveDataMod"
' SaveDataMod Module - Bridge between forms and tables

Option Explicit

' Interface for both form and direct editing
Public Enum DataSource
    FromData = 1
    CustomDefault = 2
    UserEntry = 3
End Enum

' Get current effective value (what Report Value shows)
Public Function GetEffectiveValue(tableType As String, ID As Long) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = Sheets("Save_Data")
    
    If tableType = "ISO16889" Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    GetEffectiveValue = tbl.DataBodyRange(ID, 3).value ' Report Value column
End Function

' Set value with proper priority handling
Public Sub SetValue(tableType As String, ID As Long, newValue As Variant, source As DataSource)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim targetColumn As Long
    
    Set ws = Sheets("Save_Data")
    
    ' CRITICAL: Suppress change events during programmatic updates
    Call SaveDataMod.BeginAutomatedUpdate
    
    On Error GoTo CleanupEvents
    
    If tableType = "ISO16889" Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    ' Determine target column based on source
    Select Case source
        Case DataSource.FromData
            targetColumn = 6 ' From Data column
        Case DataSource.CustomDefault
            targetColumn = 5 ' Custom Default column
        Case DataSource.UserEntry
            targetColumn = 4 ' User Entry column
    End Select
    
    ' Set the value
    tbl.DataBodyRange(ID, targetColumn).value = newValue
    
    ' Handle special ISO 16889 cases
    If tableType = "ISO16889" Then
        Call HandleISO16889ValueChange(ID, newValue, source)
    End If
    
CleanupEvents:
    ' CRITICAL: Re-enable change events
    Call SaveDataMod.EndAutomatedUpdate
End Sub

' Clear value from specific source
Public Sub ClearValue(tableType As String, ID As Long, source As DataSource)
    Call SetValue(tableType, ID, "", source)
End Sub

' Get value from specific source
Public Function GetSourceValue(tableType As String, ID As Long, source As DataSource) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sourceColumn As Long
    
    Set ws = Sheets("Save_Data")
    
    If tableType = "ISO16889" Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    Select Case source
        Case DataSource.FromData
            sourceColumn = 6
        Case DataSource.CustomDefault
            sourceColumn = 5
        Case DataSource.UserEntry
            sourceColumn = 4
    End Select
    
    GetSourceValue = tbl.DataBodyRange(ID, sourceColumn).value
End Function

' Check if value has been overridden by user
Public Function IsUserOverridden(tableType As String, ID As Long) As Boolean
    IsUserOverridden = (GetSourceValue(tableType, ID, DataSource.UserEntry) <> "")
End Function

' Reset to default (clear user entry)
Public Sub ResetToDefault(tableType As String, ID As Long)
    Call ClearValue(tableType, ID, DataSource.UserEntry)
End Sub

' Handle ISO 16889 specific value changes
Private Sub HandleISO16889ValueChange(ID As Long, newValue As Variant, source As DataSource)
    ' Only handle User Entry changes that affect analysis
    If source <> DataSource.UserEntry Then Exit Sub
    
    Select Case ID
        Case 2 ' Termination DP
            If IsNumeric(newValue) And newValue <> "" Then
                Call ValidateAndApplyDPChange(CDbl(newValue))
            End If
        Case 7 ' Filter selection
            If IsNumeric(newValue) And newValue <> "" Then
                Call ValidateAndApplyFilterChange(CInt(newValue))
            End If
        Case 8 ' Sensor selection
            If newValue <> "" Then
                Call ValidateAndApplySensorChange(CStr(newValue))
            End If
    End Select
End Sub

' Validation functions (delegate to existing logic)
Private Sub ValidateAndApplyDPChange(newDP As Double)
    ' Use existing validation from worksheet change event
    Call Sheets("Save_Data").ValidateAndApplyDPOverride(newDP)
End Sub

Private Sub ValidateAndApplyFilterChange(newFilter As Integer)
    Call Sheets("Save_Data").ValidateAndApplyFilterOverride(newFilter)
End Sub

Private Sub ValidateAndApplySensorChange(newSensor As String)
    Call Sheets("Save_Data").ValidateAndApplySensorOverride(newSensor)
End Sub

' Form interface functions
Public Function GetDisplayName(tableType As String, ID As Long) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = Sheets("Save_Data")
    
    If tableType = "ISO16889" Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    GetDisplayName = tbl.DataBodyRange(ID, 2).value ' Display Name column
End Function

' Get all IDs for a table (for populating forms)
Public Function GetAllIDs(tableType As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = Sheets("Save_Data")
    
    If tableType = "ISO16889" Then
        Set tbl = ws.ListObjects("ISO16889SaveDataTable")
    Else
        Set tbl = ws.ListObjects("SaveDataTable")
    End If
    
    GetAllIDs = tbl.DataBodyRange.Columns(1).value ' ID column
End Function

' Validation helpers for forms with informative messages
Public Function ValidateISO16889Value(ID As Long, newValue As Variant, ByRef errorMessage As String) As Boolean
    errorMessage = "" ' Clear any previous error message
    
    If ISO16889Mod.ISO16889ReportData Is Nothing Then
        ValidateISO16889Value = True ' Allow if no data loaded
        Exit Function
    End If
    
    Select Case ID
        Case 2 ' DP override
            If Not IsNumeric(newValue) Then
                errorMessage = "DP value must be numeric."
                ValidateISO16889Value = False
                Exit Function
            End If
            
            Dim maxDP As Double
            maxDP = ISO16889Mod.ISO16889ReportData.GetActualTerminationDP()
            
            If CDbl(newValue) > maxDP Then
                errorMessage = "DP override (" & newValue & ") exceeds maximum allowed value of " & maxDP & "." & vbCrLf & _
                             "You can only trim data to a lower termination point, not extrapolate beyond the actual test termination."
                ValidateISO16889Value = False
            Else
                ValidateISO16889Value = True
            End If
            
        Case 7 ' Filter selection
            If Not IsNumeric(newValue) Then
                errorMessage = "Filter selection must be numeric (1 or 2)."
                ValidateISO16889Value = False
                Exit Function
            End If
            
            If Not ISO16889Mod.ISO16889ReportData.IsValidFilterChoice(CInt(newValue)) Then
                Dim availableFilters As String
                availableFilters = ISO16889Mod.ISO16889ReportData.GetAvailableFilterOptions()
                errorMessage = "Filter " & newValue & " is not available." & vbCrLf & _
                             "Available filters: " & availableFilters
                ValidateISO16889Value = False
            Else
                ValidateISO16889Value = True
            End If
            
        Case 8 ' Sensor selection
            If Not ISO16889Mod.ISO16889ReportData.IsValidSensorChoice(CStr(newValue)) Then
                Dim availableSensors As String
                availableSensors = ISO16889Mod.ISO16889ReportData.GetAvailableSensorOptions()
                errorMessage = "Sensor '" & newValue & "' is not available." & vbCrLf & _
                             "Available sensors: " & availableSensors
                ValidateISO16889Value = False
            Else
                ValidateISO16889Value = True
            End If
            
        Case Else
            ValidateISO16889Value = True
    End Select
End Function

' Get original calculated value (From Data column)
Public Function GetOriginalValue(tableType As String, ID As Long) As Variant
    If tableType = "ISO16889" Then
        GetOriginalValue = ISO16889Mod.GetISO16889SaveFromData(ID)
    Else
        GetOriginalValue = GetSourceValue(tableType, ID, DataSource.FromData)
    End If
End Function

' Dashboard integration functions - work with User Entry column only
Public Sub SetDashboardValue(ID As Long, newValue As Variant)
    ' Dashboard changes should only affect User Entry column
    Call SetValue("ISO16889", ID, newValue, DataSource.UserEntry)
End Sub

Public Function GetDashboardValue(ID As Long) As Variant
    ' Dashboard should read from effective value (Report Value)
    GetDashboardValue = GetEffectiveValue("ISO16889", ID)
End Function

Public Function GetOriginalDashboardValue(ID As Long) As Variant
    ' Dashboard can show original value for reference
    GetOriginalDashboardValue = GetOriginalValue("ISO16889", ID)
End Function

' Clear user override (dashboard "Reset" functionality)
Public Sub ResetDashboardValue(ID As Long)
    Call SetValue("ISO16889", ID, "", DataSource.UserEntry)
End Sub

'======================================================================
'================ AUTOMATION CONTROL FUNCTIONS ======================
'======================================================================

' Control change event suppression for Save_Data worksheet
Public Sub BeginAutomatedUpdate()
    ' Call the worksheet method to suppress change events
    Sheets("Save_Data").BeginAutomatedUpdate
End Sub

Public Sub EndAutomatedUpdate()
    ' Re-enable change events
    Sheets("Save_Data").EndAutomatedUpdate
End Sub

' Safe wrapper that ensures events are always re-enabled
Public Sub ExecuteWithSuppressedEvents(ByVal codeToExecute As String)
    On Error GoTo CleanupEvents
    
    Call BeginAutomatedUpdate
    
    ' Execute the provided code (this would need to be implemented differently)
    ' For now, we'll use the direct approach in calling code
    
CleanupEvents:
    Call EndAutomatedUpdate
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.source, Err.Description
    End If
End Sub
