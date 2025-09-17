VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISO16889ClassMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************************
'*************  ISO16889 Report Details Storage - ENHANCED *************
'************************************************************************
Public WorkbookInstance As Workbook 'Sets the parent of the ClassModule to this workbook

'test termination details
Public TerminationTag As String 'what differential pressure tag was used for the terminal pressure analog column
Public TerminationDP As Double 'what pressure termination happened at
Public TerminationTime As Double 'elapsed time when first termination took place
Public TerminationFilter As Integer 'filter 1 or 2 which terminated
Public TerminationSizePhrase As String 'what particle counters evaluated the termination filter

'Set to false when running singlepass or series tests reporting on individual filters
Public RetainedMassValid As Boolean
Public InjectedMass As Variant

'Here we store the clumps of analog data
Public C_Times As Variant
Public C_Pressures As Variant
Public C_Masses As Variant

'Here we store the clumps of count data
Public CU_LB As Variant
Public CD_LB As Variant
Public Beta_LB As Variant

Public CU_LS As Variant
Public CD_LS As Variant
Public Beta_LS As Variant

Public CU_LBE As Variant
Public CD_LBE As Variant
Public Beta_LBE As Variant

'test result details
Public ISO16889SizeAtBeta2 As String
Public ISO16889SizeAtBeta10 As String
Public ISO16889SizeAtBeta75 As String
Public ISO16889SizeAtBeta100 As String
Public ISO16889SizeAtBeta200 As String
Public ISO16889SizeAtBeta1000 As String

' Cache for expensive calculations
Private Type analysisCache
    LastTerminationTime As Double
    LastTerminationDP As Double
    LastTerminationFilter As Integer
    TablesBuilt As Boolean
    IsValid As Boolean
End Type
Private analysisCache As analysisCache

Private Sub Class_Initialize()
    'Initialize cache
    analysisCache.IsValid = False
    analysisCache.TablesBuilt = False    ' NEW: Initialize tables built flag
End Sub

'======================================================================
'============== PUBLIC METHODS (Called by ISO16889Mod) ==============
'======================================================================

' Main validation method - called by module
Public Function ValidateFileCompatibility() As Boolean
    Dim testType As String
    Dim errorMsg As String
    
    ValidateFileCompatibility = False ' Assume failure
    
    testType = DataFileMod.TestData.testType
    
    ' Check for incompatible test types
    Select Case testType
        Case "Data Only"
            errorMsg = "This file contains 'Data Only' and cannot generate ISO 16889 reports."
            
        Case "PQ"
            errorMsg = "This file contains 'PQ' data and cannot generate ISO 16889 reports."
            
        Case "Cyclic Multipass"
            errorMsg = "This file contains 'Cyclic Multipass' data (ISO 23369)." & vbCrLf & _
                       "Please use the ISO 23369 report writer instead."
            
        Case "Cyclic Series Multipass"
            errorMsg = "This file contains 'Cyclic Series Multipass' data (ISO 23369)." & vbCrLf & _
                       "Please use the ISO 23369 report writer instead."
            
        Case "Single-Pass", "Multipass", "Multipass Series", ""
            ' These are valid for ISO 16889
            ValidateFileCompatibility = True
            
        Case Else
            errorMsg = "Unknown test type: '" & testType & "'." & vbCrLf & _
                       "Cannot determine if this file is compatible with ISO 16889."
    End Select
    
    ' Check for essential data availability
    If ValidateFileCompatibility Then
        If Not ValidateEssentialData(errorMsg) Then
            ValidateFileCompatibility = False
        End If
    End If
    
    ' Display error and exit if validation failed
    If Not ValidateFileCompatibility Then
        MsgBox "ISO 16889 Processing Stopped" & vbCrLf & vbCrLf & errorMsg, _
               vbCritical + vbOKOnly, "Incompatible File Type"
        
        ' Clear any partial data
        Call ClearISO16889Data
    End If
End Function

' Enhanced test type evaluation - called by module
Public Function EvaluateByTestType() As Boolean
    EvaluateByTestType = True
    
    Select Case DataFileMod.TestData.testType
        Case "Single-Pass"
            Me.RetainedMassValid = False
            If Not SetISO16889DiffPressTag() Then
                EvaluateByTestType = False
            End If
            
        Case "Multipass"
            Me.RetainedMassValid = True
            If Not SetISO16889DiffPressTag() Then
                EvaluateByTestType = False
            End If
            
        Case "Multipass Series"
            Me.RetainedMassValid = False
            If Not SetISO16889DiffPressTag() Then
                EvaluateByTestType = False
            End If
            
        Case Else
            ' Default handling for unspecified or unknown test types
            Me.RetainedMassValid = False
            If Not SetISO16889DiffPressTag() Then
                EvaluateByTestType = False
            End If
    End Select
End Function

' Enhanced sensor evaluation - called by module
Public Function EvaluateSelectedSensors() As Boolean
    EvaluateSelectedSensors = False
    
    Dim ExRaDs As Boolean
    Dim TermFilter As Integer
    Dim hasAnyCountData As Boolean
    
    ' Check if any particle count data exists
    hasAnyCountData = (hasData(DataFileMod.TestData.LB_Sizes) Or _
                      hasData(DataFileMod.TestData.LS_Sizes) Or _
                      hasData(DataFileMod.TestData.LBE_Sizes))
    
    If Not hasAnyCountData Then
        MsgBox "No particle size data found for any sensor type." & vbCrLf & _
               "ISO 16889 analysis requires particle count data.", vbCritical
        Exit Function
    End If
    
    ExRaDs = ReportFillMod.GetSaveResult(58)
    TermFilter = Me.TerminationFilter
    
    If ExRaDs Then
        If TermFilter > 1 Then
            If hasData(DataFileMod.TestData.LS_Sizes) Then
                Me.TerminationSizePhrase = "LS"
            ElseIf hasData(DataFileMod.TestData.LBE_Sizes) Then
                Me.TerminationSizePhrase = "LBE"
            Else
                MsgBox "ExRaDs system with Filter 2 termination requires LS or LBE sensor data.", vbCritical
                Exit Function
            End If
        Else
            If hasData(DataFileMod.TestData.LB_Sizes) Then
                Me.TerminationSizePhrase = "LB"
            Else
                MsgBox "LB sensor data required but not found.", vbCritical
                Exit Function
            End If
        End If
    Else
        ' Non-ExRaDs system - use default sensor
        If hasData(DataFileMod.TestData.LB_Sizes) Then
            Me.TerminationSizePhrase = ""
        Else
            MsgBox "Standard system requires LB sensor data.", vbCritical
            Exit Function
        End If
    End If
    
    EvaluateSelectedSensors = True
End Function

' Main termination calculation - called by module
Public Function SetISO16889DiffPressTag() As Boolean
    SetISO16889DiffPressTag = False
    
    Dim TerminalDP As Variant
    
    Select Case DataFileMod.TestData.TestSetup
        Case "Spin On", "Pressure", "Suction"
            ' Validate termination DP exists
            TerminalDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", 1)
            
            If IsEmpty(TerminalDP) Or TerminalDP <= 0 Then
                MsgBox "Terminal DP setpoint not found or invalid for " & DataFileMod.TestData.TestSetup & " test setup.", vbCritical
                Exit Function
            End If
            
            ' Validate pressure tag exists
            If IsEmpty(DataFileMod.TestData.GetAnalogTagData("TS_DPress")) Then
                MsgBox "Required pressure data (TS_DPress) not found in analog data.", vbCritical
                Exit Function
            End If
            
            Me.TerminationTag = "TS_DPress"
            Me.TerminationDP = TerminalDP
            Me.TerminationFilter = 1
            Call SetTerminationTime
            
        Case "Suction & Pressure", "Two Pressure"
            If DataFileMod.TestData.MidstreamFlag And HasDualFilterData() Then
                If Not SetFilter1or2TerminationFields() Then
                    Exit Function
                End If
                Call SetTerminationTime
            Else
                ' Fallback to single filter mode with validation
                TerminalDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", 1)
                
                If IsEmpty(TerminalDP) Or TerminalDP <= 0 Then
                    MsgBox "Terminal DP setpoint not found for single filter fallback mode.", vbCritical
                    Exit Function
                End If
                
                Me.TerminationTag = "TS_DPress"
                Me.TerminationDP = TerminalDP
                Me.TerminationFilter = 1
                Call SetTerminationTime
            End If
             
        Case Else
            MsgBox "Unrecognized Test Setup: '" & DataFileMod.TestData.TestSetup & "'." & vbCrLf & _
                   "Cannot proceed with ISO 16889 analysis.", vbCritical
            Exit Function
    End Select
    
    SetISO16889DiffPressTag = True
End Function

' Enhanced termination time calculation - called by module
Public Sub SetTerminationTime()
    Dim DPTag As String
    Dim DPSetP As Double
    Dim DPData As Variant
    Dim TestTimes As Variant
    
    If Me.TerminationTag = "" Then Exit Sub
    
    DPTag = Me.TerminationTag
    DPSetP = Me.TerminationDP
    DPData = DataFileMod.TestData.GetAnalogTagData(DPTag)
    TestTimes = DataFileMod.TestData.Times
    
    ' Validate data arrays
    If IsEmpty(DPData) Or IsEmpty(TestTimes) Then
        Me.TerminationTime = 0
        Exit Sub
    End If
    
    ' Use enhanced interpolation for better accuracy
    Me.TerminationTime = MathMod.FastLinearInterpolation(DPSetP, DPData, TestTimes)
End Sub

' Set clump times - called by module
Public Sub SetClumpTimes()
    Dim TerminationTime As Double
    Dim tempC_Times() As Double
    Dim i As Integer
    
    TerminationTime = Me.TerminationTime
    ReDim tempC_Times(1 To 10)
    
    For i = 1 To 10
        tempC_Times(i) = TerminationTime * (i / 10)
    Next i
    
    Me.C_Times = tempC_Times
End Sub

' Set clump pressures - called by module
Public Sub SetClumpPressures()
    Dim tempC_Pressures() As Double
    Dim i As Integer
    Dim DPData As Variant
    Dim TestTimes As Variant
    Dim targetTime As Double
    
    ' Get the actual pressure and time data from the test
    DPData = DataFileMod.TestData.GetAnalogTagData(Me.TerminationTag)
    TestTimes = DataFileMod.TestData.Times
    
    ' Validate we have the required data
    If IsEmpty(DPData) Or IsEmpty(TestTimes) Then
        ' Fallback to percentage calculation if interpolation data unavailable
        ReDim tempC_Pressures(1 To 10)
        For i = 1 To 10
            tempC_Pressures(i) = Me.TerminationDP * (i / 10)
        Next i
        Me.C_Pressures = tempC_Pressures
        Exit Sub
    End If
    
    ReDim tempC_Pressures(1 To 10)
    
    ' Calculate actual pressures at each time interval using interpolation
    For i = 1 To 10
        ' Get the target time for this interval (10%, 20%, etc. of termination time)
        targetTime = Me.TerminationTime * (i / 10)
        
        ' Interpolate the actual pressure at this time point
        tempC_Pressures(i) = MathMod.FastLinearInterpolation(targetTime, TestTimes, DPData)
    Next i
    
    Me.C_Pressures = tempC_Pressures
End Sub

' Enhanced function to get available filter options for dashboard
Public Function GetAvailableFilterOptions() As String
    Dim options As String
    options = "1"  ' Filter 1 always available
    
    If HasDualFilterData() Then
        options = options & ",2"
    End If
    
    GetAvailableFilterOptions = options
End Function

' Enhanced function to get available sensor options for dashboard
Public Function GetAvailableSensorOptions() As String
    Dim options As String
    
    If hasData(DataFileMod.TestData.LB_Sizes) Then
        options = "LB"
    End If
    
    If hasData(DataFileMod.TestData.LS_Sizes) Then
        If options <> "" Then options = options & ",LS" Else options = "LS"
    End If
    
    If hasData(DataFileMod.TestData.LBE_Sizes) Then
        If options <> "" Then options = options & ",LBE" Else options = "LBE"
    End If
    
    GetAvailableSensorOptions = options
End Function

' User override validation methods - called by module
Public Function IsValidFilterChoice(filterNum As Integer) As Boolean
    Select Case filterNum
        Case 1
            ' Filter 1 always available if we have any pressure data
            IsValidFilterChoice = True
        Case 2
            ' Filter 2 only available if we have dual filter setup and midstream flag
            IsValidFilterChoice = DataFileMod.TestData.MidstreamFlag And HasDualFilterData()
        Case Else
            IsValidFilterChoice = False
    End Select
End Function

Public Function IsValidDPOverride(newDP As Double) As Boolean
    Dim actualTerminationDP As Double
    
    ' Get the actual termination DP from the test
    actualTerminationDP = GetActualTerminationDP()
    
    ' User can only set lower DP (trim data), not extrapolate
    IsValidDPOverride = (newDP > 0 And newDP <= actualTerminationDP)
End Function

Public Function IsValidSensorChoice(sensorType As String) As Boolean
    Select Case UCase(sensorType)
        Case "LB"
            IsValidSensorChoice = hasData(DataFileMod.TestData.LB_Sizes)
        Case "LS"
            IsValidSensorChoice = hasData(DataFileMod.TestData.LS_Sizes)
        Case "LBE"
            IsValidSensorChoice = hasData(DataFileMod.TestData.LBE_Sizes)
        Case Else
            IsValidSensorChoice = False
    End Select
End Function

' Apply user overrides - called by module
Public Sub ApplyFilterOverride(newFilter As Integer)
    Me.TerminationFilter = newFilter
    
    ' Update termination tag based on new filter selection
    Select Case newFilter
        Case 1
            Me.TerminationTag = "TS_DPress"  ' or TS_PreDPress depending on setup
        Case 2
            Me.TerminationTag = "TS_FinalDPress"
    End Select
    
    ' Get the DP setpoint for this filter
    Me.TerminationDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", newFilter)
    
    ' Recalculate termination time for new filter/DP combination
    Call SetTerminationTime
End Sub

' Cache management - called by module
Public Function IsRebuildRequired() As Boolean
    With analysisCache
        If Not .IsValid Then
            IsRebuildRequired = True
            Debug.Print "Cache invalid - rebuild required"
        ElseIf Not .TablesBuilt Then
            IsRebuildRequired = True
            Debug.Print "No tables built - rebuild required"
        ElseIf .LastTerminationTime <> Me.TerminationTime Then
            IsRebuildRequired = True
            Debug.Print "Termination time changed - rebuild required"
        ElseIf .LastTerminationDP <> Me.TerminationDP Then
            IsRebuildRequired = True
            Debug.Print "Termination DP changed - rebuild required"
        ElseIf .LastTerminationFilter <> Me.TerminationFilter Then
            IsRebuildRequired = True
            Debug.Print "Termination filter changed - rebuild required"
        Else
            IsRebuildRequired = False
            Debug.Print "Cache valid - no rebuild needed"
        End If
    End With
End Function

Public Sub UpdateCache()
    With analysisCache
        .LastTerminationTime = Me.TerminationTime
        .LastTerminationDP = Me.TerminationDP
        .LastTerminationFilter = Me.TerminationFilter
        .TablesBuilt = True
        .IsValid = True
    End With
    Debug.Print "Cache updated - tables marked as built"
End Sub

Public Sub InvalidateCache()
    analysisCache.IsValid = False
    analysisCache.TablesBuilt = False    ' NEW: Mark tables as not built
    Debug.Print "Cache invalidated - tables marked as not built"
End Sub

Public Sub MarkTablesAsBuilt()
    analysisCache.TablesBuilt = True
End Sub

Public Sub CalculateBetaSizes()
    DevToolsMod.TimerStartCount
    
    ' Only calculate if we have sensor data and valid beta arrays
    If Me.TerminationSizePhrase = "" Then
        Debug.Print "No sensor selected - skipping beta size calculations"
        Exit Sub
    End If
    
    ' Validate that beta arrays are populated
    Dim betaArray As Variant
    Select Case Me.TerminationSizePhrase
        Case "LS"
            betaArray = Me.Beta_LS
        Case "LBE"
            betaArray = Me.Beta_LBE
        Case Else
            betaArray = Me.Beta_LB
    End Select
    
    If IsEmpty(betaArray) Then
        Debug.Print "Beta array is empty - cannot calculate beta sizes"
        Exit Sub
    End If
    
    ' Calculate sizes for standard beta ratios using existing function
    On Error GoTo CalculationError
    
    Me.ISO16889SizeAtBeta2 = GetISO16889SizeGivenBeta(2)
    Me.ISO16889SizeAtBeta10 = GetISO16889SizeGivenBeta(10)
    Me.ISO16889SizeAtBeta75 = GetISO16889SizeGivenBeta(75)
    Me.ISO16889SizeAtBeta100 = GetISO16889SizeGivenBeta(100)
    Me.ISO16889SizeAtBeta200 = GetISO16889SizeGivenBeta(200)
    Me.ISO16889SizeAtBeta1000 = GetISO16889SizeGivenBeta(1000)
    
    Debug.Print "Beta size calculations completed:"
    Debug.Print "  ß=2: " & Me.ISO16889SizeAtBeta2
    Debug.Print "  ß=10: " & Me.ISO16889SizeAtBeta10
    Debug.Print "  ß=75: " & Me.ISO16889SizeAtBeta75
    Debug.Print "  ß=100: " & Me.ISO16889SizeAtBeta100
    Debug.Print "  ß=200: " & Me.ISO16889SizeAtBeta200
    Debug.Print "  ß=1000: " & Me.ISO16889SizeAtBeta1000
    
    DevToolsMod.TimerEndCount "Beta Size Calculations"
    Exit Sub
    
CalculationError:
    Debug.Print "Error calculating beta sizes: " & Err.Description
    ' Set default values on error
    Me.ISO16889SizeAtBeta2 = ""
    Me.ISO16889SizeAtBeta10 = ""
    Me.ISO16889SizeAtBeta75 = ""
    Me.ISO16889SizeAtBeta100 = ""
    Me.ISO16889SizeAtBeta200 = ""
    Me.ISO16889SizeAtBeta1000 = ""
End Sub

'Returns the beta from the selected sensors according to the named range calculation.
Private Function GetISO16889SizeGivenBeta(Beta As Double) As String
    Dim Sizes As Variant
    Dim betaArray As Variant
    Dim avgBetas() As Double
    Dim i As Long, j As Long
    Dim sum As Double, count As Long
    Dim resultSize As Double
    
    On Error GoTo ErrorHandler
    
    ' ENHANCED: Get current sensor selection from SaveData table (not just class property)
    Dim currentSensor As String
    currentSensor = ISO16889Mod.GetISO16889SaveResult(8) ' ID 8 is sensor selection
    
    ' Fallback to class property if SaveData table is empty
    If currentSensor = "" Or currentSensor = "Not Found" Then
        currentSensor = Me.TerminationSizePhrase
    End If
    
    ' Default to LB if still empty
    If currentSensor = "" Then currentSensor = "LB"
    
    Debug.Print "GetISO16889SizeGivenBeta using sensor: " & currentSensor
    
    ' Get data directly from class object based on selected sensor
    Select Case UCase(currentSensor)
        Case "LS"
            Sizes = DataFileMod.TestData.LS_Sizes
            betaArray = Me.Beta_LS
        Case "LBE"
            Sizes = DataFileMod.TestData.LBE_Sizes
            betaArray = Me.Beta_LBE
        Case Else ' "LB" or default
            Sizes = DataFileMod.TestData.LB_Sizes
            betaArray = Me.Beta_LB
    End Select
    
    ' Validate we have data
    If IsEmpty(Sizes) Or IsEmpty(betaArray) Then
        Debug.Print "Empty arrays in GetISO16889SizeGivenBeta for sensor: " & currentSensor
        GetISO16889SizeGivenBeta = ""
        Exit Function
    End If
    
    ' Calculate average beta for each size across ALL available time intervals
    ReDim avgBetas(1 To UBound(Sizes))
    
    For i = 1 To UBound(Sizes)
        sum = 0: count = 0
        For j = 1 To UBound(betaArray, 1)
            If j <= UBound(betaArray, 1) And i <= UBound(betaArray, 2) Then
                Dim betaValue As Variant ' Changed to Variant to handle Empty values
                betaValue = betaArray(j, i)
                
                ' ENHANCED: Handle Empty values from short tests
                If Not IsEmpty(betaValue) And IsNumeric(betaValue) And CDbl(betaValue) > 0 Then
                    sum = sum + CDbl(betaValue)
                    count = count + 1
                End If
            End If
        Next j
        
        ' Calculate average
        If count > 0 Then
            avgBetas(i) = sum / count
        Else
            avgBetas(i) = 0  ' Mark as no data
        End If
    Next i
    
    ' Debug output showing ALL beta progression
    Debug.Print "Average betas for " & currentSensor & " (target ß=" & Beta & "):"
    For i = 1 To UBound(avgBetas)
        If avgBetas(i) > 0 Then
            If avgBetas(i) >= 100000 Then
                Debug.Print "  Size " & Sizes(i) & ": ß=100000 (capped high)"
            Else
                Debug.Print "  Size " & Sizes(i) & ": ß=" & Format(avgBetas(i), "0.00")
            End If
        End If
    Next i
    
    ' Find the actual data range (all sizes with any beta data)
    Dim firstValidSize As Long, lastValidSize As Long
    firstValidSize = 0
    lastValidSize = 0
    
    For i = 1 To UBound(avgBetas)
        If avgBetas(i) > 0 Then
            If firstValidSize = 0 Then firstValidSize = i
            lastValidSize = i
        End If
    Next i
    
    If firstValidSize = 0 Then
        Debug.Print "No valid beta data found for sensor: " & currentSensor
        GetISO16889SizeGivenBeta = ""
        Exit Function
    End If
    
    Debug.Print "Data range: Size " & Sizes(firstValidSize) & " to " & Sizes(lastValidSize)
    
    ' Check for off-scale conditions using actual min/max in the data
    Dim minBeta As Double, maxBeta As Double
    minBeta = 999999  ' Start high
    maxBeta = 0       ' Start low
    
    For i = firstValidSize To lastValidSize
        If avgBetas(i) > 0 Then
            If avgBetas(i) < minBeta Then minBeta = avgBetas(i)
            If avgBetas(i) > maxBeta Then maxBeta = avgBetas(i)
        End If
    Next i
    
    Debug.Print "Beta range: " & Format(minBeta, "0.00") & " to " & IIf(maxBeta >= 100000, "100000+", Format(maxBeta, "0.00"))
    
    If Beta < minBeta Then
        GetISO16889SizeGivenBeta = "<" & Format(Sizes(firstValidSize), "0.0")
        Debug.Print "Beta " & Beta & " is off-scale low: <" & Sizes(firstValidSize)
        Exit Function
    End If
    
    If Beta > maxBeta Then
        GetISO16889SizeGivenBeta = ">" & Format(Sizes(lastValidSize), "0.0")
        Debug.Print "Beta " & Beta & " is off-scale high: >" & Sizes(lastValidSize)
        Exit Function
    End If
    
    ' Find adjacent sizes that bracket the target beta
    For i = firstValidSize To lastValidSize - 1
        If avgBetas(i) > 0 And avgBetas(i + 1) > 0 Then
            ' Check if target beta falls between these two adjacent sizes
            ' Handle both increasing and decreasing progressions
            If (avgBetas(i) <= Beta And avgBetas(i + 1) >= Beta) Or _
               (avgBetas(i) >= Beta And avgBetas(i + 1) <= Beta) Then
                
                ' Perform interpolation between adjacent sizes
                resultSize = MathMod.LinearInterpolation(Beta, avgBetas(i), avgBetas(i + 1), Sizes(i), Sizes(i + 1))
                GetISO16889SizeGivenBeta = Format(resultSize, "0.1")
                
                Debug.Print "Beta " & Beta & " interpolated between adjacent sizes:"
                Debug.Print "  Size " & Sizes(i) & " (ß=" & IIf(avgBetas(i) >= 100000, "100000", Format(avgBetas(i), "0.00")) & ") and Size " & Sizes(i + 1) & " (ß=" & IIf(avgBetas(i + 1) >= 100000, "100000", Format(avgBetas(i + 1), "0.00")) & ")"
                Debug.Print "  Result: " & Format(resultSize, "0.1") & " microns"
                Exit Function
            End If
        End If
    Next i
    
    ' If not found between adjacent sizes, look for non-adjacent bracketing
    ' This handles cases where there might be gaps in the progression
    Dim lowerBetaIdx As Long, higherBetaIdx As Long
    Dim lowerBeta As Double, higherBeta As Double
    lowerBetaIdx = 0
    higherBetaIdx = 0
    lowerBeta = 0
    higherBeta = 999999
    
    For i = firstValidSize To lastValidSize
        If avgBetas(i) > 0 Then
            ' Look for beta just below target
            If avgBetas(i) < Beta And avgBetas(i) > lowerBeta Then
                lowerBeta = avgBetas(i)
                lowerBetaIdx = i
            End If
            
            ' Look for beta just above target
            If avgBetas(i) > Beta And avgBetas(i) < higherBeta Then
                higherBeta = avgBetas(i)
                higherBetaIdx = i
            End If
        End If
    Next i
        
    If lowerBetaIdx > 0 And higherBetaIdx > 0 Then
        ' Found bracketing values (may not be adjacent)
        resultSize = MathMod.LinearInterpolation(Beta, lowerBeta, higherBeta, Sizes(lowerBetaIdx), Sizes(higherBetaIdx))
        GetISO16889SizeGivenBeta = Format(resultSize, "0.1")
        
        Debug.Print "Beta " & Beta & " interpolated between non-adjacent sizes:"
        Debug.Print "  Size " & Sizes(lowerBetaIdx) & " (ß=" & IIf(lowerBeta >= 100000, "100000", Format(lowerBeta, "0.00")) & ") and Size " & Sizes(higherBetaIdx) & " (ß=" & IIf(higherBeta >= 100000, "100000", Format(higherBeta, "0.00")) & ")"
        Debug.Print "  Result: " & Format(resultSize, "0.1") & " microns"
        Exit Function
    End If
    
    ' If we get here, something went wrong with the bracketing
    Debug.Print "ERROR: Could not find bracketing values for Beta " & Beta & " in sensor " & currentSensor
    GetISO16889SizeGivenBeta = ""
    Exit Function
    
ErrorHandler:
    GetISO16889SizeGivenBeta = ""
    Debug.Print "Error in GetISO16889SizeGivenBeta for beta " & Beta & " sensor " & currentSensor & ": " & Err.Description
End Function

'======================================================================
'================ PRIVATE HELPER METHODS ============================
'======================================================================

' Validate essential data is present
Private Function ValidateEssentialData(ByRef errorMsg As String) As Boolean
    ValidateEssentialData = True
    
    ' Check for any particle count data
    If IsEmpty(DataFileMod.TestData.LBU_CountsData) And _
       IsEmpty(DataFileMod.TestData.LSU_CountsData) And _
       IsEmpty(DataFileMod.TestData.LBE_CountsData) Then
        errorMsg = "No particle count data found in this file." & vbCrLf & _
                   "ISO 16889 analysis requires upstream and downstream particle counts."
        ValidateEssentialData = False
        Exit Function
    End If
    
    ' Check for analog data
    If IsEmpty(DataFileMod.TestData.analogData) Then
        errorMsg = "No analog data (pressure, flow, etc.) found in this file." & vbCrLf & _
                   "ISO 16889 analysis requires pressure progression data."
        ValidateEssentialData = False
        Exit Function
    End If
    
    ' Check for time data
    If IsEmpty(DataFileMod.TestData.Times) Then
        errorMsg = "No time data found in this file." & vbCrLf & _
                   "ISO 16889 analysis requires time progression data."
        ValidateEssentialData = False
        Exit Function
    End If
    
    ' Check for test setup information
    If DataFileMod.TestData.TestSetup = "" Or DataFileMod.TestData.TestSetup = "Unknown" Then
        errorMsg = "Test setup information is missing or unknown." & vbCrLf & _
                   "ISO 16889 analysis requires valid test configuration data."
        ValidateEssentialData = False
        Exit Function
    End If
End Function

' Enhanced dual filter validation
Private Function SetFilter1or2TerminationFields() As Boolean
    SetFilter1or2TerminationFields = False
    
    Dim Filter1SetP As Double
    Dim Filter2SetP As Double
    Dim Filter1Data As Variant
    Dim Filter2Data As Variant
    
    ' Validate both setpoints exist
    Filter1SetP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", 1)
    Filter2SetP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", 2)
    
    If Filter1SetP <= 0 Then
        MsgBox "Filter 1 Terminal DP setpoint not found or invalid.", vbCritical
        Exit Function
    End If
    
    If Filter2SetP = -1 Or IsEmpty(Filter2SetP) Or Filter2SetP <= 0 Then
        ' Only Filter 1 available - validate and use it
        Filter1Data = DataFileMod.TestData.GetAnalogTagData("TS_PreDPress")
        If IsEmpty(Filter1Data) Then
            MsgBox "Filter 1 pressure data (TS_PreDPress) not found.", vbCritical
            Exit Function
        End If
        
        Me.TerminationTag = "TS_PreDPress"
        Me.TerminationFilter = 1
        Me.TerminationDP = Filter1SetP
        SetFilter1or2TerminationFields = True
        Exit Function
    End If
    
    ' Validate both pressure data arrays exist
    Filter1Data = DataFileMod.TestData.GetAnalogTagData("TS_PreDPress")
    Filter2Data = DataFileMod.TestData.GetAnalogTagData("TS_FinalDPress")
    
    If IsEmpty(Filter1Data) Then
        MsgBox "Filter 1 pressure data (TS_PreDPress) not found.", vbCritical
        Exit Function
    End If
    
    If IsEmpty(Filter2Data) Then
        MsgBox "Filter 2 pressure data (TS_FinalDPress) not found.", vbCritical
        Exit Function
    End If
    
    ' Proceed with original dual filter logic
    Dim i As Long, Filter1Index As Long, Filter2Index As Long
    Filter1Index = UBound(Filter1Data)
    Filter2Index = UBound(Filter2Data)

    For i = LBound(Filter1Data) To UBound(Filter1Data)
        If Filter1Data(i) > Filter1SetP Then
            Filter1Index = i
            Exit For
        End If
    Next i

    For i = LBound(Filter2Data) To UBound(Filter2Data)
        If Filter2Data(i) > Filter2SetP Then
            Filter2Index = i
            Exit For
        End If
    Next i

    If Filter1Index <= Filter2Index Then
        Me.TerminationTag = "TS_PreDPress"
        Me.TerminationFilter = 1
        Me.TerminationDP = Filter1SetP
    Else
        Me.TerminationTag = "TS_FinalDPress"
        Me.TerminationFilter = 2
        Me.TerminationDP = Filter2SetP
    End If
    
    SetFilter1or2TerminationFields = True
End Function

' Get the actual termination DP that occurred during the test (ORIGINAL value)
Public Function GetActualTerminationDP() As Double
    ' IMPORTANT: Always get the ORIGINAL termination DP from From Data column
    ' This ensures we always validate against the actual test termination, not user overrides
    Dim originalDP As String
    originalDP = GetISO16889SaveResult(2) ' From Data column for DP
    
    If IsNumeric(originalDP) And originalDP <> "" Then
        GetActualTerminationDP = CDbl(originalDP)
    Else
        ' Fallback to header data if From Data not available
        GetActualTerminationDP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", Me.TerminationFilter)
    End If
End Function



' Check if dual filter pressure data is available
Private Function HasDualFilterData() As Boolean
    Dim filter2DP As Variant
    filter2DP = TableMod.GetValueFromTable("HeaderData", "General Test Information", "TerminalDP", 2)
    
    ' Valid if not empty and not -1 (placeholder for missing)
    HasDualFilterData = (Not IsEmpty(filter2DP) And filter2DP <> -1)
End Function

' Clear any partial data on failure
Private Sub ClearISO16889Data()
    On Error Resume Next
    
    ' Clear the ISO16889Data sheet
    Sheets("ISO16889Data").UsedRange.Clear
    
    ' Clear cache
    analysisCache.IsValid = False
End Sub

' Helper function for pressure calculation (needs to be accessible to SetClumpPressures)
Private Function GetISO16889Pressure(DP As Double, Percentage As Double) As Double
    Dim DPData As Variant
    Dim Times As Variant
    Dim timePt As Double
    
    Times = DataFileMod.TestData.Times
    DPData = DataFileMod.TestData.GetAnalogTagData(Me.TerminationTag)
    timePt = Times(UBound(Times)) * Percentage
    
    GetISO16889Pressure = MathMod.LinInterpolation(timePt, Times, DPData, 1)
End Function

