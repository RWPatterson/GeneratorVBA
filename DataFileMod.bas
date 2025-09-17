Attribute VB_Name = "DataFileMod"
Option Explicit
' Modern DataFileMod - Enhanced Object Management with Reduced Bloat
' Refactored to consolidate redundant functions and centralize sheet configuration

Public TestData As DataFileClassMod

' Centralized sheet name configuration - modify these constants to change all sheet references
Private Const SHEET_RAWDATA As String = "RawData"
Private Const SHEET_RAWCYCLEDATA As String = "RawCycleData"
Private Const SHEET_HEADERDATA As String = "HeaderData"
Private Const SHEET_ANALOGDATA As String = "AnalogData"
Private Const SHEET_CYCLEANALOGDATA As String = "CycleAnalogData"
Private Const SHEET_LBU_COUNTS As String = "LBU_CountsData"
Private Const SHEET_LBD_COUNTS As String = "LBD_CountsData"
Private Const SHEET_LBE_COUNTS As String = "LBE_CountsData"
Private Const SHEET_LSU_COUNTS As String = "LSU_CountsData"
Private Const SHEET_LSD_COUNTS As String = "LSD_CountsData"

' High-performance cache for worksheet operations
Private Type WorksheetCache
    Name As String
    HeaderStart As Long
    HeaderEnd As Long
    DataStart As Long
    DataEnd As Long
    lastCol As Long
    RepeatCount As Long
    rowCount As Long
    IsValid As Boolean
End Type

Private wsCache As WorksheetCache
Private Const CACHE_MISS As Long = -1

'======================================================================
'================ PUBLIC INTERFACE FUNCTIONS ========================
'======================================================================

' CONSOLIDATED: Replaces EnsureTestDataReady, IsTestDataObjectValid, CreateOrRecoverTestDataObject, EnsureTestDataObject
Public Function EnsureTestDataReady() As Boolean
    DevToolsMod.TimerStartCount
    
    ' Single comprehensive validation and recovery function
    EnsureTestDataReady = False
    
    On Error GoTo ValidationFailed
    
    ' STEP 1: Check if object exists and is valid
    If Not TestData Is Nothing Then
        ' Validate object integrity
        If TypeName(TestData) = "DataFileClassMod" And _
           Not TestData.WorkbookInstance Is Nothing And _
           TestData.WorkbookInstance.Name = ThisWorkbook.Name Then
            ' Object is valid
            DevToolsMod.TimerEndCount "TestData Ready (existing valid object)"
            EnsureTestDataReady = True
            Exit Function
        End If
        
        ' Object exists but is invalid - clean it up
        Set TestData = Nothing
        Debug.Print "EnsureTestDataReady: Cleaned up invalid TestData object"
    End If
    
    ' STEP 2: Create new object
    Set TestData = New DataFileClassMod
    Set TestData.WorkbookInstance = ThisWorkbook
    
    TestData.CycleDataExist = (SheetExists(SHEET_RAWCYCLEDATA) And Sheets(SHEET_RAWCYCLEDATA).Cells(1, 1).value = ";Data Format:")
    
    Debug.Print "EnsureTestDataReady: Created new TestData object, DataExist=" & TestData.DataExist
    
    EnsureTestDataReady = True
    DevToolsMod.TimerEndCount "TestData Ready (object created)"
    Exit Function
    
ValidationFailed:
    Debug.Print "EnsureTestDataReady Error: " & Err.Description
    EnsureTestDataReady = False
    DevToolsMod.TimerEndCount "TestData Ready (failed)"
End Function

' Function specifically for Main.GenerateReport to check if processing needed
Public Function ShouldProcessRawData() As Boolean
    Debug.Print "ShouldProcessRawData: Starting check..."
    
    ' Ensure object exists first
    If Not EnsureTestDataReady() Then
        Debug.Print "ShouldProcessRawData: No valid TestData object"
        ShouldProcessRawData = False
        Exit Function
    End If
    
    ' Check if RawData sheet has data waiting to be processed
    Dim hasRawData As Boolean
    hasRawData = False
    
    If SheetExists(SHEET_RAWDATA) Then
        Dim headerValue As Variant
        headerValue = Sheets(SHEET_RAWDATA).Cells(1, 1).value
        Debug.Print "ShouldProcessRawData: RawData A1 = '" & headerValue & "'"
        
        hasRawData = (CStr(headerValue) = "HEADER")
        
        If hasRawData Then
            Dim usedRows As Long
            usedRows = Sheets(SHEET_RAWDATA).UsedRange.Rows.count
            Debug.Print "ShouldProcessRawData: RawData has " & usedRows & " rows"
            hasRawData = (usedRows >= 10)
        End If
    Else
        Debug.Print "ShouldProcessRawData: RawData sheet does not exist"
    End If
    
    ' Check if data has been actually PROCESSED
    Dim hasProcessedData As Boolean
    hasProcessedData = False
    
    If TestData.DataExist Then
        hasProcessedData = (TestData.FileName <> "" And _
                           TestData.testType <> "" And _
                           TestData.DataRowCount > 0)
        Debug.Print "ShouldProcessRawData: hasProcessedData = " & hasProcessedData
        Debug.Print "  - FileName: '" & TestData.FileName & "'"
        Debug.Print "  - testType: '" & TestData.testType & "'"
        Debug.Print "  - DataRowCount: " & TestData.DataRowCount
    End If
    
    Debug.Print "ShouldProcessRawData: hasRawData = " & hasRawData
    Debug.Print "ShouldProcessRawData: hasProcessedData = " & hasProcessedData
    
    ShouldProcessRawData = (hasRawData And Not hasProcessedData)
    Debug.Print "ShouldProcessRawData: Final result = " & ShouldProcessRawData
End Function

' Main data file processing - enhanced with better error handling
Public Sub ProcessDataFile()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    ' Ensure we have a valid object before processing
    If Not EnsureTestDataReady() Then
        Debug.Print "ProcessDataFile: Failed to create TestData object"
        GoTo CleanExit
    End If
    
    ' Validate we have data to process
    If Not ValidateRawDataExists() Then
        Debug.Print "ProcessDataFile: No valid raw data to process"
        GoTo CleanExit
    End If
    
    ' Build cache and process data
    BuildWorksheetCache SHEET_RAWDATA
    
    If Not wsCache.IsValid Then
        Debug.Print "ProcessDataFile: Failed to build valid worksheet cache"
        GoTo CleanExit
    End If
    
    ProcessHeaderData
    ExtractTestConfiguration
    ProcessDataArrays
    CalculateTimeArrays
    DeployDataToSheets
    FormatDataTables
    
    ' Mark object as having valid data
    TestData.DataExist = True
    Debug.Print "ProcessDataFile: Successfully processed data file"
    
CleanExit:
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Total Data Processing"
    If Err.Number <> 0 Then
        Debug.Print "ProcessDataFile Error: " & Err.Description
        If Not TestData Is Nothing Then TestData.DataExist = False
    End If
End Sub

' Public function to check TestData status
Public Function GetTestDataStatus() As String
    Dim status As String
    
    If TestData Is Nothing Then
        status = "TestData object is Nothing"
    ElseIf TypeName(TestData) <> "DataFileClassMod" Or TestData.WorkbookInstance Is Nothing Then
        status = "TestData object exists but is invalid"
    ElseIf Not TestData.DataExist Then
        status = "TestData object is valid but contains no data"
    Else
        status = "TestData object is valid and contains data"
        status = status & vbCrLf & "  - File: " & TestData.FileName
        status = status & vbCrLf & "  - Test Type: " & TestData.testType
        status = status & vbCrLf & "  - Row Count: " & TestData.DataRowCount
        
        If Not IsEmpty(TestData.LB_Sizes) Then
            status = status & vbCrLf & "  - LB Sizes: " & (UBound(TestData.LB_Sizes) - LBound(TestData.LB_Sizes) + 1) & " channels"
        End If
        
        If Not IsEmpty(TestData.LS_Sizes) Then
            status = status & vbCrLf & "  - LS Sizes: " & (UBound(TestData.LS_Sizes) - LBound(TestData.LS_Sizes) + 1) & " channels"
        End If
    End If
    
    GetTestDataStatus = status
End Function

' Public function to force object recreation (for troubleshooting)
Public Sub ForceTestDataRecreation()
    Debug.Print "ForceTestDataRecreation: Disposing current object and creating new"
    Set TestData = Nothing
    Call EnsureTestDataReady
    Debug.Print "ForceTestDataRecreation: " & GetTestDataStatus()
End Sub

' Public function to validate TestData integrity
Public Function ValidateTestDataIntegrity() As Boolean
    ValidateTestDataIntegrity = False
    
    If Not EnsureTestDataReady() Then
        Debug.Print "ValidateTestDataIntegrity: EnsureTestDataReady failed"
        Exit Function
    End If
    
    If Not TestData.DataExist Then
        Debug.Print "ValidateTestDataIntegrity: No data exists"
        Exit Function
    End If
    
    ' Check essential arrays
    If IsEmpty(TestData.AnalogTags) Then
        Debug.Print "ValidateTestDataIntegrity: Missing AnalogTags"
        Exit Function
    End If
    
    If IsEmpty(TestData.analogData) Then
        Debug.Print "ValidateTestDataIntegrity: Missing analogData"
        Exit Function
    End If
    
    If IsEmpty(TestData.Times) Then
        Debug.Print "ValidateTestDataIntegrity: Missing Times"
        Exit Function
    End If
    
    If TestData.DataRowCount <= 0 Then
        Debug.Print "ValidateTestDataIntegrity: Invalid DataRowCount"
        Exit Function
    End If
    
    ' Check array dimensions match
    If UBound(TestData.Times) <> TestData.DataRowCount Then
        Debug.Print "ValidateTestDataIntegrity: Times array size mismatch"
        Exit Function
    End If
    
    If UBound(TestData.analogData, 1) <> TestData.DataRowCount Then
        Debug.Print "ValidateTestDataIntegrity: analogData row count mismatch"
        Exit Function
    End If
    
    Debug.Print "ValidateTestDataIntegrity: All checks passed"
    ValidateTestDataIntegrity = True
End Function

'======================================================================
'================ PRIVATE HELPER FUNCTIONS ===========================
'======================================================================

' Validate that raw data exists and is processable
Private Function ValidateRawDataExists() As Boolean
    ValidateRawDataExists = False
    
    If Not SheetExists(SHEET_RAWDATA) Then
        Debug.Print "ValidateRawDataExists: RawData sheet does not exist"
        Exit Function
    End If
    
    If Sheets(SHEET_RAWDATA).Cells(1, 1).value <> "HEADER" Then
        Debug.Print "ValidateRawDataExists: RawData sheet does not contain expected header"
        Exit Function
    End If
    
    If Sheets(SHEET_RAWDATA).UsedRange.Rows.count < 10 Then
        Debug.Print "ValidateRawDataExists: RawData sheet appears to have insufficient data"
        Exit Function
    End If
    
    ValidateRawDataExists = True
End Function

' Build comprehensive cache for worksheet operations
Private Sub BuildWorksheetCache(wsName As String)
    DevToolsMod.TimerStartCount
    
    With wsCache
        .Name = wsName
        .IsValid = False
        
        If Not SheetExists(wsName) Then
            Debug.Print "BuildWorksheetCache: Sheet " & wsName & " does not exist"
            Exit Sub
        End If
        
        Dim ws As Worksheet
        Set ws = Sheets(wsName)
        
        ' Find key markers using bulk search
        Dim ranges As Variant
        ranges = Array("HEADER", "ENDHEADER", "DATA", "ENDDATA")
        
        Dim positions(3) As Long
        Dim i As Long
        
        For i = 0 To 3
            Dim found As Range
            Set found = ws.UsedRange.Find(ranges(i), , , xlWhole, , , True)
            positions(i) = IIf(found Is Nothing, CACHE_MISS, found.Row)
        Next i
        
        .HeaderStart = positions(0)
        .HeaderEnd = positions(1)
        .DataStart = positions(2) + 2  ' Skip format row
        .DataEnd = positions(3)
        .lastCol = ws.UsedRange.Columns.count
        
        ' Calculate derived values
        If .DataStart > 0 And .DataEnd > 0 Then
            .RepeatCount = IIf(FastStringExists(ws, "MidstreamFlag") Or FastStringExists(ws, "LSSizes"), 5, 3)
            .rowCount = (.DataEnd - .DataStart + 1) \ .RepeatCount
        End If
        
        .IsValid = (.HeaderStart > 0 And .HeaderEnd > 0 And .DataStart > 0 And .DataEnd > 0 And .rowCount > 0)
    End With
    
    DevToolsMod.TimerEndCount "Cache Build"
    
    If wsCache.IsValid Then
        Debug.Print "BuildWorksheetCache: Valid cache built - RepeatCount=" & wsCache.RepeatCount & ", RowCount=" & wsCache.rowCount
    Else
        Debug.Print "BuildWorksheetCache: Failed to build valid cache"
    End If
End Sub

' Ultra-fast string existence check
Private Function FastStringExists(ws As Worksheet, searchStr As String) As Boolean
    On Error Resume Next
    FastStringExists = Not (ws.UsedRange.Find(searchStr, , , xlWhole, , , True) Is Nothing)
    On Error GoTo 0
End Function

' Optimized header processing with bulk operations
Private Sub ProcessHeaderData()
    DevToolsMod.TimerStartCount
    
    If Not wsCache.IsValid Then
        Debug.Print "ProcessHeaderData: Invalid cache, skipping"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    ' Calculate header data dimensions
    Dim headerRows As Long
    headerRows = wsCache.HeaderEnd - wsCache.HeaderStart - 1
    
    If headerRows <= 0 Then
        Debug.Print "ProcessHeaderData: No header rows to process"
        Exit Sub
    End If
    
    ' Read header data in bulk
    Dim headerRange As Range
    Set headerRange = ws.Cells(wsCache.HeaderStart + 1, 1).Resize(headerRows, wsCache.lastCol)
    
    Dim rawData As Variant
    rawData = headerRange.value
    
    ' Process header array in memory
    Dim resultArray As Variant
    ReDim resultArray(wsCache.lastCol - 1, headerRows - 1)
    
    Dim i As Long, j As Long
    For i = 1 To headerRows
        For j = 1 To wsCache.lastCol
            If Not IsEmpty(rawData(i, j)) Then
                resultArray(j - 1, i - 1) = rawData(i, j)
            Else
                Exit For  ' Stop at first empty cell
            End If
        Next j
    Next i
    
    TestData.HeaderData = resultArray
    Call TableMod.HeaderDataToSheet(SHEET_HEADERDATA)
    
    DevToolsMod.TimerEndCount "Header Processing"
End Sub

' Enhanced test configuration extraction with robust error handling
Private Sub ExtractTestConfiguration()
    DevToolsMod.TimerStartCount
    
    With TestData
        .FileName = GetConfigValue("General Test Information", "FileName", vbString, "Unknown File")
        .FileDate = GetConfigValue("General Test Information", "TestDate", vbDate, Date)
        .TestStartTime = GetConfigValue("General Test Information", "TestTime", vbDate, Time)
        .testType = GetConfigValue("General Test Information", "TestType", vbString, "Unknown Test Type")
        .CountTime = GetConfigValue("Particle Counter Configuration", "CountTime", vbLong, 60)
        .HoldTime = GetConfigValue("Particle Counter Configuration", "HoldTime", vbLong, 0)
        .MidstreamFlag = GetConfigValue("Dilution System Configuration", "MidstreamFlag", vbBoolean, False)
        .PressureSource = GetConfigValue("Dilution System Configuration", "PressureSource", vbString, False)
        .AuxPressureFlag = GetConfigValue("Test System Configuration", "AuxPressureFlag", vbBoolean, False)
        .TestSetup = GetConfigValue("Test System Configuration", "Setup", vbString, "Spin On")
    End With
    
    DevToolsMod.TimerEndCount "Test Configuration Extraction"
End Sub

' CONSOLIDATED: Replaces GetConfigValueSafe, GetConfigValueSafeInt, GetConfigValueSafeBool
Private Function GetConfigValue(section As String, key As String, valueType As VbVarType, defaultValue As Variant) As Variant
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetValueFromTable(SHEET_HEADERDATA, section, key, 1)
    
    ' Check if result is empty, null, or error value
    If IsEmpty(result) Or IsNull(result) Or IsError(result) Then
        GoTo UseDefault
    End If
    
    ' Type-specific validation and conversion
    Select Case valueType
        Case vbString
            If Trim(CStr(result)) = "" Or CStr(result) = "#N/A" Or CStr(result) = "ERROR" Then
                GoTo UseDefault
            End If
            GetConfigValue = CStr(result)
            
        Case vbLong, vbInteger
            If Not IsNumeric(result) Then GoTo UseDefault
            Dim tempInt As Long
            tempInt = CLng(result)
            If tempInt < 0 Or tempInt > 86400 Then GoTo UseDefault  ' Sanity check for time values
            GetConfigValue = tempInt
            
        Case vbBoolean
            GetConfigValue = MathMod.ConvertToBool(result)
            
        Case vbDate
            If IsDate(result) Then
                GetConfigValue = CDate(result)
            Else
                GoTo UseDefault
            End If
            
        Case Else
            GetConfigValue = result
    End Select
    
    Exit Function
    
UseDefault:
    GetConfigValue = defaultValue
    Debug.Print "GetConfigValue: Using default for " & section & "." & key & " = " & defaultValue
End Function

' Modern array processing with enhanced error handling
Private Sub ProcessDataArrays()
    DevToolsMod.TimerStartCount
    
    If Not wsCache.IsValid Then
        Debug.Print "ProcessDataArrays: Invalid cache, cannot process"
        Exit Sub
    End If
    
    ' Extract all tag arrays in single pass
    ExtractAllTagArrays
    
    ' Validate we got essential arrays
    If IsEmpty(TestData.AnalogTags) Then
        Debug.Print "ProcessDataArrays: Warning - No analog tags found"
    End If
    
    If IsEmpty(TestData.Sizes) Then
        Debug.Print "ProcessDataArrays: Warning - No size data found"
    End If
    
    ' Process main data based on repeat count
    Select Case wsCache.RepeatCount
        Case 3
            Process3RowData
        Case 5
            Process5RowData
        Case Else
            Debug.Print "ProcessDataArrays: Unsupported repeat count: " & wsCache.RepeatCount
    End Select
    
    ' Handle cycle data if present
    If TestData.CycleDataExist Then
        ProcessCycleData
    End If
    
    DevToolsMod.TimerEndCount "Data Array Processing"
End Sub

' Bulk tag array extraction with improved error handling
Private Sub ExtractAllTagArrays()
    Dim result As Variant
    
    ' Process AnalogTags
    result = ExtractTagArray(";Data Format:", ";Data Format:", SHEET_RAWDATA)
    If Not IsEmpty(result) Then
        result = PrependArrayValue(result, "Test Time")
        TestData.AnalogTags = result
    End If
    
    ' Process CycleAnalogTags
    result = ExtractTagArray(";Data Format:", ";Data Format:", SHEET_RAWCYCLEDATA)
    If Not IsEmpty(result) Then
        result = PrependArrayValue(result, "Test Time")
        TestData.CycleAnalogTags = result
    End If
    
    ' Process Sizes
    result = ExtractTagArray(";Particle", "Sizes", wsCache.Name)
    If Not IsEmpty(result) Then TestData.Sizes = result
    
    ' Process LB_Sizes
    result = ExtractTagArray(";Particle", "LBSizes", wsCache.Name)
    If Not IsEmpty(result) Then TestData.LB_Sizes = result
    
    ' Process LBE_Sizes
    result = ExtractTagArray(";Particle", "LBESizes", wsCache.Name)
    If Not IsEmpty(result) Then TestData.LBE_Sizes = result
    
    ' Process LS_Sizes
    result = ExtractTagArray(";Particle", "LSSizes", wsCache.Name)
    If Not IsEmpty(result) Then TestData.LS_Sizes = result
End Sub

' High-performance tag array extraction
Private Function ExtractTagArray(section As String, key As String, sheetName As String) As Variant
    ExtractTagArray = Empty
    
    If Not SheetExists(sheetName) Then Exit Function
    
    Dim ws As Worksheet
    Set ws = Sheets(sheetName)
    
    ' Fast section location
    Dim sectionRange As Range
    Set sectionRange = ws.UsedRange.Find(section, , , xlPart, , , True)
    If sectionRange Is Nothing Then Exit Function
    
    ' Search within section bounds
    Dim searchRange As Range
    Set searchRange = ws.Range(sectionRange, ws.Cells(ws.UsedRange.Rows.count, 1))
    
    Dim keyRange As Range
    Set keyRange = searchRange.Find(key, , , xlWhole, , , True)
    If keyRange Is Nothing Then Exit Function
    
    ' Find data extent efficiently
    Dim lastCol As Long
    lastCol = ws.Cells(keyRange.Row, ws.Columns.count).End(xlToLeft).Column
    
    If lastCol > keyRange.Column Then
        Dim dataRange As Range
        Set dataRange = ws.Range(keyRange.offset(0, 1), ws.Cells(keyRange.Row, lastCol))
        ExtractTagArray = ConvertToArray(dataRange.value)
    End If
End Function

' Fast array prepending without ReDim Preserve
Private Function PrependArrayValue(arr As Variant, value As Variant) As Variant
    Dim newSize As Long
    newSize = UBound(arr) - LBound(arr) + 2
    
    Dim result As Variant
    ReDim result(1 To newSize)
    
    result(1) = value
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        result(i - LBound(arr) + 2) = arr(i)
    Next i
    
    PrependArrayValue = result
End Function

' Ultra-fast 3-row data processing
Private Sub Process3RowData()
    If wsCache.rowCount <= 0 Then
        Debug.Print "Process3RowData: No rows to process"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    Dim analogCols As Long, pcCols As Long
    analogCols = IIf(IsEmpty(TestData.AnalogTags), 0, UBound(TestData.AnalogTags))
    pcCols = IIf(IsEmpty(TestData.Sizes), 0, UBound(TestData.Sizes))
    
    If analogCols = 0 Or pcCols = 0 Then
        Debug.Print "Process3RowData: Missing essential arrays"
        Exit Sub
    End If
    
    ' Single massive read operation
    Dim totalRows As Long
    totalRows = wsCache.rowCount * 3
    
    Dim allData As Variant
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols + 1).value
    
    ' Pre-allocate all result arrays
    Dim analogData As Variant, lbuData As Variant, lbdData As Variant
    ReDim analogData(1 To wsCache.rowCount, 1 To analogCols + 1)
    ReDim lbuData(1 To wsCache.rowCount, 1 To pcCols + 1)
    ReDim lbdData(1 To wsCache.rowCount, 1 To pcCols + 1)
    
    ' Optimized extraction with minimal calculations
    Dim i As Long, j As Long, baseRow As Long
    For i = 1 To wsCache.rowCount
        baseRow = (i - 1) * 3
        
        ' Extract analog data (row 1)
        For j = 1 To analogCols + 1
            analogData(i, j) = allData(baseRow + 1, j)
        Next j
        
        ' Extract particle data (rows 2-3)
        For j = 1 To pcCols + 1
            lbuData(i, j) = allData(baseRow + 2, j)
            lbdData(i, j) = allData(baseRow + 3, j)
        Next j
    Next i
    
    ' Assign to TestData
    TestData.analogData = analogData
    TestData.LBU_CountsData = lbuData
    TestData.LBD_CountsData = lbdData
    TestData.LB_Sizes = TestData.Sizes
    TestData.DataRowCount = wsCache.rowCount
End Sub

' High-performance 5-row data processing
Private Sub Process5RowData()
    If wsCache.rowCount <= 0 Then
        Debug.Print "Process5RowData: No rows to process"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    Dim analogCols As Long, pcCols As Long
    analogCols = IIf(IsEmpty(TestData.AnalogTags), 0, UBound(TestData.AnalogTags))
    pcCols = IIf(IsEmpty(TestData.LB_Sizes), 0, UBound(TestData.LB_Sizes))
    
    If analogCols = 0 Or pcCols = 0 Then
        Debug.Print "Process5RowData: Missing essential arrays"
        Exit Sub
    End If
    
    ' Bulk read entire data section
    Dim totalRows As Long
    totalRows = wsCache.rowCount * 5
    
    Dim allData As Variant
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols + 1).value
    
    ' Determine 5-row variant (LBLS vs LBLB)
    Dim hasLSSizes As Boolean
    hasLSSizes = FastStringExists(ws, "LSSizes")
    
    If hasLSSizes Then
        Process5RowLBLS allData, analogCols + 1, pcCols + 1
    Else
        Process5RowLBLB allData, analogCols + 1, pcCols + 1
    End If
    
    TestData.DataRowCount = wsCache.rowCount
End Sub

' LBLS 5-row processing
Private Sub Process5RowLBLS(allData As Variant, analogCols As Long, pcCols As Long)
    Dim analogData As Variant, lbuData As Variant, lbdData As Variant
    Dim lsuData As Variant, lsdData As Variant
    
    ReDim analogData(1 To wsCache.rowCount, 1 To analogCols)
    ReDim lbuData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lbdData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lsuData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lsdData(1 To wsCache.rowCount, 1 To pcCols)
    
    Dim i As Long, j As Long, baseRow As Long
    For i = 1 To wsCache.rowCount
        baseRow = (i - 1) * 5
        
        ' Analog data
        For j = 1 To analogCols
            analogData(i, j) = allData(baseRow + 1, j)
        Next j
        
        ' Particle count data
        For j = 1 To pcCols
            lbuData(i, j) = allData(baseRow + 2, j)
            lsuData(i, j) = allData(baseRow + 3, j)
            lbdData(i, j) = allData(baseRow + 4, j)
            lsdData(i, j) = allData(baseRow + 5, j)
        Next j
    Next i
    
    TestData.analogData = analogData
    TestData.LBU_CountsData = lbuData
    TestData.LBD_CountsData = lbdData
    TestData.LSU_CountsData = lsuData
    TestData.LSD_CountsData = lsdData
End Sub

' LBLB 5-row processing (skips LSU)
Private Sub Process5RowLBLB(allData As Variant, analogCols As Long, pcCols As Long)
    Dim analogData As Variant, lbuData As Variant, lbdData As Variant, lbeData As Variant
    
    ReDim analogData(1 To wsCache.rowCount, 1 To analogCols)
    ReDim lbuData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lbdData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lbeData(1 To wsCache.rowCount, 1 To pcCols)
    
    Dim i As Long, j As Long, baseRow As Long
    For i = 1 To wsCache.rowCount
        baseRow = (i - 1) * 5
        
        ' Analog data
        For j = 1 To analogCols
            analogData(i, j) = allData(baseRow + 1, j)
        Next j
        
        ' Particle count data (skip LSU row 3)
        For j = 1 To pcCols
            lbuData(i, j) = allData(baseRow + 2, j)
            lbdData(i, j) = allData(baseRow + 4, j)
            lbeData(i, j) = allData(baseRow + 5, j)
        Next j
    Next i
    
    TestData.analogData = analogData
    TestData.LBU_CountsData = lbuData
    TestData.LBD_CountsData = lbdData
    TestData.LBE_CountsData = lbeData
End Sub

' Streamlined cycle data processing
Private Sub ProcessCycleData()
    If Not SheetExists(SHEET_RAWCYCLEDATA) Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_RAWCYCLEDATA)
    
    Dim endRow As Long
    Dim endRange As Range
    Set endRange = ws.UsedRange.Find("ENDDATA", , , xlWhole)
    If endRange Is Nothing Then Exit Sub
    
    endRow = endRange.Row
    
    Dim rowCount As Long, colCount As Long
    rowCount = endRow - 2
    colCount = IIf(IsEmpty(TestData.CycleAnalogTags), 0, UBound(TestData.CycleAnalogTags))
    
    If rowCount > 0 And colCount > 0 Then
        TestData.CycleDataRowCount = rowCount
        TestData.cycleAnalogData = ws.Cells(2, 1).Resize(rowCount, colCount + 1).value
    End If
End Sub

' Optimized time array calculations
Private Sub CalculateTimeArrays()
    DevToolsMod.TimerStartCount
    
    If TestData.DataRowCount > 0 And Not IsEmpty(TestData.analogData) Then
        TestData.Times = CalculateElapsedTimes(TestData.analogData, TestData.DataRowCount)
    End If
    
    If TestData.CycleDataExist And TestData.CycleDataRowCount > 0 And Not IsEmpty(TestData.cycleAnalogData) Then
        TestData.CycleTimes = CalculateElapsedTimes(TestData.cycleAnalogData, TestData.CycleDataRowCount)
    End If
    
    DevToolsMod.TimerEndCount "Time Calculations"
End Sub

' Generic high-performance time calculation
Private Function CalculateElapsedTimes(dataArray As Variant, rowCount As Long) As Variant
    Dim Times As Variant
    ReDim Times(1 To rowCount)
    
    If rowCount = 0 Then
        CalculateElapsedTimes = Times
        Exit Function
    End If
    
    Dim startTime As Double, prevTime As Double, currTime As Double
    startTime = dataArray(1, 1)
    prevTime = startTime
    
    Dim i As Long
    For i = 1 To rowCount
        currTime = dataArray(i, 1)
        
        ' Handle midnight rollover
        If currTime < prevTime Then currTime = currTime + 1
        
        ' Convert to elapsed minutes
        Times(i) = (currTime - startTime) * 1440
        prevTime = currTime
    Next i
    
    CalculateElapsedTimes = Times
End Function

' Deploy data to sheets using existing TableMod
Private Sub DeployDataToSheets()
    DevToolsMod.TimerStartCount
    
    If Not IsEmpty(TestData.Times) Then
        Call TableMod.TimeArrayToDataSheets(TestData.Times, "A2")
    End If
    
    Call TableMod.DataTagsToSheets("B1")
    Call TableMod.TestDataToSheets("B2")
    Call TableMod.ConvertDataToNamedTables("A1")
    
    DevToolsMod.TimerEndCount "Data Deployment"
End Sub

' CONSOLIDATED: Replaces FormatCountTable and FormatAnalogTable
Private Sub FormatDataTables()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    ' Format all count data tables using centralized sheet constants
    Dim countSheets As Variant
    countSheets = Array(SHEET_LBU_COUNTS, SHEET_LBD_COUNTS, SHEET_LBE_COUNTS, SHEET_LSU_COUNTS, SHEET_LSD_COUNTS)
    
    Dim i As Long
    For i = 0 To UBound(countSheets)
        If SheetExists(CStr(countSheets(i))) Then
            FormatDataTable CStr(countSheets(i)), "Count"
        End If
    Next i
    
    ' Format analog data tables
    If SheetExists(SHEET_ANALOGDATA) Then FormatDataTable SHEET_ANALOGDATA, "Analog"
    If SheetExists(SHEET_CYCLEANALOGDATA) Then FormatDataTable SHEET_CYCLEANALOGDATA, "Analog"
    
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Table Formatting"
End Sub

' CONSOLIDATED: Unified table formatting function (replaces FormatCountTable and FormatAnalogTable)
Private Sub FormatDataTable(wsName As String, tableType As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        With tbl
            Select Case LCase(tableType)
                Case "count"
                    ' Time format for first column only
                    If .ListColumns.count >= 1 Then
                        .ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
                    End If
                    
                    ' Number format for remaining columns
                    If .ListColumns.count > 1 Then
                        Dim numRange As Range
                        Set numRange = .DataBodyRange.Resize(, .ListColumns.count - 1).offset(, 1)
                        numRange.NumberFormat = "0.00"
                    End If
                    
                Case "analog"
                    ' Time format for first two columns
                    If .ListColumns.count >= 1 Then .ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
                    If .ListColumns.count >= 2 Then .ListColumns(2).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
                    
                    ' Number format for remaining columns
                    If .ListColumns.count > 2 Then
                        Dim numRange2 As Range
                        Set numRange2 = .DataBodyRange.Resize(, .ListColumns.count - 2).offset(, 2)
                        numRange2.NumberFormat = "0.00"
                    End If
            End Select
        End With
    Next tbl
    
    On Error GoTo 0
End Sub

'======================================================================
'================ UTILITY AND HELPER FUNCTIONS ======================
'======================================================================

' Fast sheet existence check
Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = (Sheets(sheetName).Name = sheetName)
    On Error GoTo 0
End Function

' Utility function for 2D to 1D array conversion
Private Function ConvertToArray(rangeValue As Variant) As Variant
    If IsArray(rangeValue) Then
        ' Handle 2D array from range
        If UBound(rangeValue, 1) = 1 Then
            ' Single row - convert to 1D
            Dim result As Variant
            ReDim result(1 To UBound(rangeValue, 2))
            
            Dim i As Long
            For i = 1 To UBound(rangeValue, 2)
                result(i) = rangeValue(1, i)
            Next i
            
            ConvertToArray = result
        Else
            ConvertToArray = rangeValue
        End If
    Else
        ' Single value
        ReDim result(1 To 1)
        result(1) = rangeValue
        ConvertToArray = result
    End If
End Function

' Enhanced GetValueFromTable with fallback handling
Private Function GetValueFromTable(wkSheet As String, tblName As String, tblKey As String, valueIndex As Integer) As Variant
    On Error GoTo ErrorHandler
    
    If Not SheetExists(wkSheet) Then
        GetValueFromTable = Empty
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets(wkSheet)
    
    Dim tbl As ListObject
    Set tbl = Nothing
    
    ' Find table by name
    Dim tblObj As ListObject
    For Each tblObj In ws.ListObjects
        If InStr(1, tblObj.Name, tblName, vbTextCompare) > 0 Then
            Set tbl = tblObj
            Exit For
        End If
    Next tblObj
    
    If tbl Is Nothing Then
        GetValueFromTable = Empty
        Exit Function
    End If
    
    ' Find column by key
    Dim col As ListColumn
    Set col = Nothing
    
    For Each col In tbl.ListColumns
        If InStr(1, col.Name, tblKey, vbTextCompare) > 0 Then
            Exit For
        End If
    Next col
    
    If col Is Nothing Then
        GetValueFromTable = Empty
        Exit Function
    End If
    
    ' Get value at specified index
    If valueIndex > 0 And valueIndex <= col.DataBodyRange.Rows.count Then
        GetValueFromTable = col.DataBodyRange(valueIndex, 1).value
    Else
        GetValueFromTable = Empty
    End If
    
    Exit Function
    
ErrorHandler:
    GetValueFromTable = Empty
End Function

'======================================================================
'================ SECTION ORGANIZATION FOR FUTURE REFACTORING =======
'======================================================================
'
' This module is organized into logical sections that could be split into separate modules:
'
' SECTION 1: DATA OBJECT MANAGEMENT
' - EnsureTestDataReady() [CONSOLIDATED object lifecycle management]
' - ValidateRawDataExists()
' - GetTestDataStatus(), ForceTestDataRecreation(), ValidateTestDataIntegrity() [diagnostic functions]
'
' SECTION 2: DATA PROCESSING PIPELINE
' - ProcessDataFile() [main pipeline controller]
' - BuildWorksheetCache(), ProcessHeaderData(), ExtractTestConfiguration()
' - ProcessDataArrays(), ExtractAllTagArrays(), ExtractTagArray()
' - Process3RowData(), Process5RowData(), Process5RowLBLS(), Process5RowLBLB()
' - ProcessCycleData(), CalculateTimeArrays(), CalculateElapsedTimes()
'
' SECTION 3: DATA FORMATTING AND DEPLOYMENT
' - DeployDataToSheets()
' - FormatDataTables(), FormatDataTable() [CONSOLIDATED formatting functions]
'
' SECTION 4: UTILITY AND HELPER FUNCTIONS
' - SheetExists(), ConvertToArray(), GetValueFromTable()
' - FastStringExists(), PrependArrayValue()
' - GetConfigValue() [CONSOLIDATED configuration retrieval]
'
'======================================================================

