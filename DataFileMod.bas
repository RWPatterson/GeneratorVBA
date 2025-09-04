Attribute VB_Name = "DataFileMod"
Option Explicit
' Modern DataFileMod - Enhanced Object Management
' Robust TestData lifecycle management with automatic recovery

Public TestData As DataFileClassMod

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

' Main entry point for ensuring TestData is ready for use
Public Function EnsureTestDataReady() As Boolean
    DevToolsMod.TimerStartCount
    
    ' STEP 1: Check if object exists and is valid
    If IsTestDataObjectValid() Then
        DevToolsMod.TimerEndCount "TestData Ready (existing valid object)"
        EnsureTestDataReady = TestData.DataExist
        Exit Function
    End If
    
    ' STEP 2: Object missing or invalid - attempt to create/recover
    If Not CreateOrRecoverTestDataObject() Then
        DevToolsMod.TimerEndCount "TestData Ready (failed - no object)"
        EnsureTestDataReady = False
        Exit Function
    End If
    
    ' STEP 3: Check if RawData sheet has data waiting to be processed
    Dim hasRawData As Boolean
    hasRawData = SheetExists("RawData") And (Sheets("RawData").Cells(1, 1).Value = "HEADER")
    
    If hasRawData And Not TestData.DataExist Then
        ' Raw data exists but TestData not populated - process it
        Debug.Print "EnsureTestDataReady: Found unprocessed data, processing now..."
        Call ProcessDataFile
    ElseIf Not hasRawData Then
        ' No raw data available
        TestData.DataExist = False
        Debug.Print "EnsureTestDataReady: No raw data available"
    End If
    
    DevToolsMod.TimerEndCount "TestData Ready (processed)"
    EnsureTestDataReady = TestData.DataExist
End Function

' Main data file processing - enhanced with better error handling
Public Sub ProcessDataFile()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    ' Ensure we have a valid object before processing
    If Not EnsureTestDataObject() Then
        Debug.Print "ProcessDataFile: Failed to create TestData object"
        GoTo CleanExit
    End If
    
    ' Validate we have data to process
    If Not ValidateRawDataExists() Then
        Debug.Print "ProcessDataFile: No valid raw data to process"
        GoTo CleanExit
    End If
    
    ' Build cache and process data
    BuildWorksheetCache "RawData"
    
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

'======================================================================
'================ OBJECT LIFECYCLE MANAGEMENT =======================
'======================================================================

' Check if current TestData object is valid and ready for use
Private Function IsTestDataObjectValid() As Boolean
    On Error GoTo ObjectInvalid
    
    IsTestDataObjectValid = False
    
    ' Check if object exists
    If TestData Is Nothing Then Exit Function
    
    ' Check if object is the correct type
    If TypeName(TestData) <> "DataFileClassMod" Then Exit Function
    
    ' Check if WorkbookInstance is set
    If TestData.WorkbookInstance Is Nothing Then Exit Function
    
    ' Check if WorkbookInstance points to current workbook
    If TestData.WorkbookInstance.Name <> ThisWorkbook.Name Then Exit Function
    
    ' Object appears valid
    IsTestDataObjectValid = True
    Exit Function
    
ObjectInvalid:
    IsTestDataObjectValid = False
End Function

' Create new TestData object or recover existing one
Private Function CreateOrRecoverTestDataObject() As Boolean
    On Error GoTo CreateFailed
    
    CreateOrRecoverTestDataObject = False
    
    ' Clean up any invalid existing object
    If Not TestData Is Nothing Then
        If Not IsTestDataObjectValid() Then
            Set TestData = Nothing
        End If
    End If
    
    ' Create new object if needed
    If TestData Is Nothing Then
        Set TestData = New DataFileClassMod
        Debug.Print "CreateOrRecoverTestDataObject: Created new TestData object"
    End If
    
    ' Initialize object properties
    Set TestData.WorkbookInstance = ThisWorkbook
    TestData.DataExist = False
    TestData.CycleDataExist = (SheetExists("RawCycleData") And Sheets("RawCycleData").Cells(1, 1).Value = ";Data Format:")
    
    CreateOrRecoverTestDataObject = True
    Exit Function
    
CreateFailed:
    Debug.Print "CreateOrRecoverTestDataObject Error: " & Err.Description
    CreateOrRecoverTestDataObject = False
End Function

' Ensure TestData object exists (simplified version for internal use)
Private Function EnsureTestDataObject() As Boolean
    If IsTestDataObjectValid() Then
        EnsureTestDataObject = True
    Else
        EnsureTestDataObject = CreateOrRecoverTestDataObject()
    End If
End Function

' Validate that raw data exists and is processable
Private Function ValidateRawDataExists() As Boolean
    ValidateRawDataExists = False
    
    ' Check if RawData sheet exists
    If Not SheetExists("RawData") Then
        Debug.Print "ValidateRawDataExists: RawData sheet does not exist"
        Exit Function
    End If
    
    ' Check if sheet has the expected header
    If Sheets("RawData").Cells(1, 1).Value <> "HEADER" Then
        Debug.Print "ValidateRawDataExists: RawData sheet does not contain expected header"
        Exit Function
    End If
    
    ' Check if sheet has any data beyond the header
    If Sheets("RawData").usedRange.Rows.count < 10 Then
        Debug.Print "ValidateRawDataExists: RawData sheet appears to have insufficient data"
        Exit Function
    End If
    
    ValidateRawDataExists = True
End Function

'======================================================================
'================ DATA PROCESSING FUNCTIONS =========================
'======================================================================

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
            Set found = ws.usedRange.Find(ranges(i), , , xlWhole, , , True)
            positions(i) = IIf(found Is Nothing, CACHE_MISS, found.Row)
        Next i
        
        .HeaderStart = positions(0)
        .HeaderEnd = positions(1)
        .DataStart = positions(2) + 2  ' Skip format row
        .DataEnd = positions(3)
        .lastCol = ws.usedRange.Columns.count
        
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
    FastStringExists = Not (ws.usedRange.Find(searchStr, , , xlWhole, , , True) Is Nothing)
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
    rawData = headerRange.Value
    
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
    Call TableMod.HeaderDataToSheet("HeaderData")
    
    DevToolsMod.TimerEndCount "Header Processing"
End Sub

' Enhanced test configuration extraction with robust error handling
Private Sub ExtractTestConfiguration()
    DevToolsMod.TimerStartCount
    
    With TestData
        .FileName = GetConfigValueSafe("General Test Information", "FileName", "Unknown File")
        .FileDate = GetConfigValueSafe("General Test Information", "TestDate", Date)
        .TestStartTime = GetConfigValueSafe("General Test Information", "TestTime", Time)
        .testType = GetConfigValueSafe("General Test Information", "TestType", "Unknown Test Type")
        .CountTime = GetConfigValueSafeInt("Particle Counter Configuration", "CountTime", 60)
        .HoldTime = GetConfigValueSafeInt("Particle Counter Configuration", "HoldTime", 0)
        .MidstreamFlag = GetConfigValueSafeBool("Dilution System Configuration", "MidstreamFlag", False)
        .PressureSource = GetConfigValueSafe("Dilution System Configuration", "PressureSource", False)
        .AuxPressureFlag = GetConfigValueSafeBool("Test System Configuration", "AuxPressureFlag", False)
        .TestSetup = GetConfigValueSafe("Test System Configuration", "Setup", "Spin On")
    End With
    
    DevToolsMod.TimerEndCount "Test Configuration Extraction"
End Sub

' Safe config value retrieval with comprehensive fallback handling
Private Function GetConfigValueSafe(section As String, key As String, defaultValue As Variant) As Variant
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetValueFromTable("HeaderData", section, key, 1)
    
    ' Check if result is empty, null, or error value
    If IsEmpty(result) Or IsNull(result) Or IsError(result) Then
        GoTo UseDefault
    End If
    
    ' Additional validation for string values
    If VarType(defaultValue) = vbString Then
        If Trim(CStr(result)) = "" Or CStr(result) = "#N/A" Or CStr(result) = "ERROR" Then
            GoTo UseDefault
        End If
    End If
    
    GetConfigValueSafe = result
    Exit Function
    
UseDefault:
    GetConfigValueSafe = defaultValue
    Debug.Print "GetConfigValueSafe: Using default for " & section & "." & key & " = " & defaultValue
End Function

' Safe integer config value retrieval
Private Function GetConfigValueSafeInt(section As String, key As String, defaultValue As Long) As Long
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetConfigValueSafe(section, key, defaultValue)
    
    If IsNumeric(result) Then
        Dim tempInt As Long
        tempInt = CLng(result)
        If tempInt >= 0 And tempInt <= 86400 Then  ' 0 to 24 hours in seconds
            GetConfigValueSafeInt = tempInt
            Exit Function
        End If
    End If
    
UseDefault:
    GetConfigValueSafeInt = defaultValue
    Debug.Print "GetConfigValueSafeInt: Using default for " & section & "." & key & " = " & defaultValue
End Function

' Safe boolean config value retrieval
Private Function GetConfigValueSafeBool(section As String, key As String, defaultValue As Boolean) As Boolean
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetConfigValueSafe(section, key, IIf(defaultValue, "True", "False"))
    
    GetConfigValueSafeBool = MathMod.ConvertToBool(result)
    Exit Function
    
UseDefault:
    GetConfigValueSafeBool = defaultValue
    Debug.Print "GetConfigValueSafeBool: Using default for " & section & "." & key & " = " & defaultValue
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
    result = ExtractTagArray(";Data Format:", ";Data Format:", "RawData")
    If Not IsEmpty(result) Then
        result = PrependArrayValue(result, "Test Time")
        TestData.AnalogTags = result
    End If
    
    ' Process CycleAnalogTags
    result = ExtractTagArray(";Data Format:", ";Data Format:", "RawCycleData")
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
    Set sectionRange = ws.usedRange.Find(section, , , xlPart, , , True)
    If sectionRange Is Nothing Then Exit Function
    
    ' Search within section bounds
    Dim searchRange As Range
    Set searchRange = ws.Range(sectionRange, ws.Cells(ws.usedRange.Rows.count, 1))
    
    Dim keyRange As Range
    Set keyRange = searchRange.Find(key, , , xlWhole, , , True)
    If keyRange Is Nothing Then Exit Function
    
    ' Find data extent efficiently
    Dim lastCol As Long
    lastCol = ws.Cells(keyRange.Row, ws.Columns.count).End(xlToLeft).Column
    
    If lastCol > keyRange.Column Then
        Dim dataRange As Range
        Set dataRange = ws.Range(keyRange.offset(0, 1), ws.Cells(keyRange.Row, lastCol))
        ExtractTagArray = ConvertToArray(dataRange.Value)
    End If
End Function

' Fast array prepending without ReDim Preserve
Private Function PrependArrayValue(arr As Variant, Value As Variant) As Variant
    Dim newSize As Long
    newSize = UBound(arr) - LBound(arr) + 2
    
    Dim result As Variant
    ReDim result(1 To newSize)
    
    result(1) = Value
    
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
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols + 1).Value
    
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
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols + 1).Value
    
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
    If Not SheetExists("RawCycleData") Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = Sheets("RawCycleData")
    
    Dim endRow As Long
    Dim endRange As Range
    Set endRange = ws.usedRange.Find("ENDDATA", , , xlWhole)
    If endRange Is Nothing Then Exit Sub
    
    endRow = endRange.Row
    
    Dim rowCount As Long, colCount As Long
    rowCount = endRow - 2
    colCount = IIf(IsEmpty(TestData.CycleAnalogTags), 0, UBound(TestData.CycleAnalogTags))
    
    If rowCount > 0 And colCount > 0 Then
        TestData.CycleDataRowCount = rowCount
        TestData.cycleAnalogData = ws.Cells(2, 1).Resize(rowCount, colCount + 1).Value
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

' High-performance table formatting
Private Sub FormatDataTables()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    ' Format all count data tables
    Dim countSheets As Variant
    countSheets = Array("LBU_CountsData", "LBD_CountsData", "LBE_CountsData", "LSU_CountsData", "LSD_CountsData")
    
    Dim i As Long
    For i = 0 To UBound(countSheets)
        If SheetExists(CStr(countSheets(i))) Then
            FormatCountTable CStr(countSheets(i))
        End If
    Next i
    
    ' Format analog data tables
    If SheetExists("AnalogData") Then FormatAnalogTable "AnalogData"
    If SheetExists("CycleAnalogData") Then FormatAnalogTable "CycleAnalogData"
    
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Table Formatting"
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

' Optimized count table formatting
Private Sub FormatCountTable(wsName As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        With tbl
            ' Time format for first column
            If .ListColumns.count >= 1 Then
                .ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
            End If
            
            ' Number format for remaining columns in single operation
            If .ListColumns.count > 1 Then
                Dim numRange As Range
                Set numRange = .DataBodyRange.Resize(, .ListColumns.count - 1).offset(, 1)
                numRange.NumberFormat = "0.00"
            End If
        End With
    Next tbl
    
    On Error GoTo 0
End Sub

' Optimized analog table formatting
Private Sub FormatAnalogTable(wsName As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        With tbl
            ' Time format for first two columns
            If .ListColumns.count >= 1 Then .ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
            If .ListColumns.count >= 2 Then .ListColumns(2).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
            
            ' Number format for remaining columns
            If .ListColumns.count > 2 Then
                Dim numRange As Range
                Set numRange = .DataBodyRange.Resize(, .ListColumns.count - 2).offset(, 2)
                numRange.NumberFormat = "0.00"
            End If
        End With
    Next tbl
    
    On Error GoTo 0
End Sub

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
        GetValueFromTable = col.DataBodyRange(valueIndex, 1).Value
    Else
        GetValueFromTable = Empty
    End If
    
    Exit Function
    
ErrorHandler:
    GetValueFromTable = Empty
End Function

'======================================================================
'================ PUBLIC CLEANUP AND DIAGNOSTIC =====================
'======================================================================

' Public function to check TestData status
Public Function GetTestDataStatus() As String
    Dim status As String
    
    If TestData Is Nothing Then
        status = "TestData object is Nothing"
    ElseIf Not IsTestDataObjectValid() Then
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
    
    ' Clean up existing object
    Set TestData = Nothing
    
    ' Force recreation through EnsureTestDataReady
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
