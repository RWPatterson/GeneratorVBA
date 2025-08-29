Attribute VB_Name = "DataFileMod"
Option Explicit
' Modern DataFileMod - Optimized for Performance
' Clean architecture without legacy compatibility

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

Public Sub ProcessDataFile()
    DevToolsMod.TimerStartCount
    DevToolsMod.OptimizePerformance True
    
    On Error GoTo CleanExit
    
    ' Initialize data object and validate
    If Not InitializeTestData() Then Exit Sub
    
    ' Build worksheet cache for ultra-fast lookups
    BuildWorksheetCache "RawData"
    
    ' Process in optimized order for best performance
    ProcessHeaderData
    ExtractTestConfiguration
    ProcessDataArrays
    CalculateTimeArrays
    DeployDataToSheets
    FormatDataTables
    
CleanExit:
    DevToolsMod.OptimizePerformance False
    DevToolsMod.TimerEndCount "Total Data Processing"
End Sub

' Modified initialization - create fresh instance
Private Function InitializeTestData() As Boolean
    ' Create fresh instance (old one disposed in cleanup)
    Set TestData = New DataFileClassMod
    Set TestData.WorkbookInstance = ThisWorkbook
    
    ' Quick validation - single cell read
    TestData.DataExist = (Sheets("RawData").Cells(1, 1).Value = "HEADER")
    
    InitializeTestData = TestData.DataExist
End Function

' Build comprehensive cache for worksheet operations
Private Sub BuildWorksheetCache(wsName As String)
    DevToolsMod.TimerStartCount
    
    With wsCache
        .Name = wsName
        .IsValid = False
        
        Dim ws As Worksheet
        Set ws = Sheets(wsName)
        
        ' Use bulk Find operations for all boundaries
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
            ' Determine repeat count efficiently
            .RepeatCount = IIf(FastStringExists(ws, "MidstreamFlag") Or FastStringExists(ws, "LSSizes"), 5, 3)
            .rowCount = (.DataEnd - .DataStart + 1) \ .RepeatCount
        End If
        
        .IsValid = (.HeaderStart > 0 And .HeaderEnd > 0 And .DataStart > 0 And .DataEnd > 0)
    End With
    
    DevToolsMod.TimerEndCount "Cache Build"
End Sub

' Ultra-fast string existence check
Private Function FastStringExists(ws As Worksheet, searchStr As String) As Boolean
    FastStringExists = Not (ws.UsedRange.Find(searchStr, , , xlWhole, , , True) Is Nothing)
End Function

' Optimized header processing with bulk operations
Private Sub ProcessHeaderData()
    DevToolsMod.TimerStartCount
    
    If Not wsCache.IsValid Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    ' Single bulk read of entire header section
    Dim headerRows As Long
    headerRows = wsCache.HeaderEnd - wsCache.HeaderStart - 1
    
    Dim headerRange As Range
    Set headerRange = ws.Cells(wsCache.HeaderStart + 1, 1).Resize(headerRows, wsCache.lastCol)
    
    Dim rawData As Variant
    rawData = headerRange.Value
    
    ' Process header array in memory with optimized loops
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

' Enhanced test configuration extraction with missing value fallback handling
Private Sub ExtractTestConfiguration()
    DevToolsMod.TimerStartCount
    
    With TestData
        ' Core file information - usually present in all files
        .FileName = GetConfigValueSafe("General Test Information", "FileName", "Unknown File")
        .FileDate = GetConfigValueSafe("General Test Information", "TestDate", Date)
        .TestStartTime = GetConfigValueSafe("General Test Information", "TestTime", Time)
        .testType = GetConfigValueSafe("General Test Information", "TestType", "Unknown Test Type")
        
        ' Particle counter settings - may be missing in older files
        .CountTime = GetConfigValueSafeInt("Particle Counter Configuration", "CountTime", 60)  ' Default 60 seconds
        .HoldTime = GetConfigValueSafeInt("Particle Counter Configuration", "HoldTime", 0)    ' Default 10 seconds
        
        ' System configuration - often missing in legacy files
        .MidstreamFlag = GetConfigValueSafeBool("Dilution System Configuration", "MidstreamFlag", False)
        .PressureSource = GetConfigValueSafe("Dilution System Configuration", "PressureSource", False)
        .AuxPressureFlag = GetConfigValueSafeBool("Test System Configuration", "AuxPressureFlag", False)
        .TestSetup = GetConfigValueSafe("Test System Configuration", "Setup", "Spin On")
    End With
    
    DevToolsMod.TimerEndCount "Test Configuration Extraction"
End Sub

' Safe config value retrieval with fallback defaults
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
    ' Optional: Log missing values for debugging
    Debug.Print "Config value missing - Section: " & section & ", Key: " & key & ", Using default: " & defaultValue
End Function

' Safe integer config value retrieval
Private Function GetConfigValueSafeInt(section As String, key As String, defaultValue As Long) As Long
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetConfigValueSafe(section, key, defaultValue)
    
    ' Validate that we can convert to integer
    If IsNumeric(result) Then
        Dim tempInt As Long
        tempInt = CLng(result)
        
        ' Sanity check for reasonable values
        If tempInt >= 0 And tempInt <= 86400 Then  ' 0 to 24 hours in seconds
            GetConfigValueSafeInt = tempInt
            Exit Function
        End If
    End If
    
UseDefault:
    GetConfigValueSafeInt = defaultValue
    Debug.Print "Invalid integer config - Section: " & section & ", Key: " & key & ", Using default: " & defaultValue
End Function

' Safe boolean config value retrieval
Private Function GetConfigValueSafeBool(section As String, key As String, defaultValue As Boolean) As Boolean
    On Error GoTo UseDefault
    
    Dim result As Variant
    result = GetConfigValueSafe(section, key, IIf(defaultValue, "True", "False"))
    
    ' Use the enhanced boolean conversion from MathMod
    GetConfigValueSafeBool = MathMod.ConvertToBool(result)
    Exit Function
    
UseDefault:
    GetConfigValueSafeBool = defaultValue
    Debug.Print "Invalid boolean config - Section: " & section & ", Key: " & key & ", Using default: " & defaultValue
End Function

' Enhanced version of the original GetConfigValue for backward compatibility
Private Function GetConfigValue(section As String, key As String) As Variant
    ' This now uses the safe version with Empty as default to maintain original behavior
    GetConfigValue = GetConfigValueSafe(section, key, Empty)
End Function

' Modern array processing with performance focus
Private Sub ProcessDataArrays()
    DevToolsMod.TimerStartCount
    
    ' Extract all tag arrays in single pass
    ExtractAllTagArrays
    
    ' Process main data based on repeat count
    Select Case wsCache.RepeatCount
        Case 3: Process3RowData
        Case 5: Process5RowData
    End Select
    
    ' Handle cycle data if present
    If TestData.CycleDataExist Then
        ProcessCycleData
    End If
    
    DevToolsMod.TimerEndCount "Data Array Processing"
End Sub

' Bulk tag array extraction
Private Sub ExtractAllTagArrays()
    ' Process each tag array individually to avoid ByRef issues
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
    Dim ws As Worksheet
    Set ws = Sheets(sheetName)
    
    ' Fast section location
    Dim sectionRange As Range
    Set sectionRange = ws.UsedRange.Find(section, , , xlPart, , , True)
    If sectionRange Is Nothing Then
        ExtractTagArray = Empty
        Exit Function
    End If
    
    ' Search within section bounds
    Dim searchRange As Range
    Set searchRange = ws.Range(sectionRange, ws.Cells(ws.UsedRange.Rows.count, 1))
    
    Dim keyRange As Range
    Set keyRange = searchRange.Find(key, , , xlWhole, , , True)
    If keyRange Is Nothing Then
        ExtractTagArray = Empty
        Exit Function
    End If
    
    ' Find data extent efficiently
    Dim lastCol As Long
    lastCol = ws.Cells(keyRange.Row, ws.Columns.count).End(xlToLeft).Column
    
    If lastCol > keyRange.Column Then
        Dim dataRange As Range
        Set dataRange = ws.Range(keyRange.offset(0, 1), ws.Cells(keyRange.Row, lastCol))
        ExtractTagArray = ConvertToArray(dataRange.Value)
    Else
        ExtractTagArray = Empty
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

' Remove the dynamic assignment function since we're using explicit assignments now

' Ultra-fast 3-row data processing
Private Sub Process3RowData()
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    Dim analogCols As Long, pcCols As Long
    analogCols = UBound(TestData.AnalogTags) + 1
    pcCols = UBound(TestData.Sizes) + 1
    
    ' Single massive read operation
    Dim totalRows As Long
    totalRows = wsCache.rowCount * 3
    
    Dim allData As Variant
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols).Value
    
    ' Pre-allocate all result arrays
    Dim analogData As Variant, lbuData As Variant, lbdData As Variant
    ReDim analogData(1 To wsCache.rowCount, 1 To analogCols)
    ReDim lbuData(1 To wsCache.rowCount, 1 To pcCols)
    ReDim lbdData(1 To wsCache.rowCount, 1 To pcCols)
    
    ' Optimized extraction with minimal calculations
    Dim i As Long, j As Long, baseRow As Long
    For i = 1 To wsCache.rowCount
        baseRow = (i - 1) * 3
        
        ' Extract analog data (row 1)
        For j = 1 To analogCols
            analogData(i, j) = allData(baseRow + 1, j)
        Next j
        
        ' Extract particle data (rows 2-3)
        For j = 1 To pcCols
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
    Dim ws As Worksheet
    Set ws = Sheets(wsCache.Name)
    
    Dim analogCols As Long, pcCols As Long
    analogCols = UBound(TestData.AnalogTags) + 1
    pcCols = UBound(TestData.LB_Sizes) + 1
    
    ' Bulk read entire data section
    Dim totalRows As Long
    totalRows = wsCache.rowCount * 5
    
    Dim allData As Variant
    allData = ws.Cells(wsCache.DataStart, 1).Resize(totalRows, pcCols).Value
    
    ' Determine 5-row variant (LBLS vs LBLB)
    Dim hasLSSizes As Boolean
    hasLSSizes = FastStringExists(ws, "LSSizes")
    
    If hasLSSizes Then
        Process5RowLBLS allData, analogCols, pcCols
    Else
        Process5RowLBLB allData, analogCols, pcCols
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
    Dim ws As Worksheet
    Set ws = Sheets("RawCycleData")
    
    Dim endRow As Long
    endRow = ws.UsedRange.Find("ENDDATA", , , xlWhole).Row
    
    Dim rowCount As Long, colCount As Long
    rowCount = endRow - 2
    colCount = UBound(TestData.CycleAnalogTags) + 1
    
    TestData.CycleDataRowCount = rowCount
    TestData.cycleAnalogData = ws.Cells(2, 1).Resize(rowCount, colCount).Value
End Sub

' Optimized time array calculations
Private Sub CalculateTimeArrays()
    DevToolsMod.TimerStartCount
    
    If TestData.DataRowCount > 0 Then
        TestData.Times = CalculateElapsedTimes(TestData.analogData, TestData.DataRowCount)
    End If
    
    If TestData.CycleDataExist And TestData.CycleDataRowCount > 0 Then
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
    
    Call TableMod.TimeArrayToDataSheets(TestData.Times, "A2")
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

' Fast sheet existence check
Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = (Sheets(sheetName).Name = sheetName)
    On Error GoTo 0
End Function

' Optimized count table formatting
Private Sub FormatCountTable(wsName As String)
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        With tbl
            ' Time format for first column
            .ListColumns(1).DataBodyRange.NumberFormat = "[h]:mm:ss.00"
            
            ' Number format for remaining columns in single operation
            If .ListColumns.count > 1 Then
                Dim numRange As Range
                Set numRange = .DataBodyRange.Resize(, .ListColumns.count - 1).offset(, 1)
                numRange.NumberFormat = "0.00"
            End If
        End With
    Next tbl
End Sub

' Optimized analog table formatting
Private Sub FormatAnalogTable(wsName As String)
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

Public Sub DebugNamedRanges()
    Debug.Print "=== NAMED RANGE INVESTIGATION ==="
    
    On Error Resume Next
    
    ' Check Selected_Sensor_Sizes
    Dim range1 As Range
    Set range1 = ThisWorkbook.Names("Selected_Sensor_Sizes").RefersToRange
    If Err.Number = 0 And Not range1 Is Nothing Then
        Debug.Print "Selected_Sensor_Sizes: " & range1.Address & " on " & range1.Worksheet.Name
        Debug.Print "Has values: " & (Not IsEmpty(range1.Value))
    Else
        Debug.Print "ERROR: Selected_Sensor_Sizes not found or invalid: " & Err.Description
        Err.Clear
    End If
    
    ' Check Selected16889BetasAverages
    Dim range2 As Range
    Set range2 = ThisWorkbook.Names("Selected16889BetasAverages").RefersToRange
    If Err.Number = 0 And Not range2 Is Nothing Then
        Debug.Print "Selected16889BetasAverages: " & range2.Address & " on " & range2.Worksheet.Name
        Debug.Print "Has values: " & (Not IsEmpty(range2.Value))
    Else
        Debug.Print "ERROR: Selected16889BetasAverages not found or invalid: " & Err.Description
        Err.Clear
    End If
    
    ' Check if ISO16889Data sheet has tables
    Dim ws As Worksheet
    Set ws = Sheets("ISO16889Data")
    Debug.Print "ISO16889Data sheet has " & ws.ListObjects.count & " tables"
    
    If ws.ListObjects.count > 0 Then
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            Debug.Print "Table: " & tbl.Name & " with " & tbl.DataBodyRange.Rows.count & " data rows"
        Next tbl
    End If
    
    On Error GoTo 0
    Debug.Print "=== END NAMED RANGE INVESTIGATION ==="
End Sub
