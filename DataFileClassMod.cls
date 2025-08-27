VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFileClassMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************************
'****************  DataFile Details Storage  ****************************
'************************************************************************

Public WorkbookInstance As Workbook 'Sets the parent of the ClassModule to this workbook
Public DataExist As Boolean 'True when rawdata is not empty
Public CycleDataExist As Boolean 'True when RawCycleData is not empty

'All header info
Public HeaderData As Variant

'General Test Information
Public FileName As String
Public FileDate As String
Public TestStartTime As String
Public testType As String
'Particle Counter Configuration
Public CountTime As Integer
Public HoldTime As Integer
'Dilution System Configuration
Public MidstreamFlag As String
Public PressureSource As String
'Test System Configuration
Public AuxPressureFlag As String
Public TestSetup As String

'Data Table Identifiers
Public AnalogTags As Variant 'All Analog Channel Tags
Public CycleAnalogTags As Variant 'All Cyclic Analog Channel Tags
Public Times As Variant 'Times column for each record
Public CycleTimes As Variant 'Cycle times column
Public Sizes As Variant 'All LB sizes for a 3 row sort
Public LB_Sizes As Variant 'All LB sizes for a 5 row sort
Public LBE_Sizes As Variant 'Second LB set sizes for 2023-033.
Public LS_Sizes As Variant 'All LS sizes

'Data Table Values
Public analogData As Variant 'Raw Analog Data
Public cycleAnalogData As Variant 'Raw Cyclic Data
Public LBU_CountsData As Variant 'Raw LBU Data
Public LBD_CountsData As Variant 'Raw LBD Data
Public LBE_CountsData As Variant 'Raw LBE Data Extra LB sensor for 2023-033
Public LSU_CountsData As Variant 'Raw LSU Data
Public LSD_CountsData As Variant 'Raw LSD Data

'Data Characteristics
Public DataRepeatCount As Long 'How many rows in each data record
Public DataRowCount As Long 'How many rows total
Public CycleDataRowCount As Long 'How many rows of cyclic data total

'When the DataFile object is created, it will test for data existing in the rawdata sheet, then call DataFileExplorer functions to fill its data arrays.
Private Sub Class_Initialize()
    'Standard Test Data check
    If Sheets("RawData").Range("A1") = "HEADER" Then
        DataExist = True
    Else
        DataExist = False
    End If
    
    'Cycle Test Data check
    If Sheets("RawCycleData").Range("A1") = ";Data Format:" Then
        CycleDataExist = True
    Else
        CycleDataExist = False
    End If
    
End Sub

' Helper function to get the index of the AnalogTag in the AnalogTags array
Function GetAnalogTagIndex(AnalogTag As String) As Integer
    Dim i As Integer
    
    For i = LBound(AnalogTags) To UBound(AnalogTags)
        If AnalogTags(i) = AnalogTag Then
            GetAnalogTagIndex = i
            Exit Function
        End If
    Next i
    
    GetAnalogTagIndex = "-1"
End Function

' Helper function to get the analog tag data based on the index
Function GetAnalogTagData(AnalogTag As String) As Variant

    Dim TagIndex As Integer
        TagIndex = GetAnalogTagIndex(AnalogTag)

    If TagIndex > -1 Then
        Dim DataColumn As Integer
        DataColumn = TagIndex
        
        ReDim dataArray(1 To UBound(analogData, 1))
        
        Dim i As Integer
        For i = 1 To UBound(analogData, 1)
            dataArray(i) = analogData(i, DataColumn)
        Next i
        
        GetAnalogTagData = dataArray
    Else
        GetAnalogTagData = "Not Found"
    End If
End Function
