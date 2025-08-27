Attribute VB_Name = "File_Subs"
Option Explicit

Global RWName As String 'Name of Workbook
Global RWPath As String 'Path of Workbook
Global DFName As String 'Name of data file
Global DFPath As String 'Path of data file
 
' Main function to open a data file, clear existing data, and load new data from the opened file
Function OpenDataFile(Optional Extension As String = "dat", Optional wkSheet As String = "RawData") As Boolean
    ' Declaring necessary variables
    Dim dataSheet As Worksheet
    Dim tempArray As Variant
    Dim Location As String
    Dim ExtType As String
    Dim FileToOpen As Variant

    DevToolsMod.TimerStartCount

    ' STEP 1: COMPLETE CLEANUP BEFORE LOADING NEW FILE
    Call ISO16889Mod.CleanupBeforeNewFile

    ' Setting file type and location based on the extension provided
    Call SetFileTypeAndLocation(Extension, Location, ExtType)

    ' Getting the file to open using the location and extension type
    FileToOpen = getFileToOpen(Location, ExtType)
    ' Turning off screen updating for better performance
    Application.ScreenUpdating = False

    ' Check if a valid file is selected
    If FileToOpen <> False Then
        ' Set the worksheet where data needs to be loaded
        Set dataSheet = ThisWorkbook.Sheets(wkSheet)
        ' Note: Clearing is now handled by CleanupBeforeNewFile

        ' Open the file and get the dataSet
        tempArray = OpenTextAndGetData(FileToOpen)

        ' Get the name and path of the current active workbook (the opened file)
        Dim DFName As String
        Dim DFPath As String
        DFName = ActiveWorkbook.Name
        DFPath = ActiveWorkbook.Path

        ' Copy the data from the array to the worksheet
        dataSheet.Range("A1").Resize(UBound(tempArray, 1), UBound(tempArray, 2)).Value = tempArray

        ' Close the opened workbook without showing any alerts
        Call CloseWorkbookWithoutAlerts(DFName)

        ' Return true indicating successful operation
        OpenDataFile = True
    End If

    ' Turning the screen updating back on
    Application.ScreenUpdating = True
    DevToolsMod.TimerEndCount "File Loading with Cleanup"
End Function

' Subroutine to set file type and location based on the provided extension
Sub SetFileTypeAndLocation(ByVal Extension As String, ByRef Location As String, ByRef ExtType As String)
    ' Check if the extension is "dat"
    If Extension = "dat" Then
        ExtType = "Test Data (*.DAT),*.dat"
        Location = "DefaultDir"
    Else
        ExtType = "Report File (*.SAV),*.sav"
        Location = "SaveDir"
    End If
End Sub

' Function to open a text file and get the data
Function OpenTextAndGetData(ByVal FileToOpen As Variant) As Variant
    Dim tempArray As Variant

    ' Open the file with specified parameters
    Workbooks.OpenText FileName:=FileToOpen, Origin:=xlWindows, startRow:=1, _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=True, Semicolon:=False, Comma:=True, Space:=False, Other:=False

    ' Read the data into array
    tempArray = ActiveSheet.UsedRange.Value

    ' Remove any double quotes from the data
    Call RemoveDoubleQuotes(tempArray)

    ' Return the array with data
    OpenTextAndGetData = tempArray
End Function

' Subroutine to remove any double quotes from the data in the array
Sub RemoveDoubleQuotes(ByRef tempArray As Variant)
    '@Ignore MultipleDeclarations
    Dim i As Long, j As Long
    For i = 1 To UBound(tempArray, 1)
        For j = 1 To UBound(tempArray, 2)
            ' Replace any double quotes with nothing (removes double quotes)
            tempArray(i, j) = Replace(tempArray(i, j), """", "")
        Next j
    Next i
End Sub

' Subroutine to close a workbook without displaying any alerts
Sub CloseWorkbookWithoutAlerts(ByVal DFName As String)
    ' Turning off alerts
    Application.DisplayAlerts = False
    ' Closing the workbook
    Workbooks(DFName).Close
    ' Turning the alerts back on
    Application.DisplayAlerts = True
End Sub
Function getFileToOpen(Location As String, ExtType As String)

Dim FilePath As String

    'KB changed ReportWriter to ReportWriter16889
    'Get and set to the last path used
    FilePath = GetSetting("ReportWriter16889", "Settings", Location)
    
    If FilePath <> "" Then
        On Error Resume Next
        ChDir FilePath
        On Error GoTo 0
    End If
    
    'Ask user to Open a file
     getFileToOpen = Application.GetOpenFilename(ExtType)

End Function


'=== Save workbook as a Template (*.xltm) ===
Public Sub SaveAsTemplate()
    Dim savePath As String
    
    ' Prompt user for template save location
    savePath = Application.GetSaveAsFilename( _
        Title:="Save-As New Report Template", _
        FileFilter:="Excel Template (*.xltm), *.xltm")
        
    ' Exit if Cancel pressed
    If savePath = "False" Then Exit Sub
    
   
    ' Save file in macro-enabled template format
    ThisWorkbook.SaveAs FileName:=savePath, _
                        FileFormat:=xlOpenXMLTemplateMacroEnabled
End Sub


'=== Save workbook as a Report (*.xlsm) ===
Public Sub SaveAsReport()
    Dim savePath As String
    
    ' Prompt user for report save location
    savePath = Application.GetSaveAsFilename( _
        Title:="Save-As Excel Report", _
        FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm")
        
    ' Exit if Cancel pressed
    If savePath = "False" Then Exit Sub
    
 
    ' Save file in macro-enabled workbook format
    ThisWorkbook.SaveAs FileName:=savePath, _
                        FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub



Function ParsePath(ByVal ParseString As String) As String
Dim index As Integer
Dim position As Integer

    index = 1
    Do
        position = index
        index = InStr(position + 1, ParseString, "\")
    Loop While index > 0
    ParsePath = Left(ParseString, position)
    
End Function


Function ParseFileName(ByVal ParseString As String, Optional NoExt As Boolean) As String

    Dim index As Integer
    Dim position As Integer
    Dim TempString As String
    Dim Delimiter As String
    
    #If Win32 Then
        Delimiter = "\"
    #ElseIf Mac Then
        Delimiter = ":"
    #End If
    
    index = 1
    Do
        position = index
        index = InStr(position + 1, ParseString$, Delimiter)
    Loop While index > 0
    If position > 1 Then
        TempString = Right(ParseString, Len(ParseString) - position)
    Else
        TempString = ParseString
    End If
    If NoExt = True And Len(TempString) > 4 Then
    
        TempString = Left(TempString, Len(TempString) - 4)
    End If
    
    ParseFileName = TempString
    
End Function

'AtoDbl - Convert ASCII string to DOUBLE
'         Returns 0 if not a valid string
Function AtoDbl(inString As String) As Double

    If IsNumeric(inString) = True Then
        AtoDbl = CDbl(inString)
    Else
        AtoDbl = 0
    End If
    
End Function

Function FileExists(Path, Fname) As Boolean
    On Error GoTo skip

            If Dir(Path & Fname) <> "" And Fname <> "" Then
                FileExists = True
            Else
                FileExists = False
            End If

    On Error GoTo 0
    Exit Function
    
skip:
FileExists = False
On Error GoTo 0
End Function


