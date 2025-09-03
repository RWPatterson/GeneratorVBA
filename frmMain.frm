VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "ISO 16889 Report Writer (v_.yyyy-mm)"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4650
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ModGraphBtn_Click()
'   Show the Chart Form
    Application.ScreenUpdating = False
    Me.Hide
    Sheets("ISO_16889_Page_2").Calculate
    Sheets("ISO_16889_Page_3").Calculate
    'ReptOpts_Check 'check report options to determine which charts to show
    frmchart.Show
    Sheets("Dashboard").Select
    Application.ScreenUpdating = True
    Me.Show
End Sub


Private Sub cmdLoadLogoBtn_Click()
    Dim Path As String
    Dim File As String
    Dim FilePath As String
            
    FilePath = Range("Logo_Path")
    Path = ParsePath(FilePath)
    File = ParseFileName(FilePath)
    
     Me.Hide
    frmCustLogo.Show
    Me.Show
End Sub


Private Sub ReportType_lb_Click()
    
    Dim ExRaDs As Boolean
    ExRaDs = ReportFillMod.GetSaveResult(58)
    
    If ExRaDs Then
        Call ISO16889Mod.SetISO16889SaveUserEntry(7, 1)
        Call ISO16889Mod.SetISO16889SaveUserEntry(8, "LB")
    Else
        Call ISO16889Mod.SetISO16889SaveUserEntry(7, 1)
        Call ISO16889Mod.SetISO16889SaveUserEntry(8, "")
    End If

End Sub


Private Sub ReportType_ls_Click()
    Dim FilterCount As Integer
    
    'Find the number of filters
    FilterCount = ReportFillMod.GetSaveResult(7)

    'Depending on if there is more than one filter
    If FilterCount > 1 Then
    'Series filter report
        Call ISO16889Mod.SetISO16889SaveUserEntry(7, 2)
        Call ISO16889Mod.SetISO16889SaveUserEntry(8, "LS")
    Else
    'Single filter report
        Call ISO16889Mod.SetISO16889SaveUserEntry(7, 1)
        Call ISO16889Mod.SetISO16889SaveUserEntry(8, "LS")
    End If
    
End Sub

Private Sub Units_English_Click()
    Call ReportFillMod.SetSaveUserEntry(55, "ENG")
End Sub

Private Sub Units_SI_Click()
   Call ReportFillMod.SetSaveUserEntry(55, "SI")
End Sub

Private Sub SetCustomDefaults()

    Sheets("Save_Data").Calculate
    
    frmTestInfo.MultiPage1.Pages("pg_Part_Size").Enabled = False  ' hide because a data file has not been selected yet
    frmTestInfo.MultiPage1.Pages("pg_Grph_Size").Enabled = False  ' hide because a data file has not been selected yet
    frmTestInfo.MultiPage1.Pages("pg_Inj_Sys").Enabled = False    ' the Injection System Tab does not have default values

    frmTestInfo.Show

End Sub

Private Sub ReptOpts_Set()

    'Check that data in Save_Data is current
    Sheets("Save_Data").Calculate
   
    Dim ReportUnits As String
    ReportUnits = ReportFillMod.GetSaveResult(30)
    
    'Toggle buttons based on Save_Data table.
     If ReportUnits = "SI" Then
        Me.Units_SI = True
        Me.Units_English = False
     Else
        If ReportUnits = "ENG" Then
           Me.Units_SI = False
           Me.Units_English = True
        Else
            If ReportUnits = "English" Then
               Me.Units_SI = False
               Me.Units_English = True
            Else                           ' set default value to SI
               Me.Units_SI = True
               Me.Units_English = False
            End If
        End If
     End If
        
     
     Dim ReportSensor As String
     ReportSensor = ISO16889Mod.GetISO16889SaveResult(8)
     
      If ReportSensor = "LB" Then
         Me.ReportType_lb = True
         Me.ReportType_ls = False
      Else
        If ReportSensor = "LS" Then
         Me.ReportType_lb = False
         Me.ReportType_ls = True
        Else
          Me.ReportType_lb = True
          Me.ReportType_ls = False
        End If
      End If
      
End Sub


Private Sub Language_Change()
Dim i As Control
Dim r As Integer

    Call SetSaveUserEntry(94, Me.Language.ListIndex + 1)
    'Application.CalculateFullRebuild
    
'    For Each i In Me.Controls
'        If i.Name = "Language" Then
''       If (i.Name = "language") Or (i.Name = "Info") Or (i.Name = "txtTitFile") Then
'        Else
'        If (i.Tag = "9999") Or (i.Tag = "") Then
'        Else
'            r = Application.WorksheetFunction.Match(Int(i.Tag), Sheets("Translations").Range("A11:A1400"), 0)
'            i.Caption = Sheets("Translations").Cells(r + 10, frmMain.Language.ListIndex + 2)
'            'MsgBox i.Name & "  " & i.Caption
'        End If
'        End If
        
'    Next

End Sub

Private Sub UserForm_Activate()
    Dim Path As String
    Dim File As String
    Dim FilePath As String
    
    Call DataFileMod.EnsureDataFileReady
    
    'Modified to check the Save_Data table
    If DataFileMod.TestData.DataExist = True Then
        Me.FileInfo = "File Open: " & DataFileMod.TestData.FileName & " .dat"
        
        Me.AddGravimetricsbtn.Enabled = True
        Me.AddEditTestInfoBtn.Caption = "Add / Edit Test Info"
        Me.AddEditTestInfoBtn.Enabled = True
        Me.ViewPrintReportBtn.Enabled = True
        
        Me.SaveBtn.Enabled = True
        Me.SaveBtn.Caption = "Save Excel Report"
        
        Me.cmdLoadLogoBtn.Enabled = False
        
        Me.ModGraphBtn.Enabled = True
        Me.ReportType_lb.Enabled = True
        Me.ReportType_ls.Enabled = True
        
        
        Dim FilterCount As Integer
        On Error Resume Next
        FilterCount = CInt(ReportFillMod.GetSaveResult(7))
        
        
        If FilterCount > 1 Then
            Me.Frame2.Caption = "Report Filter"
            Me.ReportType_lb.Caption = "Pre-Filter"
            Me.ReportType_ls.Caption = "Final Filter"
        Else
            Me.Frame2.Caption = "Report Counters"
            Me.ReportType_lb.Caption = "Light Blocking"
            Me.ReportType_ls.Caption = "Light Scattering"
        End If
        
    Else
        Me.FileInfo = "File Open: "
        Me.AddGravimetricsbtn.Enabled = False
        
        Me.AddEditTestInfoBtn.Caption = "Add / Edit Custom Defaults"
        Me.AddEditTestInfoBtn.Enabled = True
        
        Me.ViewPrintReportBtn.Enabled = False
        Me.ReportType_lb.Enabled = False
        Me.ReportType_ls.Enabled = False
        
        Me.SaveBtn.Enabled = True
        Me.SaveBtn.Caption = "Save Report Template"
        
        Me.ModGraphBtn.Enabled = False
    End If
    
    'FilePath = Range("Logo_Path")
    'Path = ParsePath(FilePath)
    'File = ParseFileName(FilePath)
    
'    If Range("Load_Logo").Value = True Then
'        If FileExists(Path, File) = False Then
'            If MsgBox("Report Writer has detected that no logo is loaded for this application." & vbCrLf & vbTab & vbTab & "Would you like to add one now?", vbYesNo) = vbYes Then
'          '      frmCustLogo.show vbModeless
'                frmCustLogo.Show
'                Range("Load_Logo").Value = True
'            Else
'                Range("Load_Logo").Value = False
'            End If
'        End If
'    End If
    
    ReptOpts_Set
     
    'Language_Change
End Sub

Private Sub UserForm_Initialize()
    RWName = ThisWorkbook.Name
    RWPath = ThisWorkbook.Path & "\"
    
    Dim RWVersion As String
    RWVersion = Sheets("Save_Data").ListObjects("ReportWriterNameTable").DataBodyRange(2, 3)
    
    Me.Caption = RWVersion
        
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Save_Data") 'This function is hardcoded to use the Save_Data Sheet.
    Set tbl = ws.ListObjects("DisplayLanguageTable") 'The SaveDataTable is always present in the workbook.
    
    Me.Language.List = tbl.ListColumns(2).DataBodyRange.Value 'fill language combobox

    Language.ListIndex = CInt(GetSaveResult(57)) - 1
      'Language_Change
End Sub


Private Sub AddGravimetricsbtn_Click()
    Me.Hide
  '  frmGrav.show vbModeless
  '  Me.show vbModeless
    frmGrav.Show
    Me.Show
End Sub

Private Sub AddEditTestInfoBtn_Click()
    Me.Hide
    
    Call DataFileMod.EnsureDataFileReady
    
    'If data is available, change the report options as usual.
    If DataFileMod.TestData.DataExist = True Then
        frmTestInfo.Show
    Else
    'If data is not available, change the default options.
        Call SetCustomDefaults
    End If
    
    Me.Show
End Sub

Private Sub CreateReportBtn_Click()
    Dim r As Integer

    If OpenDataFile() Then
        DataFileMod.TestData.DataExist = True
        DataFileMod.SetupDataFileModule
        
        Sheets("Dashboard").Select
        
        Application.ScreenUpdating = True
        
        Me.AddGravimetricsbtn.Enabled = True
        Me.AddEditTestInfoBtn.Enabled = True
            Me.AddEditTestInfoBtn.Caption = "Add / Edit Test Info"
        Me.ViewPrintReportBtn.Enabled = True
        Me.SaveBtn.Enabled = True
        Me.cmdLoadLogoBtn.Enabled = False
        Me.ModGraphBtn.Enabled = True
        
    End If
    
End Sub


Private Sub SaveBtn_Click()
      
    If savefile() Then
           
    Else
        MsgBox "Save Failed"
    End If
End Sub



   
Private Sub QuitBtn_Click()
    'Application.Visible = True
    Unload Me
End Sub
    
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        Unload Me
End Sub


'To be checked later

Private Function EvaluateTestType() As Boolean

   Dim testType As String
   Dim TestSetup As String
   Dim Midstream As Boolean
   
        ' Initialize variables.
    testType = GetSaveResult(14)
    TestSetup = GetSaveResult(39)
'    Midstream = Sheets(cDF_Sort).Range("RD_MidstreamFlag")
        
    'Morph main form based on test type
    Select Case testType    ' Evaluate Number.
        Case "Multipass", "Singlepass"
        
        
            'Update Window Caption based on test type
            If TestSetup = "Pressure" Then
                Me.Caption = "ISO 16889 Report Writer (ver. " & Range("RW_Version_Num") & ")"
            Else
                Me.Caption = "Suction Filter Report Writer (ver. " & Range("RW_Version_Num") & ")"
            End If
            
                Me.ReportType_lb.Caption = "Light Blocking"
                Me.ReportType_ls.Caption = "Light Scattering"


            'Return True
            EvaluateTestType = True
            
        Case "Multipass Series"
            
            
            'Update Window Caption based on test setup.
            If TestSetup = "Suction Pressure" Then
                Me.Caption = "Suction Series Report Writer (ver. " & Range("RW_Version_Num") & ")"
            Else

                Me.Caption = "Multipass Series Report Writer (ver. " & Range("RW_Version_Num") & ")"
            End If
            
            
            'Enable Filter 1 or 2 if Multipass Series is on
                Me.ReportType_lb.Caption = "PreFilter"
                Me.ReportType_ls.Caption = "Final Filter"
            
            'Return True
            EvaluateTestType = True
        
        
        Case "Data Only", "PQ", "Cyclic Multipass", "Cyclic Series Multipass"
            'No not like that.
            MsgBox "Cannot report Test Type:" & testType
            EvaluateTestType = False
            
            
        Case Else
            'Definitely not like that.
            MsgBox "Unrecognized Test Type"
            EvaluateTestType = False
        
        
        End Select

End Function
