VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestInfo 
   Caption         =   "Test Information"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6600
   OleObjectBlob   =   "frmTestInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmTestInfo - Modern Implementation with SaveDataTable Integration

' Form state tracking
Private formLoaded As Boolean
Private hasUnsavedChanges As Boolean



'======================================================================
'================== FORM INITIALIZATION =============================
'======================================================================

Private Sub UserForm_Initialize()
    formLoaded = False
    hasUnsavedChanges = False
    
    ' Set initial page
    Me.MultiPage1.Value = 0
    
    formLoaded = True
End Sub

Private Sub UserForm_Activate()
    If Not formLoaded Then Exit Sub
    
    ' Fill dropdown lists
    Call FillFormDropdowns
    
    ' Load current values from tables
    Call LoadFormData
End Sub

'======================================================================
'================== DROPDOWN POPULATION =============================
'======================================================================

Private Sub FillFormDropdowns()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("Save_Data")
    
    On Error Resume Next
    
    ' Pressure unit dropdowns
    Set tbl = ws.ListObjects("TSDiffPressUnitsTable")
    If Not tbl Is Nothing Then
        Me.TerminalDP_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.CleanHousingDP_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.CleanAssemblyDP_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.BypassDP_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    ' Area unit dropdown
    Set tbl = ws.ListObjects("MediaAreaUnitsTable")
    If Not tbl Is Nothing Then
        Me.Media_Area_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    ' Pleat dimension dropdowns
    Set tbl = ws.ListObjects("MediaPleatHightUnitsTable")
    If Not tbl Is Nothing Then
        Me.Media_PleatHeight_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.Media_PleatLength_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    ' Bubble point dropdown
    Set tbl = ws.ListObjects("BubblePointUnitsTable")
    If Not tbl Is Nothing Then
        Me.BubblePoint_Unit.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    ' Counter model dropdowns
    Set tbl = ws.ListObjects("APCCounterModelsTable")
    If Not tbl Is Nothing Then
        Me.US_CounterType.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.DS_CounterType.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    ' Sensor model dropdowns
    Set tbl = ws.ListObjects("APCSensorModelsTable")
    If Not tbl Is Nothing Then
        Me.US_SensorType.List = tbl.ListColumns(2).DataBodyRange.Value
        Me.DS_SensorType.List = tbl.ListColumns(2).DataBodyRange.Value
    End If
    
    On Error GoTo 0
End Sub

'======================================================================
'================== DATA LOADING AND SAVING =========================
'======================================================================

Private Sub LoadFormData()
    Application.EnableEvents = False
    
    ' Loop through all controls and load data based on their Tag property
    Call LoadControlsFromTables(Me)
    
    ' Load controls from MultiPage pages
    Dim page As Object
    For Each page In Me.MultiPage1.Pages
        Call LoadControlsFromTables(page)
    Next page
    
    Application.EnableEvents = True
    hasUnsavedChanges = False
End Sub

Private Sub LoadControlsFromTables(containerControl As Object)
    Dim ctrl As Control
    Dim tagValue As String
    Dim ID As Long
    Dim tableType As String
    Dim controlValue As Variant
    
    For Each ctrl In containerControl.Controls
        tagValue = ctrl.Tag
        
        If tagValue <> "" And IsNumeric(tagValue) Then
            ID = CLng(tagValue)
            
            ' Determine table type based on ID
            If ID >= 1000 Then
                ' ISO16889 table (remove offset)
                tableType = "ISO16889"
                ID = ID - 1000
            ElseIf ID = 9999 Then
                ' Named range - skip for now
                GoTo NextControl
            Else
                ' SaveData table
                tableType = "SaveData"
            End If
            
            ' Get the effective value (Report Value column)
            controlValue = SaveDataMod.GetEffectiveValue(tableType, ID)
            
            ' Set the control value
            Call SetControlValue(ctrl, controlValue)
        End If
        
NextControl:
    Next ctrl
End Sub

Private Sub SaveFormData()
    Application.EnableEvents = False
    
    ' Save all controls based on their Tag property
    Call SaveControlsToTables(Me)
    
    ' Save controls from MultiPage pages
    Dim page As Object
    For Each page In Me.MultiPage1.Pages
        Call SaveControlsToTables(page)
    Next page
    
    Application.EnableEvents = True
    hasUnsavedChanges = False
    
    MsgBox "Test information saved successfully.", vbInformation
End Sub

Private Sub SaveControlsToTables(containerControl As Object)
    Dim ctrl As Control
    Dim tagValue As String
    Dim ID As Long
    Dim tableType As String
    Dim controlValue As Variant
    
    For Each ctrl In containerControl.Controls
        tagValue = ctrl.Tag
        
        If tagValue <> "" And IsNumeric(tagValue) Then
            ID = CLng(tagValue)
            controlValue = GetControlValue(ctrl)
            
            ' Determine table type based on ID
            If ID >= 1000 Then
                ' ISO16889 table (remove offset)
                tableType = "ISO16889"
                ID = ID - 1000
            ElseIf ID = 9999 Then
                ' Named range - skip for now
                GoTo NextControl
            Else
                ' SaveData table
                tableType = "SaveData"
            End If
            
            ' Save to User Entry column
            Call SaveDataMod.SetValue(tableType, ID, controlValue, SaveDataMod.UserEntry)
        End If
        
NextControl:
    Next ctrl
End Sub

'======================================================================
'================== NAVIGATION BUTTONS ==============================
'======================================================================

Private Sub NextBtn_Click()
    If (Me.MultiPage1.Value = Me.MultiPage1.Pages.count - 1) Then
       NextBtn.Visible = False
       PrevBtn.Visible = True
    Else
      If (Me.MultiPage1.Value = Me.MultiPage1.Pages.count - 2) And (frmTestInfo.MultiPage1.Pages("pg_Grph_Size").Visible = False) Then
         NextBtn.Visible = False
         PrevBtn.Visible = True
      Else
        If (Me.MultiPage1.Value = Me.MultiPage1.Pages.count - 3) And (frmTestInfo.MultiPage1.Pages("pg_Part_Size").Enabled = False) Then
           NextBtn.Visible = False
           PrevBtn.Visible = True
        Else
            NextBtn.Visible = True
            PrevBtn.Visible = True
        
            If Me.MultiPage1(Me.MultiPage1.Value + 1).Enabled Then
               Me.MultiPage1.Value = Me.MultiPage1.Value + 1
            ElseIf Me.MultiPage1(Me.MultiPage1.Value + 2).Enabled Then
               Me.MultiPage1.Value = Me.MultiPage1.Value + 2
            End If
        End If
      End If
    End If
End Sub

Private Sub PrevBtn_Click()
    If Me.MultiPage1.Value = 0 Then
       NextBtn.Visible = True
       PrevBtn.Visible = False
    Else
      PrevBtn.Visible = True
      NextBtn.Visible = True
   
       If Me.MultiPage1(Me.MultiPage1.Value - 1).Enabled Then
          Me.MultiPage1.Value = Me.MultiPage1.Value - 1
       ElseIf Me.MultiPage1(Me.MultiPage1.Value - 2).Enabled Then
          Me.MultiPage1.Value = Me.MultiPage1.Value - 2
       End If
     End If
End Sub

'======================================================================
'================== SAVE AND CLOSE BUTTONS ==========================
'======================================================================

Private Sub TI_SaveBtn_Click()
    Call SaveFormData
    
    ' Handle post-save actions (like ISO 16889 rebuilds)
    Call HandlePostSaveActions
    
    Unload Me
End Sub

Private Sub HandlePostSaveActions()
    ' Check if critical ISO 16889 values changed and rebuild if needed
    ' You can add this logic later when needed
End Sub

'======================================================================
'================== DATE VALIDATION (Keep your existing logic) =======
'======================================================================

Private Sub TestDay_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   If Me.TestDay.Value <> "" Then
     If Me.TestDay.Value < 1 Or Me.TestDay.Value > 31 Then
       MsgBox ("Please enter a valid day")
       Me.TestDay.Value = ""
       Cancel = True
     Else
       Cancel = False
     End If
   Else
      Cancel = False
   End If
End Sub

Private Sub TestMonth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
     If Me.TestMonth.Value <> "" Then
        If Me.TestMonth.Value < 1 Or Me.TestMonth.Value > 12 Then
           MsgBox ("Please enter a valid month")
           Me.TestMonth.Value = ""
           Cancel = True
           Exit Sub
        Else
           Cancel = False
        End If
        If Me.TestMonth.Value = 2 Then
           If Me.TestDay.Value > 29 Then
              MsgBox ("Please enter a valid day/month")
              Me.TestMonth.Value = ""
              Me.TestDay.Value = ""
              Cancel = True
              Exit Sub
           Else
              Cancel = False
           End If
        Else
           If Me.TestMonth.Value = 4 Or Me.TestMonth.Value = 6 Or Me.TestMonth.Value = 9 Or Me.TestMonth.Value = 11 Then
              If Me.TestDay.Value > 30 Then
                 MsgBox ("Please enter a valid day/month")
                 Me.TestMonth.Value = ""
                 Me.TestDay.Value = ""
                 Cancel = True
                 Exit Sub
              Else
                 Cancel = False
              End If
           Else
              If Me.TestDay.Value > 31 Then
                 MsgBox ("Please enter a valid day/month")
                 Me.TestMonth.Value = ""
                 Me.TestDay.Value = ""
                 Cancel = True
                 Exit Sub
              Else
                 Cancel = False
              End If
           End If
        End If
     Else
       Cancel = False
     End If
End Sub

Private Sub TestYear_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Dim YDate As Date
    YDate = Year(Date)
        
    If Me.TestYear.Value <> "" Then
        If Me.TestYear.Value < 1980 Or Me.TestYear.Value > YDate Then
           MsgBox ("Please enter a valid year: yyyy")
           Me.TestYear.Value = ""
           Cancel = True
           Exit Sub
        Else
           Cancel = False
        End If
         
         If IsDate(TestDay & "/" & TestMonth & "/" & TestYear) = True Then
           Me.TestDate.Value = DateValue(TestDay & "/" & TestMonth & "/" & TestYear)
           Cancel = False
         Else
           MsgBox ("Please enter a valid date dd/mm/yyyy")
           Me.TestYear.Value = ""
           Cancel = True
         End If
    Else
       Cancel = False
    End If
End Sub

'======================================================================
'================== UTILITY FUNCTIONS ===============================
'======================================================================

Private Function GetControlValue(ctrl As Control) As Variant
    On Error Resume Next
    
    Select Case TypeName(ctrl)
        Case "TextBox", "ComboBox"
            GetControlValue = ctrl.Value
        Case "CheckBox", "OptionButton"
            GetControlValue = ctrl.Value
        Case Else
            GetControlValue = ctrl.Value
    End Select
End Function

Private Sub SetControlValue(ctrl As Control, Value As Variant)
    On Error Resume Next
    
    Select Case TypeName(ctrl)
        Case "TextBox", "ComboBox"
            ctrl.Value = Value
        Case "CheckBox", "OptionButton"
            ctrl.Value = CBool(Value)
        Case Else
            ctrl.Value = Value
    End Select
End Sub

'======================================================================
'================== FORM CLOSE HANDLING =============================
'======================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 And hasUnsavedChanges Then
        If MsgBox("Do you want to continue without saving?", vbYesNo + vbQuestion) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

' Track changes
Private Sub MultiPage1_Change()
    If formLoaded Then hasUnsavedChanges = True
End Sub
