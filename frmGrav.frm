VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrav 
   Caption         =   "Gravimetrics"
   ClientHeight    =   7365
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   11100
   OleObjectBlob   =   "frmGrav.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1723"
End
Attribute VB_Name = "frmGrav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Private Sub UserForm_Activate()
Dim cell As Range
Dim i As Single
    
    FillGravTable
    
    Me.Tag = Me.Table.ListCount - 1

    Me.Table.ListIndex = 0
    
    UpdateInputSection
    
    Me.GM_SpecGrav = GetSaveResult(22)
End Sub

Sub FillGravTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim numOfSamples As Long
    '@Ignore MultipleDeclarations
    Dim i As Long, j As Long ' i will loop over rows, j will loop over columns
    Dim val As Variant
    
    Set ws = ThisWorkbook.Sheets("Save_Data") ' Adjust as needed
    Set tbl = ws.ListObjects("ISO16889GravTable")
    
    With Me.Table
        .Clear
        
        Set dataRange = tbl.DataBodyRange
        numOfSamples = dataRange.Rows.count
        
        For i = 1 To numOfSamples
            .AddItem
            
            For j = 2 To tbl.ListColumns.count
                val = dataRange.Cells(i, j).value
                
                ' Display empty strings if the cell is empty or Null
                If IsNull(val) Or IsEmpty(val) Then
                    val = ""
                End If
                
                .List(i - 1, j - 2) = val
            Next j
        Next i
        
        If .ListCount > 0 Then
            .ListIndex = 0
            Me.GM_SampleName = .List(0, 0)
        End If
        
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the close box in the title bar

    If CloseMode <> 1 Then

        If MsgBox("Do you want to continue without saving?", vbOKCancel) = vbCancel Then
            Cancel = 1
        Else
            'close form
        End If
    Else
        'This section is used if the code closes the form.
    End If
End Sub
'
Private Sub GM_Volume_Change()
    If AtoDbl(Me.GM_Volume.value) = 0 Or Me.GM_Volume = Empty Then Me.GM_GravLevel.value = Empty: Exit Sub  'Me.GM_Volume.Value = 0.0001  'Prevent divison by 0
    If AtoDbl(Me.GM_DirtWeight.value) = 0 Then Me.GM_GravLevel.value = Empty: Exit Sub
    Me.GM_GravLevel.value = Format(AtoDbl(Me.GM_DirtWeight.value) * 1000 * 1000 / AtoDbl(Me.GM_Volume.value), "###0.0000")
End Sub
Private Sub GM_DirtWeight_Change()
    If AtoDbl(Me.GM_Volume.value) = 0 Then Me.GM_GravLevel.value = Empty: Exit Sub
    Me.GM_GravLevel.value = Format(AtoDbl(Me.GM_DirtWeight.value) * 1000 * 1000 / AtoDbl(Me.GM_Volume.value), "###0.0000")
End Sub
Private Sub GM_CleanPad_Change()
    If Me.GM_DirtyPad = Empty Or Me.GM_CleanPad = Empty Then Me.GM_DirtWeight.value = Empty: Exit Sub
    Me.GM_DirtWeight.value = Format((AtoDbl(Me.GM_DirtyPad.value) - AtoDbl(Me.GM_CleanPad.value)), "###0.000000")
End Sub
Private Sub GM_DirtyPad_Change()
    If Me.GM_DirtyPad = Empty Or Me.GM_CleanPad = Empty Then Me.GM_DirtWeight.value = Empty: Exit Sub
    Me.GM_DirtWeight.value = Format((AtoDbl(Me.GM_DirtyPad.value) - AtoDbl(Me.GM_CleanPad.value)), "###0.000000")
End Sub

Private Sub GM_EmptyCup_Change()
    If Me.GM_FullCup.value = Empty Or Me.GM_EmptyCup = Empty Then Me.GM_Volume.value = Empty: Exit Sub
    Me.GM_Volume.value = Format((AtoDbl(Me.GM_FullCup.value) - AtoDbl(Me.GM_EmptyCup.value)) / GetSaveResult(22), "###0.0000")
End Sub

Private Sub GM_FullCup_Change()
    If Me.GM_FullCup.value = Empty Or Me.GM_EmptyCup = Empty Then Me.GM_Volume.value = Empty: Exit Sub
    Me.GM_Volume.value = Format((AtoDbl(Me.GM_FullCup.value) - AtoDbl(Me.GM_EmptyCup.value)) / GetSaveResult(22), "###0.0000")
End Sub
'
'
Private Sub SVOption_change()
    If Me.GM_FullCup.BackColor = vbWindowBackground Then
        With Me.GM_FullCup
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With

        With Me.GM_EmptyCup
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With

        With Me.GM_Volume
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With
    Else
        With Me.GM_FullCup
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With

        With Me.GM_EmptyCup
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With

        With Me.GM_Volume
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With
    End If
End Sub
'
'
'
Private Sub MAOption_Change()
    If Me.GM_CleanPad.BackColor = vbWindowBackground Then
        With Me.GM_CleanPad
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With

        With Me.GM_DirtyPad
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With

        With Me.GM_DirtWeight
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With
    Else
        With Me.GM_CleanPad
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With

        With Me.GM_DirtyPad
            .BackColor = vbWindowBackground
            .TabStop = True
            .Locked = False
        End With

        With Me.GM_DirtWeight
            .BackColor = vbButtonFace
            .TabStop = False
            .Locked = True
        End With
    End If

End Sub

Private Sub Table_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    UpdateInputSection
End Sub
'
Private Sub Table_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UpdateInputSection
End Sub

Private Sub ApplyBtn_Click()
    UpdateTable
End Sub

'Private Sub Donebtn_Click()
'    UpdateTable
'End Sub

Private Sub SaveBtn_Click()
    Application.ScreenUpdating = False
    ExportToSheet
    Application.ScreenUpdating = True
    Unload Me
End Sub


Sub UpdateTable()
    Dim Row As Long
    
    With Me.Table
        Row = .ListIndex
        
        ' Verify selection is valid
        If Row < 0 Or Row >= .ListCount Then Exit Sub

        If .List(Row, 0) <> Me.GM_SampleName Then Exit Sub

        ' Update ListBox with formatted values from the input controls
        .List(Row, 1) = Format(Me.GM_FullCup, "###0.0000")
        .List(Row, 2) = Format(Me.GM_EmptyCup, "###0.0000")
        .List(Row, 3) = Format(Me.GM_Volume, "###0.0000")
        .List(Row, 4) = Format(Me.GM_DirtyPad, "###0.0000")
        .List(Row, 5) = Format(Me.GM_CleanPad, "###0.0000")
        .List(Row, 6) = Format(Me.GM_DirtWeight, "###0.0000")
        .List(Row, 7) = Format(Me.GM_GravLevel, "###0.0000")
        
    End With
End Sub



Sub UpdateInputSection()
    Dim RowIndex As Long
    
    ' Validate ListIndex against ListCount to avoid errors
    With Me.Table
        If .ListCount = 0 Or .ListIndex < 0 Or .ListIndex >= .ListCount Then
            ' No valid selection, clear input controls
            Me.GM_FullCup = Empty
            Me.GM_EmptyCup = Empty
            Me.GM_Volume = Empty
            Me.GM_DirtyPad = Empty
            Me.GM_CleanPad = Empty
            Me.GM_DirtWeight = Empty
            Me.GM_GravLevel = Empty
            Me.GM_SampleName = ""
            Exit Sub
        End If
        
        ' Valid row selected, populate controls
        RowIndex = .ListIndex
        
        Me.GM_SampleName = .List(RowIndex, 0)
        Me.GM_FullCup = .List(RowIndex, 1)
        Me.GM_EmptyCup = .List(RowIndex, 2)
        Me.GM_Volume = .List(RowIndex, 3)
        Me.GM_DirtyPad = .List(RowIndex, 4)
        Me.GM_CleanPad = .List(RowIndex, 5)
        Me.GM_DirtWeight = .List(RowIndex, 6)
        Me.GM_GravLevel = .List(RowIndex, 7)
    End With
End Sub
'
Sub ExportToSheet(Optional Defaults As Boolean = False)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRows As Long
    Dim r As Long

    Set ws = ThisWorkbook.Sheets("Save_Data")
    Set tbl = ws.ListObjects("ISO16889GravTable")

    Application.Calculation = xlCalculationManual
    On Error GoTo CleanExit

    dataRows = tbl.DataBodyRange.Rows.count

    If Me.Table.ListCount < dataRows Then
        MsgBox "Warning: The form table has fewer rows than the Excel data table." & vbCrLf & _
               "Only updating matching rows.", vbInformation
        dataRows = Me.Table.ListCount
    End If

    For r = 1 To dataRows
        SetValueInTableObj tbl, "Bottle Initial Weight", r, Me.Table.List(r - 1, 1)
        SetValueInTableObj tbl, "Bottle Final Weight", r, Me.Table.List(r - 1, 2)
        SetValueInTableObj tbl, "Pad Initial Weight", r, Me.Table.List(r - 1, 4)
        SetValueInTableObj tbl, "Pad Final Weight", r, Me.Table.List(r - 1, 5)
    Next r

CleanExit:
    Application.Calculation = xlCalculationAutomatic
End Sub




'
'


