VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustLogo 
   Caption         =   "Set Custom Defaults"
   ClientHeight    =   4680
   ClientLeft      =   30
   ClientTop       =   195
   ClientWidth     =   5445
   OleObjectBlob   =   "frmCustLogo.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1747"
End
Attribute VB_Name = "frmCustLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdCancel_Click()
   Unload Me
End Sub

Private Sub UserForm_Activate()
    Dim Logo As String
    
    Logo = Range("Logo_Path")
    If Logo <> "" Then ImgLogo (Logo)
            
End Sub
Private Sub CmdLogo_Click()

    Dim LogoToOpen As String
    
    LogoToOpen = Application.GetOpenFilename("Image Files (*.bmp; *.jpg; *.gif), *.bmp; *.jpg; *.gif")
    ImgLogo (LogoToOpen)
    Range("Logo_Path") = LogoToOpen
    Range("Logo_New") = True
    
End Sub

Private Sub cmdSave_Click()
    
'''    Clear_Data
    ActiveWorkbook.SaveCopyAs
    Unload Me
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the close box in the title bar
    
    If CloseMode <> 1 Then
    
        If MsgBox("Do you want to save?", vbYesNo) = vbYes Then
            cmdSave_Click
        Else
        'close form
           CmdCancel_Click
        End If
    Else
        'This section is used if the code closes the form.
    End If
End Sub
    
