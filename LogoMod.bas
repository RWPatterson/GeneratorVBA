Attribute VB_Name = "LogoMod"
Option Explicit

' List the sheet names and corresponding image control names here
Private LogoSheets As Variant
Private LogoC
Public ControlName As String

Public Sub InitLogoSheetList()
    ' List all sheet names where the logo image should be placed
    LogoSheets = Array( _
        "Home", "ISO_16889_Page_1", "ISO_16889_Page_2", "ISO_16889_Page_3", "C1_DP_v_Mass", "C2_Beta_v_Size", "C3_Beta_v_Time", "C4_Beta_v_Press", "C_Up_Counts", "C_Down_Counts")
    ControlName = "Image1"
End Sub

' === 1. Open Logo and Preview it ===

Public Sub OpenLogo()
    Dim logoPath As Variant
    Dim ws As Worksheet

    logoPath = Application.GetOpenFilename("Image Files (*.bmp; *.jpg; *.jpeg; *.gif; *.png), *.bmp; *.jpg; *.jpeg; *.gif; *.png", , "Select Logo Image")

    If logoPath = False Then Exit Sub

    Range("Logo_Path").Value = logoPath
    Set ws = ThisWorkbook.Sheets("Home")

    On Error Resume Next
    With ws.OLEObjects("LogoPreview").Object
        .Picture = LoadPicture(logoPath)
        .PictureSizeMode = 3 ' 3 = Zoom
    End With
    On Error GoTo 0
End Sub


' === 2. Remove Logo from All Sheets ===

Public Sub RemoveLogo()
    InitLogoSheetList
    Dim i As Long
    Dim ws As Worksheet
    For i = LBound(LogoSheets) To UBound(LogoSheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(LogoSheets(i))
        ws.OLEObjects(ControlName).Object.Picture = Nothing
        On Error GoTo 0
    Next i
    Range("Logo_Path").Value = ""
    MsgBox "Logo removed from all pages.", vbInformation
End Sub

' === 3. Apply Logo to All Pages ===

Public Sub ApplyLogoToAll()
    Dim logoPath As String
    InitLogoSheetList
    logoPath = Range("Logo_Path").Value
    
    If logoPath = "" Or Dir(logoPath) = "" Then
        MsgBox "No valid logo path specified. Please use 'Open Logo' first.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    Dim ws As Worksheet
    For i = LBound(LogoSheets) To UBound(LogoSheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(LogoSheets(i))
        With ws.OLEObjects(ControlName).Object
            .Picture = LoadPicture(logoPath)
            .PictureSizeMode = 3   ' fmPictureSizeModeZoom
            .BorderStyle = 0      ' fmBorderStyleNone
            .BackStyle = 0        ' fmBackStyleTransparent
            .Width = 75
            .Height = 50
            .Left = 500
        End With
        On Error GoTo 0
    Next i
    MsgBox "Logo applied to all pages.", vbInformation
End Sub

' ===== Helper Function: (optional, use if you wish to remove custom logos) =====
Public Sub ClearLogoPreview()
    On Error Resume Next
    ThisWorkbook.Sheets("Main").OLEObjects("LogoPreview").Object.Picture = Nothing
    On Error GoTo 0
End Sub

