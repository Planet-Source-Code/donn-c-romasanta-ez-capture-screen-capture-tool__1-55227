VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "EZ Capture"
   ClientHeight    =   6000
   ClientLeft      =   4680
   ClientTop       =   3930
   ClientWidth     =   9840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   1429
      ButtonWidth     =   2090
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Capture Image"
            Key             =   "Capture"
            Object.ToolTipText     =   "Capture Image"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Desktop"
                  Text            =   "Desktop"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Active"
                  Text            =   "Active Window"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Area"
                  Text            =   "Rectangular Area"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Image"
            Key             =   "Open"
            Object.ToolTipText     =   "Open a new Image."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "SaveIt"
            Object.ToolTipText     =   "Save Image As"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print Image"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            Object.ToolTipText     =   "About EZ Capture"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   1710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "&Capture"
      Begin VB.Menu mnuFullScreen 
         Caption         =   "&Desktop"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "&Active Window"
      End
      Begin VB.Menu mnuArea 
         Caption         =   "&Rectangular Area"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Cre&dits"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Dim PictWidth As Single, PictHeight As Single

Private Sub mnuFileClose_Click()
  'unload the form
Unload Me
Unload frmAbout
Unload frmCaptureRectangle
Unload frmPrintScreen
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
Unload frmAbout
Unload frmCaptureRectangle
Unload frmPrintScreen
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState <> vbMinimized Then
    With Me
    If .Width < 6500 Then
       .Width = 6500
    End If
    If .Height < 3000 Then
       .Height = 3000
    End If
    End With
End If
PictWidth = frmPreview.ScaleX(frmPreview.Picture.Width, vbHiMetric, vbTwips)
PictHeight = frmPreview.ScaleX(frmPreview.Picture.Height, vbHiMetric, vbTwips)
frmPreview.Move 0, 0, PictWidth, PictHeight

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuActive_Click()
DoEvents
Me.Hide
DoEvents
Set frmPreview.Picture = CaptureActiveWindow()
DoEvents
sndPlaySound App.Path & "/camera.wav", SND_ASYNC
Me.Show
End Sub

Private Sub mnuArea_Click()
Me.WindowState = vbMinimized
DoEvents
Me.Hide
Set frmCaptureRectangle.Picture = CaptureScreen()
frmCaptureRectangle.Show
End Sub


Private Sub mnuFileExit_Click()
Unload Me
Unload frmAbout
Unload frmCaptureRectangle
Unload frmPrintScreen
End Sub

Private Sub mnuFullScreen_Click()
DoEvents
Me.Hide
DoEvents
sndPlaySound App.Path & "/camera.wav", SND_ASYNC
DoEvents
Set frmPreview.Picture = CaptureScreen()
frmPreview.Show
Me.Show

End Sub

Private Sub mnuOpen_Click()
With CD1
    .Filter = "GIF Files (*.gif)|*.gif|JPEG Files" & _
             "(*.jpg)|*.jpg|Bitmap Files (*.bmp)|*.bmp"
    '--- Specify default filter
    .FilterIndex = 2
    '--- set starting Path
    .InitDir = "c:\aaaaaa" 'Path1
    
    .Flags = cdlOFNExplorer
    
    '--- Show the Open Dialog
    .ShowOpen
    '--- If Canceled is Pressed
    If .FileName = "" Then Exit Sub
    '--- Load the Choosen Image to the Picture Box
    frmPreview.Picture = LoadPicture(.FileName)
End With
End Sub

Private Sub mnuSave_Click()
If frmPreview.Picture = 0 Then
    MsgBox "Please capture or load an image before saving.", vbInformation, "Nothing To Save"
Else
    '--- Set the Filters
    With CD1
        .Filter = "Bitmap Files (*.bmp)|*.bmp"
        '--- Specify default filter
        .FilterIndex = 2
        '--- Hide the "Open as read only" checkbox when saving.
        .Flags = cdlOFNHideReadOnly
        '--- Show the Open Dialog
        .ShowSave
        '--- If Canceled is Pressed
        If .FileName = "" Then Exit Sub
        '--- Save the Image
        SavePicture frmPreview.Image, .FileName
    End With
End If
End Sub

Private Sub mnuPrint_Click()
    If Picture2.Picture <> 0 Then
        frmPrintScreen.PrintBitmap Picture2.Picture
    Else
        MsgBox "Please capture or load an image before printing.", vbInformation, "Nothing To Print"
    End If
End Sub
Private Sub mnuClear_Click()
frmPreview.Picture = LoadPicture()
frmPreview.Refresh
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Capture"
            mnuFullScreen_Click
        Case "Open"
            mnuOpen_Click
        Case "SaveIt"
            mnuSave_Click
        Case "Print"
            mnuPrint_Click
        Case "About"
            frmAbout.Show
    End Select

End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
              Case "Desktop"
                   mnuFullScreen_Click
              Case "Active"
                   mnuActive_Click
              Case "Area"
                   mnuArea_Click
End Select
End Sub
