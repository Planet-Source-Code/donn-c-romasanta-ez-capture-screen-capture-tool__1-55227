VERSION 5.00
Begin VB.Form frmCaptureRectangle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2265
   ClientLeft      =   3975
   ClientTop       =   5400
   ClientWidth     =   2685
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   DrawWidth       =   2
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   179
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCaptureRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDown As Boolean
Private nOldX As Integer
Private nOldY As Integer
Dim XStart, YStart As Single
Dim XPrevious, YPrevious As Single
Dim CopyWidth, CopyHeight As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Private Sub Form_Activate()
    With Me
        .Left = -2
        .Top = -2
        .Width = Screen.Width + 2
        .Height = Screen.Height + 2
        .DrawStyle = 2
    End With
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- This where we set the Begainning of the Box
'--- that will be Drawn around the Capture Area
    If Button = 1 Then
        XStart = X
        YStart = Y
        XPrevious = XStart
        YPrevious = YStart
        frmCaptureRectangle.AutoRedraw = False
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- Where we Draw the Box around the Choosen Area as you hold down the Left Mouse
'--- button and Drag in any direction to create a rectangle

    If Button <> 1 Then Exit Sub
        frmCaptureRectangle.Line (XStart, YStart)-(XPrevious, YPrevious), , B
        frmCaptureRectangle.Refresh
        frmCaptureRectangle.Line (XStart, YStart)-(X, Y), , B
        XPrevious = X
        YPrevious = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X1 As Single, Y1 As Single
Dim CopyWidth As Single, CopyHeight As Single
Dim PictWidth As Single, PictHeight As Single

frmCaptureRectangle.Line (XStart, YStart)-(XPrevious, YPrevious), , B
frmCaptureRectangle.Refresh
If X > XStart Then X1 = XStart Else X1 = X
If Y > YStart Then Y1 = YStart Else Y1 = Y
CopyWidth = Abs(X - XStart)
CopyHeight = Abs(Y - YStart)

frmPreview.Picture = CaptureWindow(frmCaptureRectangle.hwnd, False, X1, Y1, Abs(X - XStart), Abs(Y - YStart))
PictWidth = frmPreview.ScaleX(frmPreview.Picture.Width, vbHiMetric, vbTwips)
PictHeight = frmPreview.ScaleX(frmPreview.Picture.Height, vbHiMetric, vbTwips)
frmPreview.Move 0, 0, PictWidth, PictHeight

frmPreview.Show
DoEvents
sndPlaySound App.Path & "/camera.wav", SND_ASYNC
frmMain.WindowState = vbNormal
Unload Me
End Sub


