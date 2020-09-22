VERSION 5.00
Begin VB.Form frmPreview 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   5775
   ClientTop       =   3675
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PictWidth As Single, PictHeight As Single

Private Sub Form_Activate()
PictWidth = Me.ScaleX(Me.Picture.Width, vbHiMetric, vbTwips)
PictHeight = Me.ScaleX(Me.Picture.Height, vbHiMetric, vbTwips)
Me.Move 0, 0, PictWidth, PictHeight
End Sub


Private Sub Form_Resize()
PictWidth = Me.ScaleX(Me.Picture.Width, vbHiMetric, vbTwips)
PictHeight = Me.ScaleX(Me.Picture.Height, vbHiMetric, vbTwips)
Me.Move 0, 0, PictWidth, PictHeight
End Sub
