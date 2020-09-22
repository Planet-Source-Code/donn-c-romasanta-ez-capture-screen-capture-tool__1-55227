VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2093
      TabIndex        =   8
      Top             =   4155
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "An easy to use screen capture tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   810
      TabIndex        =   7
      Top             =   885
      Width           =   3645
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   4200
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Developed by: Donn C. Romasanta "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1545
      Width           =   2640
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2700
      Width           =   1725
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "for Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2925
      TabIndex        =   2
      Top             =   1875
      Width           =   1755
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FreeWare Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2475
      TabIndex        =   1
      Top             =   2325
      Width           =   2220
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSplash.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   3030
      Width           =   4905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   165
      X2              =   5145
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Label lblAppName 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "EZ Capture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   795
      TabIndex        =   4
      Top             =   135
      Width           =   3285
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "EZ Capture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   765
      TabIndex        =   9
      Top             =   195
      Width           =   3285
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblProductName_Click()

End Sub

Private Sub Command1_Click()
frmAbout.Hide
End Sub

Private Sub Form_Click()
frmAbout.Hide
End Sub

