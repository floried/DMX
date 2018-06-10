VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hilfe Flaschendrehen"
   ClientHeight    =   4065
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805.736
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   885
      Left            =   240
      Picture         =   "TutorAbout.frx":0000
      ScaleHeight     =   579.425
      ScaleMode       =   0  'User
      ScaleWidth      =   579.425
      TabIndex        =   1
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -112.686
      X2              =   5112.197
      Y1              =   2401.959
      Y2              =   2401.959
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   2130
      Left            =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "DMXControl 2.9 Flaschendrehen Plugin"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.798
      Y1              =   2401.959
      Y2              =   2401.959
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0 (Marten Jahn) "
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   2790
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Info zu " & App.Title
End Sub

