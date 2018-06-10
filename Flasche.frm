VERSION 5.00
Begin VB.Form frmFlaschendrehen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flaschendrehen Ultimate Edition"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Start_button 
      BackColor       =   &H000080FF&
      Caption         =   "Start"
      Height          =   495
      Left            =   1560
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.Timer timer_rotation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   720
   End
   Begin VB.Image Lamps_picture 
      Height          =   1020
      Index           =   4
      Left            =   6120
      Picture         =   "Flasche.frx":0000
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Image Lamps_picture 
      Height          =   1020
      Index           =   3
      Left            =   4800
      Picture         =   "Flasche.frx":039F
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Image Lamps_picture 
      Height          =   1020
      Index           =   2
      Left            =   4800
      Picture         =   "Flasche.frx":07C5
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Image Lamps_picture 
      Height          =   1020
      Index           =   1
      Left            =   4800
      Picture         =   "Flasche.frx":0BEB
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Image Lamps_picture 
      Height          =   1020
      Index           =   0
      Left            =   4800
      Picture         =   "Flasche.frx":1011
      Top             =   360
      Width           =   1185
   End
   Begin VB.Image imgPar 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image imgPar 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image imgPar 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image imgPar 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   0
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Flaschendrehen"
      BeginProperty Font 
         Name            =   "Nasalization"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFlaschendrehen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
' This is a plugin demonstration for DMXControl 2.9
' - forms module -
' You are allowed to use and adapt this code for your own plugin
' Marten Jahn, Jan 2007
' www.dmxcontrol.de
'****************************************************************************

Option Explicit

Implements IDMXCTool

' Private WithEvents Storage As IDataStorage

Dim cl As Long       'current selected lamp number
Dim push_time As Double 'duration pushing startbutton -> impacts game duration and roundtrip time
Dim game_duration As Double
Dim current_round_trip As Double 'time for whole roundtrip (4 lamps)
Dim consumed_time As Double
Dim mDimmerAddr As Integer


Private Function IDMXCTool_AskToSaveProjectData() As Boolean
  'should the plugin be asked to save data or configuration when
  'the project is changed.
  'Debug.Print "[frmTutorMaim] IDMXCTool_AskToSaveProjectData()"
  IDMXCTool_AskToSaveProjectData = False
End Function

Private Sub IDMXCTool_ClearProjectData()
  'msgbox ("[frmTutorMaim] IDMXCTool_ClearProjectData()")
End Sub

Private Sub IDMXCTool_LoadProjectData()
  'MsgBox ("[frmBallon] IDMXCTool_LoadProjectData()")
End Sub

Private Sub IDMXCTool_SaveProjectData()
  'MsgBox ("[frmBallon] IDMXCTool_SaveProjectData()")
End Sub

Private Property Let IDMXCTool_ViewMode(ByVal RHS As DMXCTypeLib.View)
  'Debug.Print "[frmBallon] IDMXCTool_ViewMode"
  'myViewmode = RHS
End Property



Public Sub lightOn(Nb As Long)
'switchs light with number Nb on
'activate picture

    imgPar(Nb).Picture = Lamps_picture(Nb).Picture

'Select Case Nb
'Case 0
'imgPar(Nb).Picture = LoadPicture("Yellow.gif")
'Case 1
'imgPar(Nb).Picture = LoadPicture("Blue.gif")
'Case 2
'imgPar(Nb).Picture = LoadPicture("Green.gif")
'Case 3
'imgPar(Nb).Picture = LoadPicture("Red.gif")

'End Select

    gLED_matrix(Nb) = LED_on
    switchDMX (Nb)
End Sub

Public Sub lightOff(Nb As Long)
'switchs light with number Nb off
'deactivate picture
    imgPar(Nb).Picture = Lamps_picture(4).Picture
    'imgPar(Nb).Picture = LoadPicture("Off.gif")
    gLED_matrix(Nb) = LED_off
    switchDMX (Nb)
End Sub


Public Sub Form_Load()
cl = 0
lightOn (cl)
    
End Sub
Public Sub readTutorLLConfigurationData()
' read only configuration data - start address
' no special form is required for this demo plugin
Dim AddressText As String
Dim ok As Boolean
While Not ok
AddressText = InputBox$("Start Address:", "Flaschendrehen Channel Configuration")
    ok = True
    'make plausibility test
    If IsNumeric(AddressText) = False Then
      MsgBox Prompt:="Please enter an integer for start address!"
      ok = False
    End If
     If AddressText > 508 Then
      MsgBox Prompt:="Please insert a value lower than 508"
      ok = False
    End If
    
    '*********************************
    ' Hier müsste Abfrage der Kanaleigenschaften erfolgen
    ' z.B. Prüfung, ob die 4 Kanäle txtStartAddr ... txtStartAddr+3 dimmbar sind
    ' nur bei positivem Test darf plugin bzw. Startbutton aktiviert werden
    '*********************************
Wend
        mDimmerAddr = Int(AddressText)
        If mDimmerAddr = 0 Then
         mDimmerAddr = 1
        End If
         
End Sub



Private Sub Start_button_Click()
'check system time
game_duration = Time

End Sub

Private Sub Start_button_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check system time
'imgPar(2) = LoadPicture("Blau_an.gif")
push_time = Timer

End Sub


Private Sub Start_button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'imgPar(2) = LoadPicture("aus.gif")
cl = 0
push_time = Timer - push_time
game_duration = push_time * 10
current_round_trip = 0.5 * push_time
timer_rotation.Interval = Int(100 * current_round_trip)
timer_rotation.Enabled = True
consumed_time = 0

End Sub



Private Sub timer_rotation_Timer()

'switchs last lamp off
'switchs next lamp on
'computes next roundtrip time

lightOff (cl)
cl = cl + 1
consumed_time = consumed_time + current_round_trip / 4
If cl = 4 Then
    current_round_trip = current_round_trip + Rnd()
    cl = 0
End If
lightOn (cl)
If (game_duration - consumed_time) > current_round_trip / 4 Then
    timer_rotation.Interval = Int(100 * current_round_trip)
Else
    timer_rotation.Enabled = False
    
End If

End Sub


Private Sub cmdHelp_Click()
'displays the help form
   Dim oF As Form
    ' Check, whether form is alrady displayed
    For Each oF In Forms
        If TypeOf oF Is frmAbout Then
           If oF.Caption = "Hilfe Flaschendrehen" Then
              oF.Show
              oF.SetFocus
              Exit Sub
           End If
        End If
    Next
    ' Create new form
    Set oF = New frmAbout
    oF.Show
    oF.lblDescription.Caption = "Bei diesem Plugin handelt es sich um das Spiel Flaschendrehen. Um die Rotation der 'Flasche', welche in Form von vier verschiedenen Farbkreisen dargestellt ist, in Gang zu setzen, bedarf es eines langen Klicks auf den 'Start'-Button. Die Dauer/Geschwindigkeit der Rotation hängt von der Druckdauer des 'Start'- Buttons ab. Jeder Farbkreis ist einem DMX-Kanal und im Sinne des Spiels einem Spieler zugeordnet. Die Lampe, welche zum Schluss leuchtet, entspricht dem auserwählten Spieler."
End Sub
