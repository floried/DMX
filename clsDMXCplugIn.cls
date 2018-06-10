VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDMXCPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************
' This is a plugin game for DMXControl 2.9
' - interface class module -
' You are allowed to use and adapt this code for your own plugin
' Marten Jahn, Febr 2007
' www.dmxcontrol.de
'****************************************************************************

Option Explicit

Implements IDMXCPlugin
Implements IDMXCModule

Public Property Get PluginName() As String
 PluginName = "Flaschendrehen"
End Property


Private Property Get IDMXCModule_ModuleInfo() As DMXCTypeLib.Type_Module_Info
 Set IDMXCModule_ModuleInfo = New Type_Module_Info: With IDMXCModule_ModuleInfo
  .Guid = "{b61b4c42-17ed-11da-93c3-00e08161165f}"  'adapt this line with unique GUID
                            'get a GUID from http://www.guidgen.com/Index.aspx
  .Name = "Flaschendrehen"                   'adapt this line
  .Description = "Flaschendrehen"       'adapt this line
  .Author = "Marten Jahn"                       'adapt this line
  .Experimental = True

  .ProvidesToolWindow = True
  .Category = CAT_Control
  .FormName = "frmFlaschendrehen" 'adapt this line - Form.Name geht nicht, würde Form laden
  Set .Frm = frmFlaschendrehen

  .ProvidesCommand = False
  .ProvidesSceneType = False

  .UsesProvidedOutput = True
  .UsesMessaging = True
  With .Messages
   .Add MSG_Channel_Set
   .Add MSG_Channel_Send
  End With
 End With
End Property

Private Function IDMXCPlugin_Configure() As Boolean
' if special plugin configuration is required call a procedure that creates configuration dialoge
 IDMXCPlugin_Configure = True
 Call frmFlaschendrehen.readTutorLLConfigurationData
End Function

Private Function IDMXCPlugin_Disable() As Boolean
 IDMXCPlugin_Disable = False
End Function

Private Function IDMXCPlugin_Enable() As Boolean
 IDMXCPlugin_Enable = True
End Function


Private Function IDMXCPlugin_Init(CallingAppTitle As String, MyFilename As String, DMXCPluginHelper As DMXCTypeLib.IPluginHelper) As Boolean
 
 Set mHelper = DMXCPluginHelper
 IDMXCPlugin_Init = True
 
End Function

Private Function IDMXCPlugin_MessageHandle(Msg As DMXCTypeLib.MSGs, Arg1 As Long, Arg2 As Long) As Long
 Select Case Msg
  Case MSG_Channel_Send
   'If gblnReceiveMode Then
    'MsgBox ("received  ch  " & Arg1 & " value " & Arg2)
    'TutorLLGetDMXvalue Arg1, Arg2
   'End If
 End Select
End Function

Private Property Get IDMXCPlugin_PluginName() As String
 IDMXCPlugin_PluginName = PluginName
End Property

Private Function IDMXCPlugin_Term() As Boolean
'operation for termination the plugin
 IDMXCPlugin_Term = True
End Function
