Attribute VB_Name = "DMXCTutorLL"
'****************************************************************************
' This is a plugin demonstration for DMXControl 2.9
' - DMX logic module -
' You are allowed to use and adapt this code for your own plugin
' Marten Jahn, Feb 2007
' www.dmxcontrol.de
'****************************************************************************

Option Explicit

'Right place for global variables

Public Enum LED_status
    LED_off             'LED cell is off - zero signal for this scene
    LED_on              'LED cell is on - maximum signal for this scene
End Enum

Public gLED_matrix(0 To 3) As LED_status ' storage of LED status per cell

Public mHelper                              'interface class reference to DMXControl

Public mDimmerAddr As Long                  'stores the start DMX address
Dim mcurDMXValues(1 To 4) As Integer

Public Sub switchDMX(n As Long)
Dim s As Long
    If gLED_matrix(n) = LED_on Then
        s = 255
    Else
        s = 0
    End If
    'MsgBox ("Ausgabe   " & n & "---" & s)
    Call mHelper.MyStream.UserInteraction   ' enables the plugin to send values, independent
                                            ' from the active program module
    Call mHelper.MyStream.SetChannel(n + mDimmerAddr + 1, s, True)
End Sub



Public Sub Main()

End Sub

