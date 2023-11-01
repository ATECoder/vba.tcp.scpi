VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScpiSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   SCPI System Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Session As cc_isr_Ieee488.TcpSession
    BeeperEnabled As Boolean
    LastDeviceError As String
    LastDeviceErrors As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="ScpiSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_Ieee488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="ScpiSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As ScpiSystem
    Set This.Session = a_session
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_Ieee488.TcpSession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.Session Is Nothing Then
        Connected = False
    Else
        Connected = This.Session.Connected
    End If

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  I/O
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the last set or get Beeper enabled condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get BeeperEnabled() As Boolean
    BeeperEnabled = This.BeeperEnabled
End Property

''' <summary>   Gets the Beeper enabled for continuit testing. </summary>
''' <returns>   [Boolean]. </returns>
Public Function BeeperEnabledGetter() As Boolean
    Dim p_value As String
    p_value = This.Session.QueryLine(":SYST:BEEP:STAT?")
    This.BeeperEnabled = IIf(p_value = "ON", True, False)
    BeeperEnabledGetter = This.BeeperEnabled
End Function

''' <summary>   Sets the Beeper enabled. </summary>
''' <param name="a_value">   [Boolean] The Beeper enabled to set. </param >
Public Sub BeeperEnabledSetter(ByVal a_value As Boolean)
    This.Session.WriteLine ":SYST:BEEP:STAT " & IIf(a_value, "ON", "OFF")
    This.BeeperEnabled = a_value
End Sub

''' <summary>   Query INPUTS switch (0=rear, 1=front). </summary>
''' <remarks>   Prepending the command with the <c>*OPC</> reduced misreadings for the
''' inputs switch state of the Keithley 2700 multimeter scanner instrument. </remarks>
''' <param name="a_delay">   [Optional, Integer, 0] the millisecond duration to add to the
'''                          <see cref="ReadAfterWriteDelay"/>. </param>
''' <value>   [Boolean] true if the inputs are set to the front panel. </value>
Public Function QueryFrontSwitch(Optional ByVal a_delay As Integer = 0) As Boolean
    QueryFrontSwitch = "1" = This.Session.QueryLine("*OPC;:SYST:FRSW?", a_delay)
End Function

''' <summary>   Return to :SYST:PRES defaults. </summary>
Public Sub Preset()
    This.Session.WriteLine ":SYST:PRES"
End Sub


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Class events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +


''' <summary>   Initializes the class properties. </summary>
Private Sub Class_Initialize()
    This.LastDeviceError = VBA.vbNullString
    This.LastDeviceErrors = VBA.vbNullString
End Sub
