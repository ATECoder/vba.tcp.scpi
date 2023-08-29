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
    ViSession As cc_isr_Ieee488.ViSession
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="ScpiSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="ScpiSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As ScpiSystem
    Set This.ViSession = a_session
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_IEEE488.VISession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.ViSession Is Nothing Then
        Connected = False
    Else
        Connected = This.ViSession.Connected
    End If

End Property

''' <summary>   Issues a beep. </summary>
Public Sub Beep()
    This.ViSession.WriteLine ("SYST:BEEP")
End Sub

''' <summary>   Gets the top error from the error queue. </summary>
''' <value>   A String; the top error on the error queue. </value>
Public Function ErrorDequeue() As Boolean
    ErrorDequeue = This.ViSession.QueryLine(":SYST:ERR?")
End Function

''' <summary>   Clears the error queue. </summary>
Public Sub ErrorQueueClear()
    This.ViSession.WriteLine (":SYST:CLE")
End Sub

''' <summary>   Query INPUTS switch (0=rear, 1=front). </summary>
''' <value>   A boolean true if the inputs are set to the front panel. </value>
Public Function QueryFrontSwitch() As Boolean
    QueryFrontSwitch = 1 = This.ViSession.QueryLine(":SYST:FRSW?")
End Function

''' <summary>   Return to :SYST:PRES defaults. </summary>
Public Sub Preset()
    This.ViSession.WriteLine (":SYST:PRES")
End Sub

