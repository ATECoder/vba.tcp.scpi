VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScpiSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const thisModuleName As String = "ScpiSystem"

Private Type ThisData
    ' reference to the IEEE488 session
    Session As Ieee488Session
End Type

Private This As ThisData

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the TCP Client socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.Session Is Nothing Then
        Connected = False
    Else
        Connected = This.Session.Connected
    End If

End Property

''' <summary>   Returns a reference to the IEEE488 Session. </summary>
''' <value>   [<see cref="Ieee488Session"/>]. </returns>
Public Property Get Session() As Ieee488Session
    Set Session = This.Session
End Property

''' <summary>   Returns a reference to the IEEE488 Session. </summary>
''' <param name="a_session">    [<see cref="<see cref="Ieee488Session"/>"/>] A connected session. </param>
''' <value>   [<see cref="ScpiSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As Ieee488Session) As ScpiSystem
    Set This.Session = a_session
    Set Initialize = Me
End Function

''' <summary>   Issues a beep. </summary>
Public Sub Beep()
    Me.Session.VISession.WriteLine ("SYST:BEEP")
End Sub

''' <summary>   Gets the top error from the error queue. </summary>
''' <value>   A String; the top error on the error queue. </value>
Public Function ErrorDequeue() As Boolean
    ErrorDequeue = Me.Session.VISession.QueryLine(":SYST:ERR?")
End Function

''' <summary>   Clears the error queue. </summary>
Public Sub ErrorQueueClear()
    Me.Session.VISession.WriteLine (":SYST:CLE")
End Sub

''' <summary>   Query INPUTS switch (0=rear, 1=front). </summary>
''' <value>   A boolean true if the inputs are set to the front panel. </value>
Public Function FrontSwitch() As Boolean
    FrontSwitch = 1 = Me.Session.VISession.QueryLine(":SYST:FRSW?")
End Function

''' <summary>   Return to :SYST:PRES defaults. </summary>
Public Sub Preset()
    Me.Session.VISession.WriteLine (":SYST:PRES")
End Sub

