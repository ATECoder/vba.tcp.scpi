VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriggerSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Trigger Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ViSession As cc_isr_Ieee488.ViSession
    Source As cc_isr_Tcp_Scpi.TriggerSource
End Type

Private This As this_

''' <summary>   Values that represent trigger sources. </summary>
Public Enum TriggerSource
    ''' <summary>   An enum constant representing the none trigger source. </summary>
    None = 0
    ''' <summary>   An enum constant representing the immediate trigger source. </summary>
    Immediate = 1
    ''' <summary>   An enum constant representing the external trigger source. </summary>
    External = 2
End Enum

''' <summary>   Returns a reference to the <see cref="TriggerSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="TriggerSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As TriggerSystem
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

''' <summary>   Gets the <see cref="TriggerSource"/>. </summary>
''' <value>   <see cref="TriggerSource"/>. </returns>
Public Property Get Source() As cc_isr_Tcp_Scpi.TriggerSource
    Source = This.Source
End Property

''' <summary>   Sets the <see cref="TriggerSource"/>. </summary>
''' <param name="a_value">   [<see cref="TriggerSource"/>] the trigger source. </param>
Public Sub SourceSetter(ByVal a_value As cc_isr_Tcp_Scpi.TriggerSource)

    If cc_isr_Tcp_Scpi.TriggerSource.Immediate = a_value Then
        This.ViSession.WriteLine ":TRIG:SOUR IMM"
    ElseIf cc_isr_Tcp_Scpi.TriggerSource.External = a_value Then
        ' for some reason the old application had the trigger source toggled
        ' if requesting an external trigger source
        ' This.ViSession.WriteLine ":TRIG:SOUR IMM"
        This.ViSession.WriteLine ":TRIG:SOUR EXT"
    End If
    This.Source = a_value
    
End Sub

''' <summary>   Gets the <see cref="TriggerSource"/>. </summary>
''' <param name="a_value">   [<see cref="TriggerSource"/>] the trigger source. </param>
Public Function SourceGetter() As cc_isr_Tcp_Scpi.TriggerSource

    Dim p_source As String
    p_source = This.ViSession.QueryLine(":TRIG:SOUR?")
    Select Case p_source
        Case "IMM"
            This.Source = TriggerSource.Immediate
        Case "EXT"
            This.Source = TriggerSource.External
        Case Else
            This.Source = TriggerSource.None
    End Select
    SourceGetter = This.Source
    
End Function

