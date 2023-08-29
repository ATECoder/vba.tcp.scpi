VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SenseSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A generic Sense Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ViSession As cc_isr_Ieee488.ViSession
    SenseFunction As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="SenseSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="SenseSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As SenseSystem
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

''' <summary>   Gets the Sense Function/ </summary>
Public Property Get SenseFunction() As String
    SenseFunction = This.SenseFunction
End Property

''' <summary>   Gets the sense function. </summary>
''' <param name="a_decorator">   [String] the decorator to trigger from the function name. </param>
Public Function SenseFunctionGetter(Optional ByVal a_decorator As String = """") As String

    This.SenseFunction = This.ViSession.QueryLine(":SENS:FUNC?")
    This.SenseFunction = cc_isr_Core.StringExtensions.TrimLeft(This.SenseFunction, a_decorator)
    This.SenseFunction = cc_isr_Core.StringExtensions.TrimRight(This.SenseFunction, a_decorator)
    SenseFunctionGetter = This.SenseFunction
    
End Function

''' <summary>   Sets the sense function. </summary>
''' <param name="a_value">       [String] the sense function. </param>
''' <param name="a_decorator">   [String] the decorator to append before and after the function name. </param>
Public Sub SenseFunctionSetter(ByVal a_value As String, _
        Optional ByVal a_decorator As String = "'")

    This.SenseFunction = a_value
    This.ViSession.WriteLine ":SENS:FUNC " & a_decorator & This.SenseFunction & a_decorator

End Sub

