VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Format Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ViSession As cc_isr_Ieee488.ViSession
    Elements As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="FormatSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="FormatSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As FormatSystem
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the last set or get format elements. </summary>
''' <value>   [String]. </value>
Public Property Get Elements() As String
    Elements = This.Elements
End Property

''' <summary>   Gets the format elements. </summary>
''' <returns>   [String]. </returns>
Public Function ElementsGetter() As String
    This.Elements = This.ViSession.QueryLine("FORM:ELEM?")
    ElementsGetter = This.Elements
End Function

''' <summary>   Sets the format elements. </summary>
''' <param name="a_value">   [String] The format elements to set. </param >
Public Function ElementsSetter(ByVal a_value As String)
    This.ViSession.WriteLine ("FORM:ELEM " & a_value)
    This.Elements = a_value
End Function

