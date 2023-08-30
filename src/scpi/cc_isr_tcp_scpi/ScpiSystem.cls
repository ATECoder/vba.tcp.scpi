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
    LastDeviceError As String
    LastDeviceErrors As String
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device errors
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   The last error that was dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastDeviceError() As String
    LastDeviceError = This.LastDeviceError
End Property

''' <summary>   The last errors that were dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastDeviceErrors() As String
    LastDeviceErrors = This.LastDeviceErrors
End Property

''' <summary>   Dequeues the top device error from the device error queue. </summary>
''' <value>   [String] the top error on the error queue. </value>
Public Function DequeueDeviceError() As String
    This.LastDeviceError = This.ViSession.QueryLine(":SYST:ERR?")
    DequeueDeviceError = This.LastDeviceError
End Function

''' <summary>   Clears the error queue. </summary>
Public Sub ClearDeviceErrors()
    This.ViSession.WriteLine (":SYST:CLE")
End Sub

''' <summary>   Queries and tries to parse a device error into a number and a message. </summary>
''' <param name="a_number">      [String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Public Function TryDequeueParseDeviceError(ByRef a_number As String, _
    ByRef a_message As String, Optional ByVal a_decorator As String = """") As Boolean
    TryDequeueParseDeviceError = Me.TryParseError(Me.DequeueDeviceError(), _
        a_number, a_message, a_decorator)
End Function

''' <summary>   Queries and concatenates device errors into a delimited replay. </summary>
''' <param name="a_errors">      [String] a delimited string with errors or the no error message
'''                              if none. </param>
''' <param name="a_delimiter">   [Optional, String, VBA.vbCrLf] the delimiter to use when building
'''                              the reply. Note that a comma separates the device error number and error
'''                              description. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if errors available. </value>
Public Function TryDequeueDeviceErrors(ByRef a_errors As String, _
    Optional ByVal a_decorator As String = """", _
    Optional ByVal a_delimiter As String = VBA.vbCrLf) As Boolean

    Dim p_delimiter As String
    Dim p_errorAvailable As Boolean
    On Error Resume Next
    
    Dim p_builder As String
    p_builder = VBA.vbNullString
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_done As Boolean
    p_done = False
    
    Do
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        If Me.TryDequeueParseDeviceError(p_errorNumber, p_errorMessage, a_decorator) Then
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & p_errorMessage
            p_done = "0" = p_errorNumber
        Else
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & p_errorMessage
            p_done = True
            p_errorAvailable = True
        End If
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        If Err.Number <> 0 Then
            p_builder = p_builder & p_delimiter & VBA.CStr(Err.Number) & "," & Err.Description
            p_done = True
            p_errorAvailable = True
        End If
    
    Loop Until p_done
    
    On Error GoTo 0
    
    a_errors = p_builder
    TryDequeueDeviceErrors = p_errorAvailable
    

End Function

''' <summary>   Tries to parse a device error into a number and a message. </summary>
''' <param name="a_value">       [String] the device error. </param>
''' <param name="a_number">      [String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Public Function TryParseError(ByVal a_value As String, _
    ByRef a_number As String, ByRef a_message As String, _
    Optional ByVal a_decorator As String = """") As Boolean
    
    On Error Resume Next
    
    Dim p_success As Boolean
    
    a_number = cc_isr_Core.StringExtensions.Pop(a_value, ",")
    
    If VBA.vbNullString = a_number Then
        a_number = "0"
        a_message = "Parsing '" & a_value & "' failed getting the error number."
        p_success = False
    Else
        a_message = cc_isr_Core.StringExtensions.Pop(a_value, ",")
        If VBA.vbNullString = a_message Then
            a_message = "Parsing '" & a_value & "' failed getting the error message."
            p_success = False
        Else
            a_message = cc_isr_Core.StringExtensions.TrimLeft(a_message, a_decorator)
            a_message = cc_isr_Core.StringExtensions.TrimRight(a_message, a_decorator)
            p_success = True
        End If
        
    End If
    
    ' if an exception, set the error number and message to the error info.
    If Err.Number <> 0 Then
        a_number = VBA.CStr(Err.Number)
        a_message = Err.Description
        p_success = False
    End If
    
    On Error GoTo 0

    TryParseError = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Misc
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Query INPUTS switch (0=rear, 1=front). </summary>
''' <value>   [Boolean] true if the inputs are set to the front panel. </value>
Public Function QueryFrontSwitch() As Boolean
    QueryFrontSwitch = 1 = This.ViSession.QueryLine(":SYST:FRSW?")
End Function

''' <summary>   Return to :SYST:PRES defaults. </summary>
Public Sub Preset()
    This.ViSession.WriteLine (":SYST:PRES")
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
