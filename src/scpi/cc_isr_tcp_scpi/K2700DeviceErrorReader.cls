VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700DeviceErrorReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 Implementation of the Device Error Reader interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements IDeviceErrorReader

Private Type this_
    ViSession As cc_isr_Ieee488.ViSession
    LastDeviceError As String
    LastDeviceErrors As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="IDeviceErrorReader"/>]. </returns>
Private Function IDeviceErrorReader_Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As IDeviceErrorReader
    Set This.ViSession = a_session
    Set IDeviceErrorReader_Initialize = Me
End Function

''' <summary>   The last error that was dequeued from the device. </summary>
''' <value>   [String] </value>
Private Property Get IDeviceErrorReader_LastDeviceError() As String
    IDeviceErrorReader_LastDeviceError = This.LastDeviceError
End Property

''' <summary>   The last errors that were dequeued from the device. </summary>
''' <value>   [String] </value>
Private Property Get IDeviceErrorReader_LastDeviceErrors() As String
    IDeviceErrorReader_LastDeviceErrors = This.LastDeviceErrors
End Property

''' <summary>   Dequeues the top device error from the device error queue. </summary>
''' <value>   [String] the top error on the error queue. </value>
Private Function IDeviceErrorReader_DequeueDeviceError() As String
    This.LastDeviceError = This.ViSession.QueryLine(":SYST:ERR?")
    IDeviceErrorReader_DequeueDeviceError = This.LastDeviceError
End Function

''' <summary>   Clears the error queue. </summary>
Private Sub IDeviceErrorReader_ClearDeviceErrors()
    This.ViSession.WriteLine (":SYST:CLE")
End Sub

''' <summary>   Queries and tries to parse a device error into a number and a message. </summary>
''' <param name="a_number">      [String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Private Function IDeviceErrorReader_TryDequeueParseDeviceError(ByRef a_number As String, _
    ByRef a_message As String, Optional ByVal a_decorator As String = """") As Boolean
    IDeviceErrorReader_TryDequeueParseDeviceError = IDeviceErrorReader_TryParseError(IDeviceErrorReader_DequeueDeviceError(), _
        a_number, a_message, a_decorator)
End Function

''' <summary>   Queries and concatenates device errors into a delimited replay. </summary>
''' <param name="a_errors">      [String] a delimited string with errors or the no error message
'''                              if none. </param>
''' <param name="a_delimiter">   [Optional, String, VBA.vbCrLf] the delimiter to use when building
'''                              the reply. Note that a comma separates the device error number and error
'''                              description. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Integer] the number of dequeued errors. </value>
Private Function IDeviceErrorReader_TryDequeueDeviceErrors(ByRef a_errors As String, _
    Optional ByVal a_decorator As String = """", _
    Optional ByVal a_delimiter As String = VBA.vbCrLf) As Integer

    Dim p_delimiter As String
    Dim p_errorsCount As Integer
    On Error Resume Next
    
    Dim p_builder As String
    p_builder = VBA.vbNullString
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_done As Boolean
    p_done = False
    
    Do
    
        DoEvents
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        If IDeviceErrorReader_TryDequeueParseDeviceError(p_errorNumber, p_errorMessage, a_decorator) Then
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & p_errorMessage
            p_done = "0" = p_errorNumber
        Else
            p_builder = p_builder & p_delimiter & p_errorNumber & "," & p_errorMessage
            p_done = True
            p_errorsCount = p_errorsCount + 1
        End If
        
        p_delimiter = IIf(VBA.vbNullString = p_builder, VBA.vbNullString, a_delimiter)
        
        ' enqueue a VBA Error as a device error.
        If Err.Number <> 0 Then
            p_builder = p_builder & p_delimiter & VBA.CStr(Err.Number) & ",Application error: " & Err.Description
            p_done = True
            p_errorsCount = p_errorsCount + 1
        End If
    
    Loop Until p_done
    
    On Error GoTo 0
    
    a_errors = p_builder
    IDeviceErrorReader_TryDequeueDeviceErrors = p_errorsCount
    

End Function

''' <summary>   Tries to parse a device error into a number and a message. </summary>
''' <param name="a_value">       [String] the device error. </param>
''' <param name="a_number">      [String] the error number; set to the <see cref="Err.Number"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_message">     [String] the error message; set to the <see cref="Err.Description"/>
'''                              if exception occurred parsing. </param>
''' <param name="a_decorator">   [String] the decorator around the message part of the device error. </param>
''' <value>   [Boolean] true if parsed. </value>
Private Function IDeviceErrorReader_TryParseError(ByVal a_value As String, _
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

    IDeviceErrorReader_TryParseError = p_success

End Function

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
