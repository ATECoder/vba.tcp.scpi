VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDeviceErrorReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can read device errors. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Returns a reference to the <see cref="IDeviceErrorReader"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="IDeviceErrorReader"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As IDeviceErrorReader
End Function

''' <summary>   The last error that was dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastDeviceError() As String
End Property

''' <summary>   The last errors that were dequeued from the device. </summary>
''' <value>   [String] </value>
Public Property Get LastDeviceErrors() As String
End Property

''' <summary>   Dequeues the top device error from the device error queue. </summary>
''' <value>   [String] the top error on the error queue. </value>
Public Function DequeueDeviceError() As String
End Function

''' <summary>   Clears the error queue. </summary>
Public Sub ClearDeviceErrors()
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
End Function

