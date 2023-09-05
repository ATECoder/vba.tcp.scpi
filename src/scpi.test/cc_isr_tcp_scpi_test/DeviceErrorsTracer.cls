VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeviceErrorsTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can trace device error messages. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    K2700 As K2700
End Type

Private This As this_

Implements cc_isr_Test_Fx.IErrTracer

''' <summary>   Initializes this object. </summary>
''' <para name="a_k2700">   [K2700] reference tot he instrument object. </param>
Public Function Initialize(ByVal a_k2700 As K2700) As DeviceErrorsTracer
    Set This.K2700 = a_k2700
    Set Initialize = Me
End Function

''' <summary>   Asserts and traces leftover errors from the device error queue. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Private Function IErrTracer_AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_actualErrorMessages As String
    Dim p_hasErrors As Boolean
    p_hasErrors = This.K2700.DeviceErrorReader.TryDequeueDeviceErrors(p_actualErrorMessages)
    Set p_outcome = Assert.IsFalse(p_hasErrors, _
        "Found errors: " & VBA.vbCrLf & p_actualErrorMessages)
    
    Set IErrTracer_AssertLeftoverErrors = p_outcome

End Function

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Private Sub IErrTracer_TraceError(ByVal a_message As String)
    UnitTestSheet.LastErrorMessage = a_message
End Sub





