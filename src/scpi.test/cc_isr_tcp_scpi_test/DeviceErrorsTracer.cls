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
    K2700 As cc_isr_Tcp_Scpi.K2700
End Type

Private This As this_

Implements cc_isr_Test_Fx.IErrTracer

''' <summary>   Initializes this object. </summary>
''' <para name="a_k2700">   [K2700] reference tot he instrument object. </param>
Public Function Initialize(ByVal a_k2700 As cc_isr_Tcp_Scpi.K2700) As DeviceErrorsTracer
    Set This.K2700 = a_k2700
    Set Initialize = Me
End Function

''' <summary>   Asserts and traces leftover errors from the device error queue
''' as well as the user defined errors queue and archive. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Private Function IErrTracer_AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_deviceErrors As String
    Dim p_hasDeviceErrors As Boolean
    If This.K2700 Is Nothing Then
        p_deviceErrors = VBA.vbNullString
    ElseIf This.K2700.Connected Then
        p_hasDeviceErrors = This.K2700.DeviceErrorReader.TryDequeueDeviceErrors(p_deviceErrors)
    End If
    
    ' if has device errors, enqueue the device error(s) to be asserted below.
    If p_hasDeviceErrors Then
        cc_isr_Core_IO.UserDefinedErrors.EnqueueError _
            cc_isr_Core_IO.UserDefinedErrors.Ieee488DeviceError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".IErrTracer_AssertLeftoverErrors", _
            " Device error(s): " & p_deviceErrors & "."
    End If
    
    Dim p_archivedErrorCount As Integer
    p_archivedErrorCount = cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount
    
    Dim p_archivedErrors As String
    Dim p_queuedErrorCount As Integer
    p_queuedErrorCount = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
    If p_archivedErrorCount > 0 Or p_queuedErrorCount > 0 Then
        
        ' if we have leftover errors, move the queued errors to the archive
        Dim p_queuedErrors As String
        p_queuedErrors = cc_isr_Core_IO.UserDefinedErrors.ArchiveQueuedErrors()
        
        ' build the errors to report from the archive.
        p_archivedErrors = cc_isr_Core_IO.UserDefinedErrors.PopArchivedErrors()
        
        Set p_outcome = Assert.Fail( _
            VBA.CStr(p_queuedErrorCount) & " errors were found in the error queue and " & _
            VBA.CStr(p_archivedErrorCount) & " errors were found in the error stack:" & _
            VBA.vbCrLf & p_archivedErrors)
    Else
        
        Set p_outcome = Assert.Pass("There were no leftover errors to assert.")
        
    End If
    
    Set IErrTracer_AssertLeftoverErrors = p_outcome

End Function

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Private Sub IErrTracer_TraceError(ByVal a_message As String)
    UnitTestSheet.LastErrorMessage = a_message
End Sub





