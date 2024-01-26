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
    Device As cc_isr_Ieee488.Device
End Type

Private This As this_

Implements cc_isr_Test_Fx.IErrTracer

''' <summary>   Initializes this object. </summary>
''' <para name="a_device">   [Device] reference tot he instrument object. </param>
Public Function Initialize(ByVal a_device As cc_isr_Ieee488.Device) As DeviceErrorsTracer
    Set This.Device = a_device
    Set Initialize = Me
End Function

''' <summary>   Asserts and traces leftover errors from the error queue and error stack. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
''' <returns>   [String]. </returns>
Friend Function EnumerateErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    Dim p_archivedErrorCount As Integer
    p_archivedErrorCount = cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount
    Dim p_queuedErrorCount As Integer
    p_queuedErrorCount = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
    
    Dim p_outcome As String
    p_outcome = VBA.vbNullString
    If p_archivedErrorCount + p_queuedErrorCount > 0 Then
        
        ' if we have leftover errors, move the queued errors to the archive
        Dim p_queuedErrors As String
        p_queuedErrors = cc_isr_Core_IO.UserDefinedErrors.ArchiveQueuedErrors()
        
        ' build the errors to report from the archive.
        Dim p_archivedErrors As String
        p_archivedErrors = cc_isr_Core_IO.UserDefinedErrors.PopArchivedErrors()
        
        If p_queuedErrorCount > 0 Then _
            p_outcome = "Found Enqueued error(s) (" & VBA.CStr(p_queuedErrorCount) & ")"

        If p_archivedErrorCount > 0 Then _
            p_outcome = p_outcome & IIf(p_queuedErrorCount > 0, " and ", "Found ") & _
                "Archived error(s) (" & VBA.CStr(p_queuedErrorCount) & ")"
        
        p_outcome = p_outcome & ":" & VBA.vbCrLf & p_archivedErrors
        
    End If
    
    EnumerateErrors = p_outcome
    
End Function

''' <summary>   Asserts and traces leftover errors from the device error queue
''' as well as the user defined errors queue and archive. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Private Function IErrTracer_AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_errors As String
    If This.Device Is Nothing Then
        p_errors = Me.EnumerateErrors(a_delimiter)
    ElseIf This.Device.Connected Then
        p_errors = This.Device.EnumerateErrors(a_delimiter)
    Else
        p_errors = Me.EnumerateErrors(a_delimiter)
    End If
    If VBA.vbNullString = p_errors Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("There are no leftover errors to assert.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_errors)
    End If
    
    Set IErrTracer_AssertLeftoverErrors = p_outcome

End Function

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Private Sub IErrTracer_TraceError(ByVal a_message As String)
    UnitTestSheet.LastErrorMessage = a_message
End Sub





