Attribute VB_Name = "K2700DeviceErrorTracerTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 Device Error Tracer Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    K2700 As cc_isr_Tcp_Scpi.K2700
    Host As String
    Port As Long
    SocketReceiveTimeout As Integer
    ErrTracer As IErrTracer
End Type

Private This As this_

Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestParsingDeviceError
        Case Else
    End Select
    AfterEach
End Sub

Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 4
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

Public Sub BeforeAll()

    This.TestNumber = 0
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.SocketReceiveTimeout = 100
    
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.ErrTracer = New ErrTracer
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700().Initialize(This.ErrTracer)
    
    ' trap errors in case connection fails rendering all tests inconclusive.
    
    On Error Resume Next
    
    This.K2700.OpenConnection This.Host, This.Port, This.SocketReceiveTimeout
    
    Dim p_leftoverErrorMessage As String
    p_leftoverErrorMessage = VBA.vbNullString
    
    If Err.Number <> 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf This.K2700.Connected Then
        Set This.BeforeAllAssert = Assert.IsTrue(True, "Connected")
    Else
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 should be connected")
    End If
    
    This.ErrTracer.TraceError p_leftoverErrorMessage
    
    ' clear the error object.
    
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    Set This.BeforeEachAssert = Assert.IsTrue(True, "initialize the pre-test assert.")

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = IIf(This.K2700.Connected, _
            Assert.IsTrue(True, "Connected"), _
            Assert.Inconclusive("K2700 should be connected"))
    
    Else
    
        Set This.BeforeEachAssert = Assert.Inconclusive(This.BeforeAllAssert.AssertMessage)
    
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.AreEqual(0, Err.Number, _
            "Error Number should be 0.")
            
    End If
    
    This.TestNumber = This.TestNumber + 1
    
    ' clear execution state before each test.
    
    If This.BeforeEachAssert.AssertSuccessful Then _
        This.K2700.Device.ClearExecutionState
    
End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    
    ' disconnect if connected
    If Not This.K2700 Is Nothing Then _
        This.K2700.CloseConnection

    If Not This.K2700 Is Nothing Then This.K2700.Dispose
    Set This.K2700 = Nothing

    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts parsing device error. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestParsingDeviceError() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNotNothing(This.K2700.ScpiSystem, _
            "Scpi System should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        
        p_success = This.K2700.ScpiSystem.TryDequeueParseDeviceError(p_errorNumber, p_errorMessage)
        Set p_outcome = Assert.IsTrue(p_success, _
            "Scpi System should dequeue and parse the last device error.")

    End If

    Dim p_expectedErrorNumber As String: p_expectedErrorNumber = "0"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorNumber, p_errorNumber, _
            "Scpi System should dequeue the 'No Error' error number.")

    End If

    Dim p_expectedErrorMessage As String: p_expectedErrorMessage = "No error"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessage, p_errorMessage, _
            "Scpi System should dequeue the 'No Error' error message.")

    End If

    Dim p_actualErrorMessages As String
    Dim p_expectedErrorMessages As String: p_expectedErrorMessages = "0,No error"
    If p_outcome.AssertSuccessful Then
        
        p_success = Not This.K2700.ScpiSystem.TryDequeueDeviceErrors(p_actualErrorMessages)
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessages, p_actualErrorMessages, _
            "Scpi System should dequeue the '0,No Error' error messages.")

    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors

    Debug.Print p_outcome.BuildReport("TestParsingDeviceError")
    
    Set TestParsingDeviceError = p_outcome
    
End Function



