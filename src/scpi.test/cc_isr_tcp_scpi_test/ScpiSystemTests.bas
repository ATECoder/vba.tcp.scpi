Attribute VB_Name = "ScpiSystemTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Scpi System Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    BeforeEachAssert As Assert
    K2700 As cc_isr_Tcp_Scpi.K2700
    Host As String
    Port As Long
    SocketReceiveTimeout As Integer
    ErrTracer As ErrTracer
End Type

Private This As this_

Public Sub BeforeAll()

    This.Host = "192.168.0.252"
    This.Port = 1234
    This.SocketReceiveTimeout = 100
    
    Set This.ErrTracer = New ErrTracer
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700().Initialize(This.ErrTracer)
    
    ' trap errors in case connection fails rendering all tests inconclusive.
    On Error Resume Next
    This.K2700.OpenConnection This.Host, This.Port, This.SocketReceiveTimeout
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    Set This.BeforeEachAssert = IIf(This.K2700.Connected, _
        Assert.IsTrue(True, "Connected"), _
        Assert.Inconclusive("View Model should be connected"))
                        
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

    Debug.Print p_outcome.BuildReport("TestParsingDeviceError")
    
    Set TestParsingDeviceError = p_outcome
    
End Function

Public Sub RunTests()
    BeforeAll
    BeforeEach
    TestParsingDeviceError
    AfterEach
    AfterAll
End Sub

