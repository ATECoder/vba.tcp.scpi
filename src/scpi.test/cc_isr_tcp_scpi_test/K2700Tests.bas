Attribute VB_Name = "K2700Tests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Device Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    Device As cc_isr_Tcp_Scpi.K2700
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
            TestQueryOperationCompletion
        Case 2
            TestRecoveryFromSyntaxFromError
        Case 3
            TestRecoveryFromReadAfterWriteTrue
        Case Else
    End Select
    AfterEach
End Sub

Public Sub RunOneTest()
    BeforeAll
    RunTest 3
    AfterAll
End Sub

Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 3
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
    
    Set This.Device = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize()
    
    ' trap errors in case connection fails rendering all tests inconclusive.
    
    On Error Resume Next
    
    This.Device.OpenConnection This.Host, This.Port, This.SocketReceiveTimeout
    
    Dim p_leftoverErrorMessage As String
    p_leftoverErrorMessage = VBA.vbNullString
    
    If Err.Number <> 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 Device failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 Device failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf This.Device.Connected Then
        Set This.BeforeAllAssert = Assert.IsTrue(True, "Connected")
    Else
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 Device should be connected")
    End If
    
    This.ErrTracer.TraceError p_leftoverErrorMessage
    
    ' clear the error object.
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = IIf(This.Device.Connected, _
            Assert.IsTrue(True, "Connected"), _
            Assert.Inconclusive("K2700 Device should be connected"))
    
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
        This.Device.Device.ClearExecutionState

                        
End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    
    ' disconnect if connected
    If Not This.Device Is Nothing Then _
        This.Device.CloseConnection

    If Not This.Device Is Nothing Then This.Device.Dispose
    Set This.Device = Nothing

    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts querying operation completion. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestQueryOperationCompletion() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
        Dim p_actualReply As String
        Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "K2700 Device should query operation completion.")

    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestQueryOperationCompletion")
    
    Set TestQueryOperationCompletion = p_outcome
    
End Function

''' <summary>   Unit test. Asserts recovery from Syntax error. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRecoveryFromSyntaxFromError() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
        Dim p_actualReply As String
        Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "K2700 Device should query operation completion.")
    End If

    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.Device.ViSession.WriteLine ("**OPC")
        On Error GoTo 0
        
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
        p_expectedReply = "1"
        p_actualReply = This.Device.Device.ClearExecutionState()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "K2700 Device should query operation completion.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors

    Debug.Print p_outcome.BuildReport("TestQueryOperationCompletion")
    
    Set TestRecoveryFromSyntaxFromError = p_outcome
    
End Function

''' <summary>   Unit test. Asserts recovery from read after write true condition. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRecoveryFromReadAfterWriteTrue() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
        Dim p_actualReply As String
        Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful And This.Device.ViSession.UsingGpibLan Then
        ' turn on read after write condition.
        This.Device.ViSession.GpibLan.ReadAfterWriteEnabledSetter True
        Set p_outcome = Assert.IsTrue(This.Device.ViSession.GpibLan.ReadAfterWriteEnabledGetter, _
            "Read after write should be true.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Device.CloseConnection
        Set p_outcome = Assert.IsFalse(This.Device.Connected, _
            "K2700 Device should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.Device.OpenConnection This.Host, This.Port
        Set p_outcome = Assert.IsTrue(This.Device.Connected, _
            "K2700 Device should be connected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.Device.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "K2700 Device should query operation completion.")
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors

    Debug.Print p_outcome.BuildReport("TestRecoveryFromReadAfterWriteTrue")
    
    Set TestRecoveryFromReadAfterWriteTrue = p_outcome
    
End Function


