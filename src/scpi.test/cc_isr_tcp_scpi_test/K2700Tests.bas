Attribute VB_Name = "K2700Tests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Device Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    Device As cc_isr_Tcp_Scpi.K2700
    Host As String
    Port As Long
    SocketReceiveTimeout As Integer
    ErrTracer As cc_isr_Test_FX.IErrTracer
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
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

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 3
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests initialize and cleanup.
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Prepares all tests. </summary>
''' <remarks>   This method sets up the 'Before All' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to set the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_FX.Assert: Set p_outcome = Assert.Pass("Primed to run all tests.")

    This.Name = "K2700Tests"
    
    This.TestNumber = 0
    
    ' set a temporate error tracer.
    Set This.ErrTracer = New DeviceErrorsTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.SocketReceiveTimeout = 100
    
    Set This.Device = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize()
    
    Dim p_errTracer As New DeviceErrorsTracer
    Set This.ErrTracer = p_errTracer.Initialize(This.Device)
    
    This.Device.OpenConnection This.Host, This.Port, This.SocketReceiveTimeout
    
    If This.Device.Connected Then
        Set p_outcome = Assert.Pass("Primed to run all tests; K2700 is connected.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; K2700 should be connected.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming all tests;" & _
                VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeAllAssert = p_outcome
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Prepares each test before it is run. </summary>
''' <remarks>   This method sets up the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to initialize the <see cref="cc_isr_Test_Fx.Assert"/> of each test.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    This.TestNumber = This.TestNumber + 1

    Dim p_outcome As cc_isr_Test_FX.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.Device.Connected, _
            Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & "; K2700 is Connected."), _
            Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                "; K2700 should be connected."))
    Else
        Set p_outcome = Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
   
    ' clear execution state before each test.
    
    If p_outcome.AssertSuccessful Then _
        This.Device.Device.ClearExecutionState

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
             Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_FX.Assert
    Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    ' cleanup after each test.
    If This.BeforeEachAssert.AssertSuccessful Then
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' assert.
    Set This.BeforeEachAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_FX.Assert: Set p_outcome = Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    If This.BeforeAllAssert.AssertSuccessful Then
    End If
    
    ' disconnect if connected
    If Not This.Device Is Nothing Then _
        This.Device.Dispose

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up all tests;" & _
            VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    Set This.ErrTracer = Nothing
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts querying operation completion. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestQueryOperationCompletion() As Assert

    Const p_procedureName As String = "TestQueryOperationCompletion"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
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

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestQueryOperationCompletion")
    
    Set TestQueryOperationCompletion = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Unit test. Asserts recovery from Syntax error. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestRecoveryFromSyntaxFromError() As Assert

    Const p_procedureName As String = "TestRecoveryFromSyntaxFromError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
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
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestQueryOperationCompletion")
    
    Set TestRecoveryFromSyntaxFromError = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Unit test. Asserts recovery from read after write true condition. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestRecoveryFromReadAfterWriteTrue() As Assert

    Const p_procedureName As String = "TestRecoveryFromReadAfterWriteTrue"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
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

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestRecoveryFromReadAfterWriteTrue")
    
    Set TestRecoveryFromReadAfterWriteTrue = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function


