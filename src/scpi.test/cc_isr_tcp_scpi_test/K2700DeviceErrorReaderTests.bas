Attribute VB_Name = "K2700DeviceErrorReaderTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 Device Error Reader Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_FX.Assert
    BeforeEachAssert As cc_isr_Test_FX.Assert
    Device As cc_isr_Tcp_Scpi.K2700
    Host As String
    Port As Long
    SocketReceiveTimeout As Integer
    ErrTracer As IErrTracer
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
            TestNoErrorShouldParse
        Case 2
            TestUndefinedHeaderErrorShouldParse
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
    For p_testNumber = 1 To 2
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

    This.Name = "K2700DeviceErrorReaderTests"
    
    This.TestNumber = 0
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.SocketReceiveTimeout = 100
    
    Set This.Device = cc_isr_Tcp_Scpi.Factory.NewK2700().Initialize()
    
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

    Set This.Device = Nothing

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

''' <summary>   Unit test. Asserts the device <c>No Error</c> should. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestNoErrorShouldParse() As cc_isr_Test_FX.Assert

    Const p_procedureName As String = "TestNoErrorShouldParse"

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
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNotNothing(This.Device.DeviceErrorReader, _
            "Device Error Reader should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        
        p_success = This.Device.DeviceErrorReader.TryDequeueParseDeviceError(p_errorNumber, p_errorMessage)
        Set p_outcome = Assert.IsTrue(p_success, _
            "Device Error Reader should dequeue and parse the last device error.")

    End If

    Dim p_expectedErrorNumber As String: p_expectedErrorNumber = "0"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorNumber, p_errorNumber, _
            "Device Error Reader should dequeue the 'No Error' error number.")

    End If

    Dim p_expectedErrorMessage As String: p_expectedErrorMessage = "No error"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessage, p_errorMessage, _
            "Device Error Reader should dequeue the 'No Error' error message.")

    End If

    Dim p_actualErrorMessages As String
    Dim p_expectedErrorMessages As String: p_expectedErrorMessages = "0,No error"
    If p_outcome.AssertSuccessful Then
        
        p_success = Not This.Device.DeviceErrorReader.TryDequeueDeviceErrors(p_actualErrorMessages)
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessages, p_actualErrorMessages, _
            "Device Error Reader should dequeue the '0,No Error' error messages.")

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestNoErrorShouldParse")
    
    Set TestNoErrorShouldParse = p_outcome
    
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


''' <summary>   Unit test. Asserts parsing device <c>UndefinedHeader</c> Error. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestUndefinedHeaderErrorShouldParse() As cc_isr_Test_FX.Assert
    
    Const p_procedureName As String = "TestUndefinedHeaderErrorShouldParse"

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
    
    ' validate the existence of no errors.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = TestNoErrorShouldParse()
    
    If p_outcome.AssertSuccessful Then
    
        ' create and fetch the error.
        This.Device.Device.ViSession.WriteLine "**CLS", False
        cc_isr_Core_IO.Factory.NewStopwatch.Wait 50
        p_success = This.Device.DeviceErrorReader.TryDequeueParseDeviceError(p_errorNumber, p_errorMessage)
        Set p_outcome = Assert.IsTrue(p_success, _
            "Device Error Reader should dequeue and parse the last device error.")

    End If

    Dim p_expectedErrorMessage As String: p_expectedErrorMessage = "Undefined header"
    Dim p_expectedErrorNumber As String: p_expectedErrorNumber = "-113"
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorNumber, p_errorNumber, _
            "Device Error Reader should dequeue the '" & _
            p_expectedErrorMessage & "' error number.")

    End If

    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessage, p_errorMessage, _
            "Device Error Reader should dequeue the expected error message.")

    End If

    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedReply As String: p_expectedReply = "1"
        Dim p_actualReply As String: p_actualReply = This.Device.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "The Device shoudl query operation completion.")

    End If

    Dim p_actualErrorMessages As String
    Dim p_expectedErrorMessages As String: p_expectedErrorMessages = "0,No error"
    If p_outcome.AssertSuccessful Then
        
        p_success = Not This.Device.DeviceErrorReader.TryDequeueDeviceErrors(p_actualErrorMessages)
        Set p_outcome = Assert.AreEqual(p_expectedErrorMessages, p_actualErrorMessages, _
            "Device Error Reader should dequeue the '0,No Error' error messages.")

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestUndefinedHeaderErrorShouldParse")
    
    Set TestUndefinedHeaderErrorShouldParse = p_outcome
   
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

