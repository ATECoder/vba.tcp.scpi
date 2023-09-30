Attribute VB_Name = "K2700ViewModelTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 View Model Tests. </summary>
''' <remarks>   Dependencies: cc_isr_Core_Tcp_Scpi.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    
    ' unit test settings
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As cc_isr_Test_Fx.IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
    
    ' initial observer settings
    Observer As K2700Observer
    Address As String
    ReadingOffset As Double
    GpibLanControllerPort As Integer
    
    ' initial view model settings
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    SessionTimeout As Integer
    
    ContinuousSenseFunctionName As String
    ImmediateSenseFunctionName As String
    ExternalSenseFunctionName As String
    PrimaryGpibAddress As Integer
    
    ' known information
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestShouldInitialize
        Case 2
            Set p_outcome = TestShouldBeConnected
        Case 3
            Set p_outcome = TestShouldReadCards
        Case 4
            Set p_outcome = TestShouldRestoreInitialState
        Case 5
            Set p_outcome = TestShouldRecoverFromSyntaxError
        Case 6
            Set p_outcome = TestShouldRestoreFromClosedConnection
        Case 7
            Set p_outcome = TestShouldConfigureImmediateMode
        Case 8
            Set p_outcome = TestShouldConfigureExternalMode
        Case 9
            Set p_outcome = TestShouldMonitorTriggering
        Case Else
    End Select
    AfterEach
    Set RunTest = p_outcome
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 9
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' Test 01 TestShouldInitialize passed. Elapsed time: 16.4 ms.
''' Test 02 TestShouldBeConnected passed. Elapsed time: 31.4 ms.
'''     Serial Poll is 16 in 11.8ms.
''' Test 03 TestShouldReadCards passed. Elapsed time: 14.9 ms.
''' Test 04 TestShouldRestoreInitialState passed. Elapsed time: 12253.3 ms.
''' Test 05 TestShouldRecoverFromSyntaxError passed. Elapsed time: 160.4 ms.
'''    Serial Poll is 4 in 3.9 ms.
''' Test 06 TestShouldRestoreFromClosedConnection passed. Elapsed time: 5745.9 ms.
''' Test 07 TestShouldConfigureImmediateMode passed. Elapsed time: 5629.6 ms.
''' Test 08 TestShouldConfigureExternalMode passed. Elapsed time: 4306.3 ms.
'''
''' Test 01 TestShouldInitialize passed. Elapsed time: 11.6 ms.
''' Test 02 TestShouldBeConnected passed. Elapsed time: 35.5 ms.
'''     Serial Poll is 80 in 12.5 ms.
''' Test 03 TestShouldReadCards passed. Elapsed time: 14.5 ms.
''' Test 04 TestShouldRestoreInitialState passed. Elapsed time: 12232.9 ms.
''' Test 05 TestShouldRecoverFromSyntaxError passed. Elapsed time: 157.1 ms.
'''     Serial Poll is 68 in 4.3 ms.
''' Test 06 TestShouldRestoreFromClosedConnection passed. Elapsed time: 5740.3 ms.
''' Ran 6 out of 6 tests.
''' Passed: 6; Failed: 0; Inconclusive: 0.
''' </code>
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 6
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
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

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")

    This.Name = "K2700ViewModelTests"
    
    ' initialize test settings
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    
    ' initialize observer settings
    This.Address = "192.168.0.252:1234"
    This.ReadingOffset = 0.01
    This.GpibLanControllerPort = 1234
    
    ' initialize view model settings
    This.SessionTimeout = 3000
    
    ' initialize known data.
    This.TopCard = "7700"
    This.BottomCard = VBA.vbNullString
    This.ContinuousSenseFunctionName = "FRES"
    This.ImmediateSenseFunctionName = "RES"
    This.ExternalSenseFunctionName = "FRES"
    This.PrimaryGpibAddress = 16
    
    ' set a temporary error tracer
    Dim p_errTrace As New DeviceErrorsTracer
    Set This.ErrTracer = p_errTrace
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
    ' card scan list uses immediate mode sense function
    This.TopCardFunctionScanList = ":FUNC 'RES',(@101,120)"
    This.BottomCardFunctionScanList = VBA.vbNullString
    
    Set This.ViewModel = cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
    
    Set This.ErrTracer = p_errTrace.Initialize(This.ViewModel.Device)
    
    ' set observer user interface settings
    K2700Observer.SocketAddress = This.Address
    K2700Observer.ReadingOffset = This.ReadingOffset
    K2700Observer.GpibLanControllerPort = This.GpibLanControllerPort
    
    ' initialize the observer before initializing the view mode
    ' but after the observer setting are set. The observer initial
    ' settings are then applied to the view model.
    Set This.Observer = K2700Observer.Initialize(This.ViewModel)
    
    ' set view model initial settings
    This.ViewModel.SessionTimeout = This.SessionTimeout
    
    ' issue the open connection command. This initializes the view model.
    This.ViewModel.OpenConnectionCommand
    
    If This.ViewModel.Connected Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests; K2700 View Model is connected.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive( _
            "Failed priming all tests; K2700 View Model should be connected.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming all tests;" & _
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

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.ViewModel.Connected, _
            Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & "; K2700 View Model is Connected."), _
            Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                "; K2700 View Model should be connected."))
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    Dim p_details As String: p_details = VBA.vbNullString
   
    If p_outcome.AssertSuccessful Then
        
        ' clear execution state before each test.
        ' clear errors if any so as to leave the instrument without errors.
        ' here we add *OPC? to prevent the query unterminated error.
        
        Dim p_command As String
        p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.ViewModel.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
    
    Dim p_reply As String
    If p_outcome.AssertSuccessful Then
        If 0 > This.ViewModel.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            "; Operation completion query should return the correct reply.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryClearExecutionState(p_details), _
            p_details)
    End If
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
             Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    
    This.TestStopper.Restart
    
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

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    ' check if we can proceed with cleanup.
    
    If Not This.BeforeEachAssert.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to cleanup test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeEachAssert.AssertMessage)

    ' cleanup after each test.
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_command As String
        Dim p_reply As String
        Dim p_details As String: p_details = VBA.vbNullString
    
        ' clear errors if any so as to leave the instrument without errors.
        p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.ViewModel.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
        If 0 > This.ViewModel.Session.TryRead(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' assert.
    Set This.BeforeEachAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
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
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    If This.BeforeAllAssert.AssertSuccessful Then
        This.ViewModel.ResetKnownStateCommand
    End If
    
    ' disconnect if connected
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.Dispose

    Set This.ViewModel = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up all tests;" & _
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
' Asserts
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' summary>   Asserts that the status byte bits value are correct. </summary>
''' <param name="a_bitsStatus"/>   [Integer] The expected status of the specified status bits. </param>
''' <param name="a_statusBits"/>   [Integer] The expected status bits. </param>
''' <param name="a_statusByte"/>   [Out, Integer] The status byte. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Private Function AssertSerialPollShouldValidate(ByVal a_bitsStatus As Integer, _
    ByVal a_statusBits As Integer, ByRef a_statusByte As Integer) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_details As String
    Dim p_polled As Boolean
    Dim p_elapsed As Double
    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    p_stopper.Restart
    p_polled = This.ViewModel.Session.AwaitStatusBits(a_bitsStatus, a_statusBits, 3000, a_statusByte, p_details)
    p_elapsed = p_stopper.ElapsedMilliseconds
    If a_statusByte < 0 Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    ElseIf p_polled Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Serial Poll is " & VBA.CStr(a_statusByte) & _
            " in " & Format(p_elapsed, "0.0") & " ms.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail("    Status byte '" & _
            VBA.CStr(a_statusByte) & "' bits '" & VBA.CStr(a_statusBits) & _
            "' not matching the expected bits '" & VBA.CStr(a_bitsStatus) & "' value.")
    End If
    
    Set AssertSerialPollShouldValidate = p_outcome

End Function

''' summary>   Asserts that external trigger mode should be configured. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertExternalTriggerModeShouldStart(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' configure external mode with front switch.
        This.ViewModel.FrontInputsRequired = True
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ConfigureExternalTriggerReadingCommand(p_details), _
            p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External trigger reading mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.Observer.CurrentMeasurementMode, _
            "Observer measuerment mode should equal expected value for external trigger reading mode.")
    End If
    
    Set AssertExternalTriggerModeShouldStart = p_outcome

End Function

''' summary>   Asserts that external trigger mode should be validated. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertExternalTriggerModeShouldValidate(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with test validations.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External trigger reading mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.Observer.CurrentMeasurementMode, _
            "Observer measuerment mode should equal expected value for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = This.ExternalSenseFunctionName
        Dim p_actualSenseFunctionName As String:
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqualString(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            VBA.VbCompareMethod.vbTextCompare, _
            "External mode sense function name should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentChannelNumber, This.Observer.CurrentChannelNumber, _
            "Observer current channel number should equal the view model channel number.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSourceOption.External, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "External trigger source should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.K2700.SenseSystem.SenseSystem.AutoRangeEnabledGetter(), _
            "Auto range should be enabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1#, _
            This.ViewModel.K2700.SenseSystem.SenseSystem.PowerLineCyclesGetter(), _
            "The integration rate in power line cycles should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop requested should be on in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be disabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off in external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off in external trigger reading mode.")
    End If
    
    Set AssertExternalTriggerModeShouldValidate = p_outcome

End Function

''' summary>   Asserts that immediate trigger mode should be configured. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertImmediateModeShouldStart(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' configure immediate mode with front switch.
        This.ViewModel.FrontInputsRequired = True
        
        ' returns true of if success. Otherwise, the error should be in the
        ' last error, if the inputs are invalid or the last error message
        ' if the configuration failed.
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ConfigureImmediateTriggerReadingsCommand(p_details), _
            p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.Observer.CurrentMeasurementMode, _
            "Observer measuerment mode should equal expected value for immediate trigger reading mode.")
    End If
    
    Set AssertImmediateModeShouldStart = p_outcome

End Function

''' summary>   Asserts that immediate trigger mode should be be validated. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertImmediateModeShouldValidate(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with validations
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.Observer.CurrentMeasurementMode, _
            "Observer measuerment mode should equal expected value for immediate trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentChannelNumber, This.Observer.CurrentChannelNumber, _
            "Observer current channel number should equal the view model channel number.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSourceOption.Immediate, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "Immediate trigger source should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop requested should be on in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.MeasureExecutable, _
            "Measure command should be enabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.MeasureExecutable, _
            "Observer Measure button should be ensabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be disabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled in immediate mode.")
    End If
    
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off in immediate mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' turn off auto increment so that we can take a single reading.
        Dim p_autoIncrementChannelNumber As Boolean: p_autoIncrementChannelNumber = False
        This.ViewModel.AutoIncrementChannelNoEnabled = p_autoIncrementChannelNumber
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_autoIncrementChannelNumber, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    End If
    
    Set AssertImmediateModeShouldValidate = p_outcome

End Function

''' summary>   Asserts that immediate measurement should take a reading. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMeasureImmediatelyShouldReadValue(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' make sure we are in immediate trigger mode.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' take a reading
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.MeasureImmediatelyCommand(p_details), _
            p_details)
        
    End If
    
    Dim p_reading As String
    Dim p_channelNumber As Integer
    Dim p_readingValue As Double
    
    If p_outcome.AssertSuccessful Then
        
        ' wait for the reading event to take shape.
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
        ' get the reading from the observer.
        p_reading = This.Observer.MeasuredReading
        
        p_channelNumber = This.Observer.MeasuredChannelNumber

        p_readingValue = This.Observer.MeasuredValue
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, p_channelNumber, _
            "Reading should come from the expected channel number.")
            
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_channelNumber, This.Observer.SelectedChannelNumber, _
            "The observer selected channel number should equals the view model channel number.")
            
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(VBA.vbNullString = p_reading, _
            "Reading should not be empty.")
            
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_readingValue > 0, _
            "Reading value should be positive.")
            
    End If
    
    Dim p_epsilon As Double: p_epsilon = 0.0000000001
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreCloseDouble(p_readingValue, VBA.CDbl(p_reading), p_epsilon, _
            "Reading should equal the parsed value.")
            
    End If
    
    Set AssertMeasureImmediatelyShouldReadValue = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode should be configured. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldStart(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' check that the external trigger mode was set.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External Trigger mode should be as expected.")
    End If
    
    ' start monitoring here.
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.StartMonitoringExternalTriggersCommand
        
        ' allow the monitoring to commenct.
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External trigger monitoring mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentMeasurementMode, _
            This.Observer.CurrentMeasurementMode, "Observer external trigger monitoring mode should be as expected.")
    End If
    
    Set AssertMonitoringModeShouldStart = p_outcome

End Function


''' summary>   Asserts that trigger monitoring mode should be validated. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldValidate(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' start validating.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External trigger monitoring mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentMeasurementMode, _
            This.Observer.CurrentMeasurementMode, "Observer external trigger monitoring mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.PauseRequested, _
            "Pause Requested should be off after starting the monitoring timer.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopRequested, _
            "Stop requested should be off in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger should get initiated after starting the monitoring timer.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be enabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be enabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be disabled in trigger monitoring mode.")
    End If
    
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be disabled in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be on in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be on in trigger monitoring mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' turn on auto increment so that we can monitor readings.
        Dim p_autoIncrementChannelNumber As Boolean: p_autoIncrementChannelNumber = True
        This.ViewModel.AutoIncrementChannelNoEnabled = p_autoIncrementChannelNumber
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_autoIncrementChannelNumber, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_selectedChannelNumber As Integer: p_selectedChannelNumber = 1
        This.ViewModel.SelectedChannelNumber = p_selectedChannelNumber
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_selectedChannelNumber, This.ViewModel.CurrentChannelNumber, _
            "Current channel number should be set to the selected channel number.")
    
    End If
   
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentChannelNumber, This.Observer.CurrentChannelNumber, _
            "Observer current channel number should be set to the selected channel number.")
    
    End If
   
    Set AssertMonitoringModeShouldValidate = p_outcome

End Function

''' summary>   Asserts that measurements should get triggered. </summary>
''' <param name="a_outcome">          [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <param name="a_timeoutSeconds">   [Optional, Double, 30] The time to wait for some triggered values. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMeasurementsShouldGetTriggered(ByVal a_outcome As cc_isr_Test_Fx.Assert, _
    Optional ByVal a_timeoutSeconds As Double = 30) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    
    ' get the first channel number
    Dim p_channel As Integer
    p_channel = This.Observer.CurrentChannelNumber
    
    Dim p_reading As String
    p_reading = This.Observer.MeasuredReading
    
    ' loop for some time waiting for triggered measurements.
    
    Dim p_endTime As Double
    p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow() + _
        (a_timeoutSeconds / cc_isr_Core_IO.CoreExtensions.SecondsPerDay)
    While p_endTime > cc_isr_Core_IO.CoreExtensions.DaysNow()
        
        DoEvents
    
        If p_channel <> This.Observer.CurrentChannelNumber Then
        
            DoEvents
            p_channel = This.Observer.CurrentChannelNumber
            DoEvents
            p_reading = This.Observer.MeasuredReading
            DoEvents
            Debug.Print p_channel; ": "; p_reading

        End If
    
    Wend
    
    Set AssertMeasurementsShouldGetTriggered = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode should be stopped. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldStop(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' verify that we stop from an active monitoring mode.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External trigger monitoring mode should be as expected.")
    End If
    
    ' stop monitoring here
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.StopMonitoringExternalTriggersCommand
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop Requested should be on off after stopping monitoring.")
        
    End If
    
    ' allow time for monitoring to stop
    
    If p_outcome.AssertSuccessful Then
    
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.PauseRequested, _
            "Pause should be requested after monitoring stopped.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.None
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Measurement mode should be as expected after monitoring stopped.")
    End If
    
    Set AssertMonitoringModeShouldStop = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode stop should be validated. </summary>
''' <param name="a_outcome">   [<see cref="cc_isr_Test_Fx.Assert"/>] The current outcome. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeStopShouldValidate(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_outcome
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' validate monitoring stop state
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.None
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Measurement mode should be as expected after monitoring stopped.")
    End If
    
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "The stop monitoring executable to should disabled after monitoring stopped.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External monitoring enabled should be off after monitoring stopped.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger should not get initiated after monitoring stopped.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be diabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be enabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled after monitoring stopped.")
    End If
    
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off after monitoring stopped.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CurrentChannelNumber, This.Observer.CurrentChannelNumber, _
            "Observer current channel number should be set to the selected channel number.")
    
    End If
   
    Set AssertMonitoringModeStopShouldValidate = p_outcome

End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts that view model should initialize. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldInitialize passed. in 11.3 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldInitialize() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldInitialize"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ToggleConnectionExecutable, _
            "Toggle connection should be executable after initializing the View Model.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Host, This.Observer.Host, _
            "Observer and view model 'Host' setting should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Port, This.Observer.Port, _
            "Observer and view model 'Port' setting should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.Address, This.Observer.SocketAddress, _
            "Observer 'SocketAddress' setting should equal initial setting.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.Observer.SocketAddress, _
            "Observer and view model 'SocketAddress' setting should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ReadingOffset, This.ViewModel.ReadingOffset, _
            "View Model 'ReadingOffset' setting should equal initial value.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.GpibLanControllerPort, This.ViewModel.GpibLanControllerPort, _
            "View Model 'GpibLanControllerPort' setting should equal initial value.")

    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")
        
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldInitialize") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldInitialize = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should connect. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldInitialize passed. in 11.3 ms.
''' TestShouldBeConnected passed. in 16.8 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldBeConnected() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldBeConnected"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Connected, _
            "View model should be connected.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.PrimaryGpibAddress, This.ViewModel.GpibAddress, _
            "View model Gpib address should be set.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.CloseConnectionExecutable, _
            "View model close connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.CloseConnectionExecutable, _
            "View model Observer close connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ToggleConnectionExecutable, _
            "View model Observer toggle connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.OpenConnectionExecutable, _
            "View model open connection executable should be disabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer External trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoning command should be disabled.")
        
    ' test serial polling
    If p_outcome.AssertSuccessful Then
            
        Dim p_details As String: p_details = VBA.vbNullString
        Dim p_command As String: p_command = "*CLS;*WAI;*OPC?"
        If 0 >= This.ViewModel.Session.TryWriteLine(p_command, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
    End If
    
    Dim p_serialPollDetails As String: p_serialPollDetails = VBA.vbNullString
    
    If p_outcome.AssertSuccessful And This.ViewModel.Session.GpibLanControllerAttached Then
    
        Dim p_expectedValue As Integer: p_expectedValue = cc_isr_Ieee488.ServiceRequestFlags.MessageAvailable
        Dim p_testBit As Integer: p_testBit = cc_isr_Ieee488.ServiceRequestFlags.MessageAvailable
        Dim p_statusByte As Integer
        Set p_outcome = AssertSerialPollShouldValidate(p_expectedValue, p_testBit, p_statusByte)
        p_serialPollDetails = p_outcome.AssertMessage
            
        ' set the serial poll and service request bytes
        This.ViewModel.SerialPollByte = p_statusByte
        This.ViewModel.StatusByte = p_statusByte
            
        ' get the operation completion values
        Dim p_reply As String: p_reply = VBA.vbNullString
        If 0 >= This.ViewModel.Session.TryRead(p_reply, p_details) Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail( _
                "Failed reading the operation completion reply after serial poll; " & p_details)
        
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SerialPollByte, _
                This.Observer.SerialPollByte, _
                "Observer and view model serial poll bytes should be equal.")
            
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.StatusByte, _
                This.Observer.StatusByte, _
                "Observer and view model status bytes should be equal.")
                
                        
                
    End If
    
        
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldBeConnected") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    Debug.Print VBA.vbTab & p_serialPollDetails
    
    Set TestShouldBeConnected = p_outcome
    
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

''' <summary>   Asserts that view model should read cards. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertShouldReadCards() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "AssertShouldReadCards"

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.TopCard, This.ViewModel.TopCard, _
            "View Model should be read the top card.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.BottomCard, This.ViewModel.BottomCard, _
            "View Model should be read the bottom card.")

    ' the view module initializes in continuous mode.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ContinuousSenseFunctionName, _
            This.ViewModel.SenseFunctionName, _
            "View Model should set the sense function name.")

    ' the cards are set for immediate mode.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.TopCardFunctionScanList, _
            This.ViewModel.TopCardFunctionScanList, _
            "View Model should be read the top card function scan list.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.BottomCardFunctionScanList, _
            This.ViewModel.BottomCardFunctionScanList, _
            "View Model should be read the top card function scan list.")
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

    Set AssertShouldReadCards = p_outcome
End Function

''' <summary>   Unit test. Asserts that view model should read cards. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldReadCards passed. in 13.2 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldReadCards() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldReadCards"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldReadCards") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldReadCards = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should restore initial state. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldRestoreInitialState passed. in 12234.4 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreInitialState() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldRestoreInitialState"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.

    Dim p_details As String: p_details = VBA.vbNullString

    ' check if we need to restore the GPIB-Lan initial state.
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not require restoration to inital state after connecting; " & p_details)
    End If

    Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = "VOLT:DC"
    If p_outcome.AssertSuccessful Then
        
        ' change function mode to voltage
        This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionSetter p_expectedSenseFunctionName
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedSenseFunctionName, _
            This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionName, _
            "Sense function name should be set to the expected value.")
            
    End If
    
    Dim p_actualSenseFunctionName As String
    If p_outcome.AssertSuccessful Then
        
        ' validate the actual function
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            "Actual sense function should be set to the expected value.")
            
    End If
    
    ' now that the function was changed, a resore should be required
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, _
                p_details), _
            "Restore should be required after setting the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
    
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #1; " & p_details)
            
    End If
        
    If p_outcome.AssertSuccessful Then
    
        ' once restored, restore of sense function should no longer be required
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, p_details), _
            "Restore of sense function should not be required after restoring the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring state #1; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.Session.ReadTimeoutSetter This.SessionTimeout - 1
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not be in its expected known state after setting session timeout to " & _
            VBA.CStr(This.ViewModel.Session.ReadTimeout) & " ms.")
    
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #2; " & p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring initial state #2; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.Session.AutoAssertTalkSetter True
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not be in its expected known state after setting auto assert TALK to true.")
    
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #3; " & p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring initial state #3; " & p_details)
    End If
    
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldRestoreInitialState") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRestoreInitialState = p_outcome
    
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
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldRecoverFromSyntaxError passed. in 145.6 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromSyntaxError() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromSyntaxError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    Dim p_details As String
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ClearExecutionStateCommand(p_details), _
            "View Model should clear execution state and query operation completion #1; " & p_details)
    End If

    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.ViewModel.Session.WriteLine ("**OPC")
        On Error GoTo 0
        
        
    End If

    Dim p_serialPollDetails As String: p_serialPollDetails = VBA.vbNullString

    If p_outcome.AssertSuccessful And This.ViewModel.Session.GpibLanControllerAttached Then
    
        Dim p_expectedValue As Integer: p_expectedValue = cc_isr_Ieee488.ServiceRequestFlags.ErrorAvailable
        Dim p_testBit As Integer: p_testBit = cc_isr_Ieee488.ServiceRequestFlags.ErrorAvailable
        Dim p_statusByte As Integer
        Set p_outcome = AssertSerialPollShouldValidate(p_expectedValue, p_testBit, p_statusByte)
        p_serialPollDetails = p_outcome.AssertMessage
        
        ' set the serial poll and service request bytes
        This.ViewModel.SerialPollByte = p_statusByte
        
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SerialPollByte, _
                This.Observer.SerialPollByte, _
                "Observer and view model serial poll bytes should be equal.")
            
        This.ViewModel.QueryStatusByteCommand
       
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.StatusByte, _
                This.Observer.StatusByte, _
                "Observer and view model status bytes should be equal.")
    
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
    
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ClearExecutionStateCommand(p_details), _
            "View Model should clear execution state and query operation completion #2; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldRecoverFromSyntaxError") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    Debug.Print VBA.vbTab & p_serialPollDetails
    
    Set TestShouldRecoverFromSyntaxError = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should restore its initial state from a closed connection. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldRestoreFromClosedConnection passed. in 5733.7 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreFromClosedConnection() As Assert

    Const p_procedureName As String = "TestShouldRestoreFromClosedConnection"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    Dim p_details As String
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Session.Socket.TryCloseConnection(p_details), _
            "View Model should close connection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.Device.Connected, _
            "View Model should be disconnected.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore its initial state; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Device.Connected, _
            "View Model should be connected after restoring its initial state.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.Device.QueryOperationCompleted()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should query operation completion.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldRestoreFromClosedConnection") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldRestoreFromClosedConnection = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should configure immediate mode. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConfigureImmediateMode() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConfigureImmediateMode"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' start the immediate trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertImmediateModeShouldStart(p_outcome)
    
    ' validate the immediate trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertImmediateModeShouldValidate(p_outcome)
    
    ' Assert taking a measurement
    
    Set p_outcome = AssertMeasureImmediatelyShouldReadValue(p_outcome)
    
    
    ' Finally, verify that no error message was recorded.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldConfigureImmediateMode") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldConfigureImmediateMode = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should configure external mode. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConfigureExternalMode() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConfigureExternalMode"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' start the external trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalTriggerModeShouldStart(p_outcome)
    
    ' validate the external trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalTriggerModeShouldValidate(p_outcome)
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldConfigureExternalMode") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldConfigureExternalMode = p_outcome
    
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


''' <summary>   Unit test. Asserts that view model should monitor triggering. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldMonitorTriggering() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldMonitorTriggering"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' start the external trigger mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalTriggerModeShouldStart(p_outcome)
    
    ' start the monitoring mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStart(p_outcome)
    
    ' validate the monitoring mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldValidate(p_outcome)
    
    ' get some data here.

    ' start with no monitoring
    Dim p_monitorSeconds As Double: p_monitorSeconds = 0.1
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMeasurementsShouldGetTriggered(p_outcome, p_monitorSeconds)

    ' stop monitoring here
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStop(p_outcome)
    
    ' Finally, verify that no error message was recorded.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestShouldMonitorTriggering") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldMonitorTriggering = p_outcome
    
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



















''' <summary>   Unit test. Asserts that view model should . </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestViewModelTestTemplate() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestViewModelTestTemplate"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    
    
    
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport("TestViewModelTestTemplate") & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestViewModelTestTemplate = p_outcome
    
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


