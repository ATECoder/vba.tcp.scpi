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
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As cc_isr_Test_Fx.IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
    
    ' The k2700 view
    K2700Sheet As K2700Sheet
    
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    
    ' initial observer settings
    ContinuousSenseFunctionName As String
    ImmediateSenseFunctionName As String
    ExternalSenseFunctionName As String
    
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
            Set p_outcome = TestInitialStateShouldRestore
        Case 5
            Set p_outcome = TestSyntaxErrorShouldRecover
        Case 6
            Set p_outcome = TestClosedConnectionShouldRestore
        Case 7
            Set p_outcome = TestImmediateModeShouldConfigure
        Case 8
            Set p_outcome = TestExternalModeShouldConfigure
        Case 9
            Set p_outcome = TestTriggerPollingShouldStartStop
        Case 10
            Set p_outcome = TestTriggerPollingShouldRead
        Case 11
            Set p_outcome = TestTriggerMonitoringShouldStartStop
        Case 12
            Set p_outcome = TestTriggerMonitoringShouldRead
        Case 13
            Set p_outcome = TestUserViewShouldMeasureImmediately
        Case 14
            Set p_outcome = TestUserViewMonitoringShouldStartStop
        Case 15
            Set p_outcome = TestUserViewMonitoringShouldRead
        Case 16
            Set p_outcome = TestOpenConnectionWithPowerOnResetShouldConnect
        Case Else
    End Select
    AfterEach
    Set RunTest = p_outcome
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 16
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' Test 01 TestShouldInitialize passed. Elapsed time: 4609.5 ms.
''' Test 02 TestShouldBeConnected passed. Elapsed time: 34.2 ms.
'''     Serial Poll is 16 in 8.6 ms.
''' Test 03 TestShouldReadCards passed. Elapsed time: 10.8 ms.
''' Test 04 TestInitialStateShouldRestore passed. Elapsed time: 13155.9 ms.
''' Test 05 TestSyntaxErrorShouldRecover passed. Elapsed time: 166.4 ms.
'''     Serial Poll is 4 in 4.9 ms.
''' Test 06 TestClosedConnectionShouldRestore passed. Elapsed time: 7647.8 ms.
''' Test 07 TestImmediateModeShouldConfigure passed. Elapsed time: 2396.0 ms.
''' Test 08 TestExternalModeShouldConfigure passed. Elapsed time: 2474.9 ms.
''' Test 09 TestTriggerPollingShouldStartStop passed. Elapsed time: 2155.3 ms.
''' Awaiting triggers...
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00136353E+02'.
'''  1 : 100.136353
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00134216E+02'.
'''  2 : 100.134216
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00134689E+02'.
'''  3 : 100.134689
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00135223E+02'.
'''  4 : 100.135223
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00134422E+02'.
'''  5 : 100.134422
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00135017E+02'.
'''  6 : 100.135017
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00135025E+02'.
'''  7 : 100.135025
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00133934E+02'.
'''  8 : 100.133934
''' Test 10 TestTriggerPollingShouldRead passed. Elapsed time: 9017.2 ms.
''' Test 11 TestTriggerMonitoringShouldStartStop passed. Elapsed time: 6234.8 ms.
''' Waiting for trigger....
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00134773E+02'.
'''  1 : 100.134773
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00135361E+02'.
'''  2 : 100.135361
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00135979E+02'.
'''  3 : 100.135979
''' Test 12 TestTriggerMonitoringShouldRead passed. Elapsed time: 11310.5 ms.
'''  1 : 100.146591
''' 19:08:24 Power on reset starting. This could take 3 seconds. Please wait...
''' 19:08:31 done power on reset.
''' Test 13 TestOpenConnectionWithPowerOnResetShouldConnect passed. Elapsed time: 7321.8 ms.
''' Ran 13 out of 13 tests.
''' Passed: 13; Failed: 0; Inconclusive: 0.
''' </code>
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 13
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
        VBA.DoEvents
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

    ' initialize the current and previous test numbers.
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")

    This.Name = "K2700ViewModelTests"
    
    ' initialize test settings
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    
    ' initialize known data.
    This.TopCard = "7700"
    This.BottomCard = VBA.vbNullString
    This.ContinuousSenseFunctionName = "FRES"
    This.ImmediateSenseFunctionName = "RES"
    This.ExternalSenseFunctionName = "FRES"
    
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
    
    ' initialize the observer before initializing the view mode
    ' but after the observer setting are set. The observer initial
    ' settings are then applied to the view model.
    Set This.Observer = K2700Observer.Initialize(This.ViewModel)
    Dim a_dataSheet As DataSheet
    Set a_dataSheet = DataSheet.Initialize(This.ViewModel)
    Set This.DataView = DataView.Instance
    Set This.UserView = UserView.Instance
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreNotEqual(0, This.DataView.GpibLanControllerPort, _
            "Data view GPIB Lan Controller Port must be non-zero.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.DataView.GpibLanControllerPort, _
            This.ViewModel.Session.GpibLanControllerPort, _
            "Data view and Session should define the same GPIB Lan Controller Port.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreNotEqual(VBA.vbNullString, _
            This.DataView.ChannelNumberCaptionPrefix, _
            "Data view channel number caption prefix must not be empty.")
    
    ' issue the open connection command. This initializes the view model.
    If p_outcome.AssertSuccessful Then _
        This.ViewModel.OpenConnectionCommand This.DataView.SocketAddress, This.DataView.SessionTimeout
    
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

    ' increment the test number if running under the test executive.
    If This.TestNumber = This.PreviousTestNumber Then This.TestNumber = This.PreviousTestNumber + 1
    
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

    ' set the previous test number to the current test number.
    This.PreviousTestNumber = This.TestNumber
    
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
        Dim p_details As String
        This.ViewModel.ResetKnownStateCommand p_details
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
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertExternalModeShouldConfigure(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = p_outcome.AssertSuccessful
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        p_success = This.ViewModel.ConfigureMeasureCommand(a_mode, p_details)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.FrontInputsHasValue, _
            "View model front inputs should be validated.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsRequired, _
            This.ViewModel.FrontInputsValue, _
            "View model front input value should equal the required value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsValue, _
            This.Observer.FrontInputsValue, _
            "Observer Front inputs state should equal view model inputs state for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsRequired, _
            This.UserView.FrontInputsRequired, _
            "User View Front inputs state should equal view model inputs state for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "External trigger reading mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.DataView.MeasurementMode, _
            "Data acquisition view measurement mode should equal expected value for external trigger reading mode.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for external trigger reading mode.")
    End If
    
    Set AssertExternalModeShouldConfigure = p_outcome

End Function

''' summary>   Asserts that external trigger mode should be validated. </summary>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertExternalModeShouldValidate(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with test validations.
    
    If p_outcome.AssertSuccessful Then
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "External trigger reading mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for external trigger reading mode.")

    ' testing trigger monitoring uses auto increment to detect changes
    ' in channel number as readings are triggered.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_mode.AutoIncrementChannel, _
            "Auto increment channel number should be true for testing external trigger monitoring.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.AutoIncrementChannel, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "View Model Auto Increment Channel No Enabled should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.AutoIncrementChannelNoEnabled, _
            This.UserView.AutoIncrementChannelNoEnabled, _
            "User View and View Model Auto Increment Channel No Enabled should equal.")
   
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.SingleRead, _
            This.ViewModel.SingleReadEnabled, _
            "View Model Single Read Enabled should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SingleReadEnabled, _
            This.UserView.SingleReadEnabled, _
            "User View and View Model Single read enabeld should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.SenseFunction, _
            This.ViewModel.SenseFunctionName, _
            "View Model Sense channel should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = This.ExternalSenseFunctionName
        Dim p_actualSenseFunctionName As String
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqualString(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            VBA.VbCompareMethod.vbTextCompare, _
            "External mode sense function name should be as expected.")
    End If
 
    If p_outcome.AssertSuccessful Then _
       Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.TargetChannelNumber, This.Observer.TargetChannelNumber, _
            "Observer Target Channel Number should equal the view model channel number.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSourceOption.External, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "External trigger source should be as expected.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.K2700.SenseSystem.SenseSystem.AutoRangeEnabledGetter(), _
            "Auto range should be enabled.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1#, _
            This.ViewModel.K2700.SenseSystem.SenseSystem.PowerLineCyclesGetter(), _
            "The integration rate in power line cycles should be as expected.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger initiation should be off in external trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.PauseRequested, _
            "Pause requested should be on in external trigger reading mode before monitoring started.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop requested should be on in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
            "User View immediate scan button should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
            "User View immediate single button should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be disabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be enabled in external trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                "User View manual scan button should be disabled in external trigger single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                "User View manual scan button should be disabled in external trigger multi-reading mode.")
        End If
    End If
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                "User View manual single button should be enabled in external trigger single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                "User View manual single button should be disabled in external trigger multi-reading mode.")
        End If
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be enabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off in external trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off in external trigger reading mode.")
    
    Set AssertExternalModeShouldValidate = p_outcome

End Function

''' summary>   Asserts that immediate trigger mode should be configured. </summary>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">          [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertImmediateModeShouldConfigure(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = p_outcome.AssertSuccessful
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        p_success = This.ViewModel.ConfigureMeasureCommand(a_mode, p_details)
        
        ' returns true of if success. Otherwise, the error should be in the
        ' last error, if the inputs are invalid or the last error message
        ' if the configuration failed.
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
        
    End If
    
    ' the card scan lists are set when entring the immediate mode.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.TopCardFunctionScanList, _
            This.ViewModel.TopCardFunctionScanList, _
            "View Model should be read the top card function scan list.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.BottomCardFunctionScanList, _
            This.ViewModel.BottomCardFunctionScanList, _
            "View Model should be read the bottom card function scan list.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.ChannelNumber, _
            This.ViewModel.SelectedChannelNumber, _
            "View model selected channelnumber should equal the expected value.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.FrontInputsHasValue, _
            "View model front inputs should be validated.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.FrontInputs, _
            This.ViewModel.FrontInputsRequired, _
            "View model required front inputs should equal the expected value.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsRequired, _
            This.ViewModel.FrontInputsValue, _
            "View model front input value should equal the required value.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsValue, _
            This.Observer.FrontInputsValue, _
            "Observer Front inputs state should equal view model inputs state for immediate trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsRequired, _
            This.UserView.FrontInputsRequired, _
            "User View Front inputs state should equal view model inputs state for immediate trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "Immediate measurement mode should be as expected when starting.")
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for immediate trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for immediate trigger reading mode.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for immediate trigger reading mode.")
    
    Set AssertImmediateModeShouldConfigure = p_outcome

End Function


''' summary>   Asserts that immediate trigger mode should be be validated. </summary>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertImmediateModeShouldValidate(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' proceed with validations
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for immediate trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for immediate trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for immediate trigger reading mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(a_mode.AutoIncrementChannel, _
            "Auto increment channel number should be False for testing immeidate trigger monitoring.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.AutoIncrementChannel, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "View Model Auto Increment Channel No Enabled should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.AutoIncrementChannelNoEnabled, _
            This.UserView.AutoIncrementChannelNoEnabled, _
            "User View and View Model Auto Increment Channel No Enabled should equal.")
   
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.SingleRead, _
            This.ViewModel.SingleReadEnabled, _
            "View Model Single Read Enabled should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SingleReadEnabled, _
            This.UserView.SingleReadEnabled, _
            "User View and View Model Single read enabeld should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.SenseFunction, _
            This.ViewModel.SenseFunctionName, _
            "View Model Sense channel should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSourceOption.Immediate, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "Immediate trigger source should be as expected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop requested should be on in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.MeasureExecutable, _
            "Measure command should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.MeasureExecutable, _
            "Observer Measure button should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
            "User View manual scan button should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
            "User View manual single button should be disabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
                "User View auto scan button should be disabled in immediate single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                "User View auto scan button should be enabled in immediate multi-reading mode.")
        End If
    End If
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                "User View auto single button should be enabled in immediate single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
                "User View auto single button should be disabled in immediate multi-reading mode.")
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off in immediate mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.AutoIncrementChannel, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    ' with immediate mode and single reading, the selected channel is used to set the
    ' measured channel after a reading is triggered and the measurement event is handled
    ' by the observer.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.ChannelNumber, This.ViewModel.SelectedChannelNumber, _
            "The View Model selected channel number should equal the expected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.SelectedChannelNumber > 0, _
            "The View Model selected channel number '" & VBA.CStr(This.ViewModel.SelectedChannelNumber) & _
            "' should be positive.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.SelectedChannelNumber <= This.ViewModel.ChannelCount, _
            "The View Model selected channel number '" & VBA.CStr(This.ViewModel.SelectedChannelNumber) & _
            "' should be smaller or equal the channel count '" & VBA.CStr(This.ViewModel.ChannelCount) & ".")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.Observer.SelectedChannelNumber, _
            "The Observer selected channel number should be set to the View Model selected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.UserView.SelectedChannelNumber, _
            "The User View selected channel number should be set to the View Model selected channel number.")
    
    Set AssertImmediateModeShouldValidate = p_outcome

End Function

''' summary>   Returns the expected target channel number given the measured channel number. </summary>
''' <returns>   [Integer]. </returns>
Public Function ExpectedTargetChannelNumber() As Integer
    
    If This.ViewModel.AutoIncrementChannelNoEnabled Then

        ' with multiple measurement, the target channel number increments after the measurement is made
        ExpectedTargetChannelNumber = IIf(This.ViewModel.MeasuredChannelNumber < This.ViewModel.ChannelCount, _
                This.ViewModel.MeasuredChannelNumber + 1, 1)
    Else
    
        ' with single measurements, the channel number is the selected channel number.
        ExpectedTargetChannelNumber = This.ViewModel.SelectedChannelNumber
    
    End If
    

End Function

''' summary>   Asserts that immediate measurement should take a reading. </summary>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMeasureImmediatelyShouldReadValue(ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    
    ' make sure we are in immediate trigger mode.
    
    If p_outcome.AssertSuccessful Then
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    ' immediate mode is tested with single measurements. Auto increment is off
    ' and the measured channel is the selected channel
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.SelectedChannelNumber > 0, _
            "The selected channel number for immediate measurement should be positive.")
    
    ' proceed with test assertions.
    
    ' take a reading
    If p_outcome.AssertSuccessful Then _
        p_success = This.ViewModel.MeasureImmediatelyCommand(This.UserView.ReadingOffset, p_details)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, p_details)
    
    If p_outcome.AssertSuccessful Then
        
        ' wait for the reading event to take shape.
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.Observer.MeasuredChannelNumber, _
            "Observer measured channel number should equal the selected channel number.")
            
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.UserView.SelectedChannelNumber, _
            "The User View selected channel number should be set to the View Model selected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.DataView.MeasuredChannelNumber, _
            "The Data View measured channel number should be set to the View Model selected channel number.")
    
    Dim p_reading As String
    Dim p_channelNumber As Integer
    Dim p_readingValue As Double
    
    If p_outcome.AssertSuccessful Then
        
        ' get the reading from the observer.
        p_reading = This.DataView.MeasuredReading
        
        p_channelNumber = This.DataView.MeasuredChannelNumber

        p_readingValue = This.DataView.MeasuredValue
        
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.DataView.MeasuredChannelNumber, _
            This.ViewModel.MeasuredChannelNumber, _
            "View Model measured channel number should equal the Data View measured channel.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.Observer.MeasuredChannelNumber, _
            This.ViewModel.MeasuredChannelNumber, _
            "View Model measured channel number should equal the Observer measured channel.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(ExpectedTargetChannelNumber(), _
            This.ViewModel.SelectedChannelNumber, _
            "The expected target channel number should equal the selected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.Observer.SelectedChannelNumber, _
            "The observer Selected Channel Number should equal the view model selected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, _
            This.UserView.SelectedChannelNumber, _
            "The User View Selected Channel Number should equal the view model selected channel number.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(VBA.vbNullString = p_reading, _
            "Reading should not be empty.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_readingValue > 0, _
            "Reading value should be positive.")
    
    Dim p_epsilon As Double: p_epsilon = 0.0000000001
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreCloseDouble(p_readingValue, VBA.CDbl(p_reading), p_epsilon, _
            "Reading should equal the parsed value.")
    
    Set AssertMeasureImmediatelyShouldReadValue = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode should be configured. </summary>
''' <param name="a_timerInteger">   [Integer] the timer interval; 0 for poll monitoring of the timer event. </value>
''' <param name="a_assert">         [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldStart(ByVal a_timerInterval As Integer, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' check that the external trigger mode was set.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "External Trigger mode should be as expected.")
    End If
    
    ' start monitoring here.
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.StartMonitoringExternalTriggers This.UserView.ReadingOffset, a_timerInterval
        
        ' allow the monitoring to commence.
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_timerInterval, This.ViewModel.TimerInterval, _
            "Timer interval should expected the expected value.")
    
    If p_outcome.AssertSuccessful And (0 = a_timerInterval) Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_timerInterval, This.ViewModel.TimerInterval, _
            "Timer interval should expected the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "External trigger monitoring mode should be as expected.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for  trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for  trigger monitoring mode.")
    
    Set AssertMonitoringModeShouldStart = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode should be validated. </summary>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldValidate(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' start validating.
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "External trigger monitoring mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.Observer.MeasurementMode, _
            "Observer measurement mode should equal expected value for external trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View measurement mode should equal expected value for external trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View measurement mode should equal expected value for external trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.PauseRequested, _
            "Pause Requested should be off after starting the monitoring.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopRequested, _
            "Stop requested should be off in trigger monitoring mode.")
    
    ' the external trigger is initiated immediate when the timer is started in timer control
    If p_outcome.AssertSuccessful And This.ViewModel.TimerControlled Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger should get initiated after starting the monitoring timer.")
    
    If p_outcome.AssertSuccessful And This.ViewModel.TimerControlled Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TimerStarted, _
            "Timer started should be True after monitoring started under timer control.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be enabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be enabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
            "User View auto scan button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
            "User View auto single button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                "User View Manual scan button should be disabled in trigger monitoring single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                "User View Manual scan button should be enabled in trigger monitoring multi-reading mode.")
        End If
    End If
    
    If p_outcome.AssertSuccessful Then
        If This.UserView.SingleReadEnabled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                "User View Manual single button should be enabled in trigger monitoring single-reading mode.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                "User View Manual single button should be disabled in trigger monitoring multi-reading mode.")
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be disabled in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be on in trigger monitoring mode.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be on in trigger monitoring mode.")
    
    ' testing trigger monitoring uses auto increment to detect changes
    ' in channel number as readings are triggered.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_mode.AutoIncrementChannel, _
            "Auto increment channel number should be true for testing external trigger monitoring.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.AutoIncrementChannel, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    ' with triggered mode and multiple reading, the Target Channel Number is used to set the
    ' measured channel after a reading is triggered and the measurement event is handled
    ' by the observer. The target channel number must then be set to between 1 and the
    ' channel count (see below).
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_mode.ChannelNumber, This.ViewModel.TargetChannelNumber, _
            "The View Model Target channel number should equals the settings channel number.")
    
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TargetChannelNumber > 0, _
            "The View Model Target channel number '" & VBA.CStr(This.ViewModel.TargetChannelNumber) & _
            "' should be positive.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TargetChannelNumber <= This.ViewModel.ChannelCount, _
            "The View Model Target channel number '" & VBA.CStr(This.ViewModel.TargetChannelNumber) & _
            "' should be smaller or equal the channel count '" & VBA.CStr(This.ViewModel.ChannelCount) & ".")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.TargetChannelNumber, _
            This.Observer.TargetChannelNumber, _
            "Observer Target Channel Number should be set to the selected channel number.")
   
    Set AssertMonitoringModeShouldValidate = p_outcome

End Function

''' summary>   Asserts that measurements should get triggered. </summary>
''' <param name="a_assert">     [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <param name="a_duration">   [Optional, Double, 30] The time to wait for some triggered values. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMeasurementsShouldGetTriggered(ByVal a_assert As cc_isr_Test_Fx.Assert, _
    Optional ByVal a_duration As Double = 30) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    
    ' Auto increment (multiple readings) is used in triggered measurement tests in which
    ' case, the target channel number is measured. Following each reading, the target channel
    ' number is incremented in a circular fashion.
    
    ' get the first channel number
    Dim p_channel As Integer
    p_channel = This.DataView.MeasuredChannelNumber
    
    Dim p_reading As String
    p_reading = This.DataView.MeasuredReading
    
    VBA.DoEvents
    Debug.Print "Waiting for trigger...."
    
    ' loop for some time waiting for triggered measurements.
    
    Dim p_endTime As Double
    p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow() + _
        (a_duration / cc_isr_Core_IO.CoreExtensions.SecondsPerDay)
    While p_endTime > cc_isr_Core_IO.CoreExtensions.DaysNow()
        
        VBA.DoEvents
    
        If p_channel <> This.DataView.MeasuredChannelNumber Then
        
            VBA.DoEvents
            p_channel = This.DataView.MeasuredChannelNumber
            
            VBA.DoEvents
            p_reading = This.DataView.MeasuredReading
            
            VBA.DoEvents
            Debug.Print p_channel; ": "; p_reading
            
            ' verify that measured channel numbers propagated correctly.
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.Observer.MeasuredChannelNumber, _
                    This.ViewModel.MeasuredChannelNumber, _
                    "View Model measured channel number should equal the Observer measured channel.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.DataView.MeasuredChannelNumber, _
                    This.ViewModel.MeasuredChannelNumber, _
                    "View Model measured channel number should equal the Data View measured channel.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(ExpectedTargetChannelNumber(), _
                    This.ViewModel.TargetChannelNumber, _
                    "The target channel number should equal the expected target channel number after a triggered reading.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.TargetChannelNumber, _
                    This.Observer.TargetChannelNumber, _
                    "The observer Target Channel Number should equal the view model target channel number.")

        End If
    
    Wend
    
    Set AssertMeasurementsShouldGetTriggered = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode should be stopped. </summary>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeShouldStop(ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' verify that we stop from an active monitoring mode.
    
    ' stop monitoring here
    
    If p_outcome.AssertSuccessful Then
    
        ' monitoring might have been stopped already.
        
        If This.ViewModel.MeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring Then
        
            This.ViewModel.StopMonitoringExternalTriggersCommand
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.StopRequested, _
                "Stop Requested should be on off after stopping monitoring.")
            
        End If
    
    End If
    
    ' allow time for monitoring to stop
    
    If p_outcome.AssertSuccessful Then _
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.PauseRequested, _
            "Pause should be requested after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.TimerStarted, _
            "Timer started should be false after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.MeasurementModeOption.None, This.ViewModel.MeasurementMode, _
            "Measurement mode should be as expected after monitoring stopped.")
    
    Set AssertMonitoringModeShouldStop = p_outcome

End Function

''' summary>   Asserts that trigger monitoring mode stop should be validated. </summary>
''' <param name="a_assert">   [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertMonitoringModeStopShouldValidate(ByVal a_assert As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' validate monitoring stop state
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.None
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.MeasurementMode, _
            "Measurement mode should be as expected after monitoring stopped.")
    End If
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "The stop monitoring executable to should disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External monitoring enabled should be off after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger should not get initiated after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StopMonitoringExecutable, _
            "Observer stop monitoring button should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.MeasureExecutable, _
            "Measure command should be disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.MeasureExecutable, _
            "Observer Measure button should be disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
            "User View auto scan command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
            "User View auto single command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
            "User View manual scan command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
            "User View manual single command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.StartMonitoringExecutable, _
            "Observer start monitoring button should be disabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ImmediateTriggerOptionExecutable, _
            "Observer immediate trigger option button should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option command should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer external trigger option button should be enabled after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External trigger monitoring state should be off after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.Observer.ExternalTrigMonitoringEnabled, _
            "Observer external trigger monitoring state should be off after monitoring stopped.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.TargetChannelNumber, This.Observer.TargetChannelNumber, _
            "Observer Target Channel Number should be set to the selected channel number.")
    
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
    Dim p_details As String
    
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
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.Observer.SocketAddress, _
            "Observer 'Socket Address' setting should equal the view model value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.DataView.SocketAddress, _
            "Data View 'Socket Address' setting should equal the view model initial setting.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.Observer.SocketAddress, _
            "Observer and view model 'Socket Address' setting should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.DataView.SocketAddress, _
            "Data View and view model 'Socket Address' setting should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(DataSheet.GpibLanControllerPort, This.ViewModel.GpibLanControllerPort, _
            "View Model 'GpibLanControllerPort' setting should equal data sheet value.")
            
    ' check the Data View Status
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Connected, _
            "View Model Connected state should be true.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.OpenConnectionExecutable, _
            "View Model Open Connection Executable should be False.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.CloseConnectionExecutable, _
            "View Model Close Connection Executable should be True.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.OpenConnectionExecutable, _
            This.DataView.OpenConnectionExecutable, _
            "Data View and View Model Open Connection Executables should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CloseConnectionExecutable, _
            This.DataView.CloseConnectionExecutable, _
            "Data View and View Model Close Connection Executables should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SocketAddress, This.DataView.SocketAddress, _
            "Data View and View Model Socket Addresses should equal.")
   
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SessionTimeout, This.DataView.SessionTimeout, _
            "Data View and View Model Session Timeouts should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty.")

    This.ViewModel.OnError "test: no error"
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastErrorMessage, This.DataView.LastErrorMessage, _
            "Data View and View Model Last Error Messages should equal.")
    This.ViewModel.OnError VBA.vbNullString

    Dim p_errorMessage As String: p_errorMessage = "last error message"
    Dim p_message As String: p_message = "last Message"
    This.ViewModel.ClearMessages p_errorMessage, p_message
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_message, This.ViewModel.LastMessage, _
            "View model last message should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_errorMessage, This.ViewModel.LastErrorMessage, _
            "View model last error message should equal the expected value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastMessage, This.DataView.LastMessage, _
            "Data View and View Model Last Messages should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastErrorMessage, This.DataView.LastErrorMessage, _
            "Data View and View Model Last error Messages should equal.")

    p_errorMessage = VBA.vbNullString
    p_message = VBA.vbNullString
    This.ViewModel.ClearMessages

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_message, This.ViewModel.LastMessage, _
            "View model last message should equal the expected value after clear.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_errorMessage, This.ViewModel.LastErrorMessage, _
            "View model last error message should equal the expected value after clear.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastMessage, This.DataView.LastMessage, _
            "Data View and View Model Last Messages should equal after clear.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastErrorMessage, This.DataView.LastErrorMessage, _
            "Data View and View Model Last error Messages should equal after clear.")


    Dim p_expectedChannelNumber As Integer
    Dim p_expectedReading As String
    Dim p_expectedMeasuredValue As Double
    
    p_expectedChannelNumber = 0
    p_expectedReading = VBA.vbNullString
    p_expectedMeasuredValue = 0#

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedChannelNumber, This.ViewModel.MeasuredChannelNumber, _
            "View Model Measured Channel Number should be expected before taking a measurement.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReading, This.ViewModel.MeasuredReading, _
            "Measured Reading should be expected before taking a measurement.")

    ' emulate a channel reading
    p_expectedChannelNumber = 1
    p_expectedReading = "0.0"
    p_expectedMeasuredValue = 0#
    
    This.ViewModel.K2700.OnChannelMeasured p_expectedChannelNumber, p_expectedReading
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedChannelNumber, This.ViewModel.MeasuredChannelNumber, _
            "View Model Measured Channel Number should be expected after emulating a measurement.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReading, This.ViewModel.MeasuredReading, _
            "Measured Reading should be expected after emulating a measurement.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasuredChannelNumber, This.Observer.MeasuredChannelNumber, _
            "Observer and View Model Measured Channel Numbers should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasuredChannelNumber, This.DataView.MeasuredChannelNumber, _
            "Data View and View Model Measured Channel Numbers should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasuredReading, This.DataView.MeasuredReading, _
            "Data View and View Model Measured Readings should equal.")
    
    p_expectedReading = VBA.vbNullString
    This.ViewModel.K2700.OnChannelMeasured p_expectedChannelNumber, p_expectedReading
    p_expectedMeasuredValue = cc_isr_Ieee488.Syntax.NotANumber
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReading, This.ViewModel.MeasuredReading, _
            "Measured Reading should be expected after emulating a failed measurement.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasuredValue, This.ViewModel.MeasuredValue, _
            "Measured Value should be expected after emulating a failed measurement.")

    p_expectedReading = "101"
    This.ViewModel.K2700.OnChannelMeasured p_expectedChannelNumber, p_expectedReading
    p_expectedMeasuredValue = 101
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReading, This.ViewModel.MeasuredReading, _
            "Measured Reading should be expected after emulating a good measurement.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedMeasuredValue, This.ViewModel.MeasuredValue, _
            "Measured Value should be expected after emulating a good measurement.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous, _
            This.ViewModel.MeasurementMode, _
            "Measurement Model should be continuous.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.DataView.MeasurementMode, _
            "Data View and View Model Measurement Modes should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ClearReadingsExecutable, _
            "View Model Clear Reading executable should be true.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.ClearReadingsExecutable, This.DataView.ClearReadingsExecutable, _
            "Data View and View Model Clear Readings Executables should equal.")

    ' check the User View Status
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Connected, This.UserView.Connected, _
            "Data View and View Model connection state should equal when connected.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.FrontInputsRequired, This.UserView.FrontInputsRequired, _
            "User View and View Model Front Inputs Required should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, This.ViewModel.SelectedChannelNumber, _
            "View Model Selected Channel Numbers should equal 1.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SelectedChannelNumber, This.UserView.SelectedChannelNumber, _
            "User View and View Model Selected Channel Numbers should equal.")

    p_expectedChannelNumber = 2
    This.UserView.SelectedChannelNumber = p_expectedChannelNumber
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedChannelNumber, This.UserView.SelectedChannelNumber, _
            "Data View Selected Channel Numbers should equal after expected value.")
    
    p_expectedChannelNumber = 1
    This.UserView.SelectedChannelNumber = p_expectedChannelNumber
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedChannelNumber, This.UserView.SelectedChannelNumber, _
            "Data View Selected Channel Numbers should equal after expected value after restoring value.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.AutoIncrementChannelNoEnabled, _
            This.UserView.AutoIncrementChannelNoEnabled, _
            "Data View and View Model Auto Increment Channel No Enabled should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SingleReadEnabled, _
            This.UserView.SingleReadEnabled, _
            "Data View and View Model Single Read Enabled should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, _
            This.UserView.MeasurementMode, _
            "User View and View Model Measurement Modes should equal.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.Measuring, _
            "Measuring state should be false.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Measuring, _
            This.UserView.Measuring, _
            "User View and View Model Measuring state should equal.")

    ' check how measurement mode changes the values
    
    If p_outcome.AssertSuccessful Then
        'On Error Resume Next
        ' get into design mode.
        ' these cause error 91 although these work from the
        ' immediate window.
        ' CommandBars("Exit Design Mode").Controls(1).Execute
        'If CommandBars.GetEnabledMso("DesignMode") Then _
        '    CommandBars.ExecuteMso "DesignMode"
        'On Error GoTo exit_Handler:
        This.UserView.DesignMode = True
    End If
    
    This.ViewModel.MeasurementModeUnitTestSetter cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View and View Model Measurement Modes should equal.")
    
    Dim p_measurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    Dim p_info As String
    Dim p_measuring As Boolean
    Dim p_singleRead As Boolean
    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To 4
        If i = 1 Then
            p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous
            p_info = "continuous"
        ElseIf i = 2 Then
            p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
            p_info = "immediate"
        ElseIf i = 3 Then
            p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
            p_info = "external"
        ElseIf i = 4 Then
            p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
            p_info = "monitoring"
        End If
        
        For j = 1 To 2
            If j = 1 Then
                p_singleRead = False
            Else
                p_singleRead = True
            End If
            
            If p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous Or _
                p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.None Then
                p_measuring = False
            
                This.ViewModel.OnMeasurementStateChanged p_measurementMode, p_info, p_measuring, p_singleRead
                
                If p_outcome.AssertSuccessful Then _
                   Set p_outcome = AssertUserInterfaceState(p_outcome)
                   
            Else
            
                For k = 1 To 2
                    If k = 1 Then
                       p_measuring = False
                       p_info = p_info & " done"
                    Else
                       p_measuring = True
                       p_info = p_info & " started"
                    End If
                    
                    This.ViewModel.OnMeasurementStateChanged p_measurementMode, p_info, p_measuring, p_singleRead
                    
                    If p_outcome.AssertSuccessful Then _
                       Set p_outcome = AssertUserInterfaceState(p_outcome)
                       
                Next
            
            End If
        
        Next
        
    Next
    
    ' make sure to restore the none measurement mode.
    p_measuring = False
    p_info = "none"
    p_measurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.None
    This.ViewModel.OnMeasurementStateChanged p_measurementMode, p_info, p_measuring, True
    If p_outcome.AssertSuccessful Then _
       Set p_outcome = AssertUserInterfaceState(p_outcome)
    
    'On Error Resume Next
    ' exit design mode
    ' CommandBars.GetPressedMso ("DesignMode")
    'CommandBars("Exit Design Mode").Controls(1).Reset
    'On Error GoTo exit_Handler:
    This.UserView.DesignMode = False

    ' close connection and check status of user interface.
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Session.Socket.TryCloseConnection(p_details), _
            "View Model should close connection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.Device.Connected, _
            "View Model device should be disconnected.")
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.OpenConnectionExecutable, _
            This.DataView.OpenConnectionExecutable, _
            "Data View and View Model Open Connection Executables should equal after disconnection.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.CloseConnectionExecutable, _
            This.DataView.CloseConnectionExecutable, _
            "Data View and View Model Close Connection Executables should equal after disconnection.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Connected, This.UserView.Connected, _
            "User View and View Model Connected states should equal after disconnection.")

    If p_outcome.AssertSuccessful Then _
       Set p_outcome = AssertUserInterfaceState(p_outcome)

    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.OpenConnectionCommand This.DataView.SocketAddress, This.DataView.SessionTimeout
    
        If Not This.ViewModel.Connected Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail("Failed reconnecting after initialize.")

    End If

    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
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


Public Function AssertUserInterfaceState(ByVal a_outcome As cc_isr_Test_Fx.Assert) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = a_outcome
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.MeasurementMode, This.UserView.MeasurementMode, _
            "User View and View Model Measurement Modes should equal for testing user interface controls.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Connected, This.UserView.Connected, _
            "User View and View Model Connected states should equal for testing user interface controls.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.SingleReadEnabled, This.UserView.SingleReadEnabled, _
            "User View and View Model single read enabled should equal for testing user interface controls.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.Measuring, This.UserView.Measuring, _
            "User View and View Model Measuring states should equal for testing user interface controls.")
    
    If Not This.ViewModel.Connected Then
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
                "User view Auto Scan Toggle should not be executable when connected.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
                "User view Auto Single Toggle should not be executable when connected.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                "User view Manual Scan Toggle should not be executable when connected.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                "User view Manual Single Toggle should not be executable when connected.")
    
    ElseIf cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous = This.ViewModel.MeasurementMode Or _
           cc_isr_Tcp_Scpi.MeasurementModeOption.None = This.ViewModel.MeasurementMode Then

        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                "User view Auto Scan Toggle should be executable where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                "User view Auto Single Toggle should be executable where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                "User view Manual Scan Toggle should be executable where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                "User view Manual Single Toggle should be executable where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                "User view Auto Scan Toggle should be released (false) where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                "User view Auto Single Toggle should be released (false) where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                "User view Manual Scan Toggle should be released (false) where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
    
        If p_outcome.AssertSuccessful Then _
            Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                "User view Manual Single Toggle should be released (false) where " & _
                "Measurement Mode is Continuous, Single Read, and Measuring.")
        
    ElseIf This.ViewModel.Measuring Then
    
        If This.ViewModel.SingleReadEnabled Then
        
            Select Case This.ViewModel.MeasurementMode
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be pressed (true) where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and Measuring.")
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, cc_isr_Tcp_Scpi.MeasurementModeOption.External
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be pressed (true) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and Measuring.")
                
            End Select
            
        Else
        
            Select Case This.ViewModel.MeasurementMode
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should not be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be pressed (true) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and Measuring.")
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, cc_isr_Tcp_Scpi.MeasurementModeOption.External
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should not be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be pressed (true) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and Measuring.")
                
            End Select
        
        End If
    
    Else
    
        If This.ViewModel.SingleReadEnabled Then
        
            Select Case This.ViewModel.MeasurementMode
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Single Read, and not Measuring.")
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, cc_isr_Tcp_Scpi.MeasurementModeOption.External
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Single Read, and not Measuring.")
                
            End Select
            
        Else
        
            Select Case This.ViewModel.MeasurementMode
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should be executable where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is Immediate, Multi-Read, and not Measuring.")
                
                Case cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, cc_isr_Tcp_Scpi.MeasurementModeOption.External
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
                            "User view Auto Scan Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
                            "User view Auto Single Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
                            "User view Manual Scan Toggle should executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
                            "User view Manual Single Toggle should be executable where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoScanToggleValue, _
                            "User view Auto Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.AutoSingleToggleValue, _
                            "User view Auto Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualScanToggleValue, _
                            "User view Manual Scan Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
                    If p_outcome.AssertSuccessful Then _
                        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.UserView.ManualSingleToggleValue, _
                            "User view Manual Single Toggle should be released (false) where " & _
                            "Measurement Mode is external or monitoring, Multi-Read, and not Measuring.")
                
            End Select
        
        End If
    
    End If
    
    Set AssertUserInterfaceState = p_outcome
    
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
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(DataView.PrimaryGpibAddress, This.ViewModel.GpibAddress, _
            "View model Gpib address should be set to the user view value.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.CloseConnectionExecutable, _
            "View model close connection executable should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.CloseConnectionExecutable, _
            "View model Observer close connection executable should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ToggleConnectionExecutable, _
            "View model Observer toggle connection executable should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.OpenConnectionExecutable, _
            "View model open connection executable should be disabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer External trigger option button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoring command should be disabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoScanToggleExecutable, _
            "User View auto scan button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.AutoSingleToggleExecutable, _
            "User View auto Single button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualScanToggleExecutable, _
            "User View Manual scan button should be enabled upon connection.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.UserView.ManualSingleToggleExecutable, _
            "User View Manual Single button should be enabled upon connection.")
        
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
        'This.ViewModel.SerialPollByte = p_statusByte
        'This.ViewModel.StatusByte = p_statusByte
            
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
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.LastErrorMessage, _
            This.DataView.LastErrorMessage, _
            "Data View Last error message should be the same as the view model.")
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
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
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
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
''' TestInitialStateShouldRestore passed. in 12234.4 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestInitialStateShouldRestore() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestInitialStateShouldRestore"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    
    ' proceed with test assertions.

    Dim p_details As String: p_details = VBA.vbNullString

    ' check if we need to restore the GPIB-Lan initial state.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not require restoration to initial state after connecting; " & p_details)

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
    
    ' now that the function was changed, a restore should be required
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, _
                p_details), _
            "Restore should be required after setting the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #1; " & p_details)
            
    If p_outcome.AssertSuccessful Then
        ' once restored, restore of sense function should no longer be required
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, p_details), _
            "Restore of sense function should not be required after restoring the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
    End If
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring state #1; " & p_details)
    
    If p_outcome.AssertSuccessful Then
        This.ViewModel.Session.ReadTimeoutSetter This.ViewModel.Session.ReadTimeout - 1
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not be in its expected known state after setting session timeout to " & _
            VBA.CStr(This.ViewModel.Session.ReadTimeout) & " ms.")
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #2; " & p_details)
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring initial state #2; " & p_details)
            
    If p_outcome.AssertSuccessful Then
        This.ViewModel.Session.AutoAssertTalkSetter True
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should not be in its expected known state after setting auto assert TALK to true.")
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore initial state #3; " & p_details)
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.ShouldRestoreInitialState(p_details), _
            "The View Model should be in its expected known state after restoring initial state #3; " & p_details)
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestInitialStateShouldRestore = p_outcome
    
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
''' TestSyntaxErrorShouldRecover passed. in 145.6 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestSyntaxErrorShouldRecover() As Assert

    Const p_procedureName As String = "TestSyntaxErrorShouldRecover"

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
        ' This.ViewModel.SerialPollByte = p_statusByte
        
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
        
        VBA.DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
    
    End If
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.ClearExecutionStateCommand(p_details), _
            "View Model should clear execution state and query operation completion #2; " & p_details)
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldReadCards()
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    Debug.Print VBA.vbTab & p_serialPollDetails
    
    Set TestSyntaxErrorShouldRecover = p_outcome
    
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
''' TestClosedConnectionShouldRestore passed. in 5733.7 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestClosedConnectionShouldRestore() As Assert

    Const p_procedureName As String = "TestClosedConnectionShouldRestore"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    Dim p_details As String
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Session.Socket.TryCloseConnection(p_details), _
            "View Model should close connection.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.Device.Connected, _
            "View Model should be disconnected.")
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.TryRestoreInitialState(p_details), _
            "View Model should restore its initial state; " & p_details)
            
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Device.Connected, _
            "View Model should be connected after restoring its initial state.")
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.Device.QueryOperationCompleted()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should query operation completion.")
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertShouldReadCards()

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestClosedConnectionShouldRestore = p_outcome
    
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
Public Function TestImmediateModeShouldConfigure() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestImmediateModeShouldConfigure"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' setup conditions for immediate triggering
    
    ' Immediate trigger mode is tested in single readings
    ' by turning auto increment off.
    
    ' With single reading mode (auto increment is off),
    ' the selected channel number becomes the
    ' measured channel number after the immediate reading is
    ' triggered and the observer event handler handles the
    ' measurement completion event. Thus, start with channel 1 and
    ' turn off auto increment in order to take single readings.
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = False
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ImmediateFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ImmediateSenseFunctionName
    p_mode.SingleRead = True
    p_mode.TimerInterval = This.DataView.TimerInterval
    
    ' start the immediate trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertImmediateModeShouldConfigure(p_mode, p_outcome)
    
    ' validate the immediate trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertImmediateModeShouldValidate(p_mode, p_outcome)
    
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
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestImmediateModeShouldConfigure = p_outcome
    
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
Public Function TestExternalModeShouldConfigure() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestExternalModeShouldConfigure"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = True
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ExternalFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ExternalSenseFunctionName
    p_mode.SingleRead = False
    p_mode.TimerInterval = This.DataView.TimerInterval
    
    ' start the external trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalModeShouldConfigure(p_mode, p_outcome)
    
    ' validate the external trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalModeShouldValidate(p_mode, p_outcome)
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestExternalModeShouldConfigure = p_outcome
    
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

''' summary>   Asserts that triggered readings should get polled. </summary>
''' <param name="a_assert">     [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <param name="a_duration">   [Optional, Double, 30] The time to wait for some triggered values. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssertTriggeredReadingsShouldPoll(ByVal a_assert As cc_isr_Test_Fx.Assert, _
    Optional ByVal a_duration As Double = 30) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    
    ' get some data here.
    
    ' get the pre-trigger measured channel
    Dim p_channel As Integer
    p_channel = This.DataView.MeasuredChannelNumber
    
    ' get the pre-trigger reading
    Dim p_reading As String
    p_reading = This.DataView.MeasuredReading
    
    If p_outcome.AssertSuccessful Then
    
        ' prime triggering so that we can get the trigger state.
        This.ViewModel.HandleTimerEvent
    
        ' the external trigger monitor is initiated when the timer event is
        ' handled using the above call
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger initiation should be on after first call to handle the timer event.")
    
    End If
    
    If p_outcome.AssertSuccessful Then _
        Debug.Print "Awaiting triggers..."
    
    ' loop for some time waiting for triggered measurements.
    Dim p_endTime As Double: p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow() + _
        (a_duration / cc_isr_Core_IO.CoreExtensions.SecondsPerDay)
        
    Do Until This.ViewModel.PauseRequested
        
        VBA.DoEvents
    
        ' on failure, send a stop requested.
        ' this is processed on the next timer event handler, which then
        ' sets the pause requested, which stops the timer
        If Not p_outcome.AssertSuccessful Then
        
            This.ViewModel.StopMonitoringExternalTriggersCommand
        
        ElseIf p_endTime < cc_isr_Core_IO.CoreExtensions.DaysNow() Then
            
            ' check if failed to stop on expiration.
            If This.ViewModel.StopRequested Then
            
                If p_outcome.AssertSuccessful Then
                    Set p_outcome = cc_isr_Test_Fx.Assert.Fail( _
                        "Trigger monitoring loop failed to terminate after stop was requested.")
                Else
                    Set p_outcome = cc_isr_Test_Fx.Assert.Fail( _
                        "Trigger monitoring loop failed to terminate after stop was requested after failure; " & _
                        p_outcome.AssertMessage)
                End If
                    
                ' force an exit from the loop as Pause requested fails to materialize,
                ' which is a bug that needs to be fixed.
                
                Exit Do
                
            Else
            
                ' on expiration, set stop request.
                This.ViewModel.StopMonitoringExternalTriggersCommand
            
                ' set a new end time to timeout in one second if the loop does not exit
                p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow() + _
                    (1 / cc_isr_Core_IO.CoreExtensions.SecondsPerDay)
            End If
        
        Else
        
            ' invoke the event handler, which either:
            ' handles a trigger or
            ' issues a pause request if stop was requested and
            '   moves the measurement mode to none when done.
    
            On Error Resume Next
            This.ViewModel.HandleTimerEvent
            If Err.Number <> 0 Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail( _
                    "Error #" & Err.Number & " handling timer event; " & Err.Description & ".")
            End If
            On Error GoTo 0
            
        End If
        
        ' record reading if the measured channel number changed.
        
        If p_channel <> This.DataView.MeasuredChannelNumber Then
        
            VBA.DoEvents
            p_channel = This.DataView.MeasuredChannelNumber
            
            VBA.DoEvents
            p_reading = This.DataView.MeasuredReading
            
            VBA.DoEvents
            Debug.Print p_channel; ": "; p_reading

            ' delay processing the next event by the presumed timer interval.
            ' cc_isr_Core_IO.Factory.NewStopwatch().Wait 500
            
            ' verify that measured channel numbers propagated correctly.
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.Observer.MeasuredChannelNumber, _
                    This.ViewModel.MeasuredChannelNumber, _
                    "View Model measured channel number should equal the Observer measured channel.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.DataView.MeasuredChannelNumber, _
                    This.ViewModel.MeasuredChannelNumber, _
                    "View Model measured channel number should equal the Data View measured channel.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(ExpectedTargetChannelNumber(), _
                    This.ViewModel.TargetChannelNumber, _
                    "The target channel number should equal the expected target channel number after a triggered reading.")
            
            If p_outcome.AssertSuccessful Then _
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(This.ViewModel.TargetChannelNumber, _
                    This.Observer.TargetChannelNumber, _
                    "The observer Target Channel Number should equal the view model target channel number.")

        End If
    
    Loop
    
    Set AssertTriggeredReadingsShouldPoll = p_outcome

End Function

''' <summary>   Asserts that view model should poll triggered readings. </summary>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">     [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <param name="a_enabled">    [Optional, Boolean, True] True to enable reading triggered values. </param>
''' <param name="a_duration">   [Optional, Double, 30] The time to wait for some triggered values. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssetTriggersShouldPoll(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert, _
    Optional ByVal a_enabled As Boolean = True, _
    Optional ByVal a_duration As Double = 30) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' setup conditions for monitoring
    
    ' Multiple readings (auto increment on) is used for testing
    ' trigger monitoring. The test checks that channel numbers change
    ' with each trigger.
    
    ' With multiple readings (auto increment is on),
    ' channel numbers start with the Target Channel Number.
    ' Start with channel 1
    
    ' start the external trigger mode
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(0, a_mode.TimerInterval, _
            "Timer interval should be zero when polling triggers.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalModeShouldConfigure(a_mode, p_outcome)
    
    ' validate the external trigger reading mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalModeShouldValidate(a_mode, p_outcome)
    
    ' start the monitoring mode turning timer monitoring off.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStart(0, p_outcome)
    
    ' validate the monitoring mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldValidate(a_mode, p_outcome)
    
    If p_outcome.AssertSuccessful And a_enabled Then _
        Set p_outcome = AssertTriggeredReadingsShouldPoll(p_outcome, a_duration)
    
    ' stop monitoring here
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStop(p_outcome)
    
    ' Finally, verify that no error message was recorded.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

    Set AssetTriggersShouldPoll = p_outcome

End Function

''' <summary>   Unit test. Asserts that view model should start and stop polling triggered readings. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestTriggerPollingShouldStartStop() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTriggerPollingShouldStartStop"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_enabled As Boolean: p_enabled = False ' for now
    Dim p_duration As Double: p_duration = 5  ' in seconds
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = True
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ExternalFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ExternalSenseFunctionName
    p_mode.SingleRead = False
    p_mode.TimerInterval = 0 ' this uses polling rather than timer
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssetTriggersShouldPoll(p_mode, p_outcome, p_enabled, p_duration)
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTriggerPollingShouldStartStop = p_outcome
    
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


''' <summary>   Unit test. Asserts that view model should poll triggering. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' Awaiting triggers...
'''  1 : 100.115234
'''  2 : 100.114975
'''  3 : 100.116783
'''  4 : 100.117149
'''  5 : 100.115334
'''  6 : 100.115814
'''  7 : 100.116417
''' Test 10 TestTriggerPollingShouldRead passed. Elapsed time: 11113.5 ms.
''' This was recorded after adding *CLS after receiving the trigger.
''' Awaiting triggers...
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00121643E+02'.
'''  1 : 100.121643
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00123245E+02'.
'''  2 : 100.123245
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00122681E+02'.
'''  3 : 100.122681
''' Test 10 TestTriggerPollingShouldRead passed. Elapsed time: 16133.3 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestTriggerPollingShouldRead() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTriggerPollingShouldRead"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_enabled As Boolean: p_enabled = True
    Dim p_duration As Double: p_duration = 5  ' in seconds
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = True
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ExternalFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ExternalSenseFunctionName
    p_mode.SingleRead = False
    p_mode.TimerInterval = 0 ' this uses polling rather than timer
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssetTriggersShouldPoll(p_mode, p_outcome, p_enabled, p_duration)
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTriggerPollingShouldRead = p_outcome
    
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

''' <summary>   Asserts that view model should monitor triggered readings. </summary>
''' <param name="a_timerInterval">   [Integer] the timer inteva; 0 if polling the timer event. </param>
''' <param name="a_mode">     [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_assert">          [<see cref="cc_isr_Test_Fx.Assert"/>] The assert status of the test method. </param>
''' <param name="a_enabled">         [Optional, Boolean, True] True to enable reading triggered values. </param>
''' <param name="a_duration">        [Optional, Double, 30] The time to wait for some triggered values. </param>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function AssetTriggersShouldMonitor(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByVal a_assert As cc_isr_Test_Fx.Assert, _
    Optional ByVal a_enabled As Boolean = True, _
    Optional ByVal a_duration As Double = 30) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = a_assert
    Dim p_details As String: p_details = VBA.vbNullString

    ' setup conditions for monitoring
    
    ' Multiple readings (auto increment on) is used for testing
    ' trigger monitoring. The test checks that channel numbers change
    ' with each trigger.
    
    ' With multiple readings (auto increment is on),
    ' channel numbers start with the Target Channel Number.
    ' Start with channel 1 and
    
    ' start the external trigger mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(a_mode.TimerInterval > 0, _
            "Timer interval should be greater than zero monitoring triggers.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertExternalModeShouldConfigure(a_mode, p_outcome)
        
    
    ' start the monitoring mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStart(a_mode.TimerInterval, p_outcome)
    
    ' validate the monitoring mode
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldValidate(a_mode, p_outcome)
    
    ' get some data here.

    If p_outcome.AssertSuccessful And a_enabled Then _
        Set p_outcome = AssertMeasurementsShouldGetTriggered(p_outcome, a_duration)

    ' stop monitoring here
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertMonitoringModeShouldStop(p_outcome)
    
    ' Finally, verify that no error message was recorded.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

    Set AssetTriggersShouldMonitor = p_outcome
    
End Function

''' <summary>   Unit test. Asserts that view model should start and stop trigger monitoring. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' Test 11 TestTriggerMonitoringShouldStartStop passed. Elapsed time: 8220.0 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestTriggerMonitoringShouldStartStop() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTriggerMonitoringShouldStartStop"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_enabled As Boolean: p_enabled = False ' for now
    Dim p_duration As Double: p_duration = 5  ' in seconds
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = True
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ExternalFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ExternalSenseFunctionName
    p_mode.SingleRead = False
    p_mode.TimerInterval = This.DataView.TimerInterval
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssetTriggersShouldMonitor(p_mode, p_outcome, p_enabled, p_duration)
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTriggerMonitoringShouldStartStop = p_outcome
    
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
''' Waiting for trigger....
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00122772E+02'.
'''  1 : 100.122772
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00122467E+02'.
'''  2 : 100.122467
''' Test 12 TestTriggerMonitoringShouldRead passed. Elapsed time: 13419.9 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestTriggerMonitoringShouldRead() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestTriggerMonitoringShouldRead"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_enabled As Boolean: p_enabled = True
    Dim p_duration As Double: p_duration = 5  ' in seconds
    
    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.BeepEnabled = False
    p_mode.AutoIncrementChannel = True
    p_mode.ChannelNumber = This.UserView.SelectedChannelNumber
    p_mode.FrontInputs = This.DataView.ExternalFrontInputsRequired
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = This.UserView.ReadingOffset
    p_mode.SenseFunction = This.DataView.ExternalSenseFunctionName
    p_mode.SingleRead = False
    p_mode.TimerInterval = This.DataView.TimerInterval
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssetTriggersShouldMonitor(p_mode, p_outcome, p_enabled, p_duration)
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTriggerMonitoringShouldRead = p_outcome
    
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

''' <summary>   Unit test. Asserts that view model should connect after power on reset. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' 8:40:13 Power on reset starting. This could take 3 seconds. Please wait...
''' 8:40:19 done power on reset.
''' Test 16 TestOpenConnectionWithPowerOnResetShouldConnect passed. Elapsed time: 6517.1 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestOpenConnectionWithPowerOnResetShouldConnect() As Assert

    Const p_procedureName As String = "TestOpenConnectionWithPowerOnResetShouldConnect"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_details As String
    Dim p_success As Boolean
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_actualReply As String
    Dim p_expectedReply As String

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Session.Socket.TryCloseConnection(p_details), _
            "View Model should close connection.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse(This.ViewModel.Device.Connected, _
            "View Model should be disconnected.")

    If p_outcome.AssertSuccessful Then
    
        Dim p_delay As Double: p_delay = 3
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " Power on reset starting. This could take "; _
            VBA.CStr(p_delay); " seconds. Please wait..."
        
        p_success = This.ViewModel.TryOpenConnectionPowerOnReset(This.DataView.SocketAddress, _
            This.DataView.SessionTimeout, p_delay, p_details)

        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, _
            "View Model should open connection with power on reset; " & p_details)
            
        Debug.Print VBA.Format$(Now, "h:mm:ss"); " done power on reset."
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Session.Connected, _
            "View Model session should be connected after restoring its initial state.")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(This.ViewModel.Device.Connected, _
            "View Model device should be connected after restoring its initial state.")
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.Device.QueryOperationCompleted()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should query operation completion.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestOpenConnectionWithPowerOnResetShouldConnect = p_outcome
    
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


