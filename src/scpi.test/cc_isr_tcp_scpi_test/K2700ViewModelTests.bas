Attribute VB_Name = "K2700ViewModelTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 View Model Tests. </summary>
''' <remarks>   Dependencies: cc_isr_Core_Tcp_Scpi.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    Observer As K2700Observer
    Host As String
    Port As Long
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    ContinuousSenseFunctionName As String
    ImmediateSenseFunctionName As String
    ExternalSenseFunctionName As String
    ErrTracer As cc_isr_Test_Fx.IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestShouldInitialize
        Case 2
            Set p_outcome = TestShouldBeConnected
        Case 3
            Set p_outcome = TestShouldReadCards
        Case 4
            Set p_outcome = TestShouldRestoreKnownState
        Case 5
            Set p_outcome = TestShouldRecoverFromSyntaxError
        Case 6
            Set p_outcome = TestShouldRecoverFromReadAfterWriteTrue
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
    RunTest 4
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 8
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

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("Primed to run all tests.")

    This.Name = "K2700ViewModelTests"
    
    This.TestNumber = 0
    
    ' set a temporary error tracer
    Set This.ErrTracer = New DeviceErrorsTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
    ' initialize known data.
    This.TopCard = "7700"
    This.BottomCard = VBA.vbNullString
    This.ContinuousSenseFunctionName = "FRES"
    This.ImmediateSenseFunctionName = "RES"
    This.ExternalSenseFunctionName = "FRES"
    ' card scan list uses immediate mode sense function
    This.TopCardFunctionScanList = ":FUNC 'RES',(@101,120)"
    This.BottomCardFunctionScanList = VBA.vbNullString
    
    Set This.ViewModel = cc_isr_Tcp_Scpi.K2700ViewModel
    
    Set This.Observer = K2700Observer.Initialize(This.ViewModel)
    
    This.ViewModel.Host = "192.168.0.252"
    This.ViewModel.Port = 1234
    This.ViewModel.SocketReceiveTimeout = 100
    
    ' initialize the view model.
    This.ViewModel.Initialize

    ' set the final error tracer capable of reporting device errors.
    Dim p_errTracer As New DeviceErrorsTracer
    Set This.ErrTracer = p_errTracer.Initialize(This.ViewModel.Device)
    
    ' connect
    This.ViewModel.OpenConnectionCommand
    
    If This.ViewModel.Connected Then
        Set p_outcome = Assert.Pass("Primed to run all tests; K2700 View Model is connected.")
    Else
        Set p_outcome = Assert.Inconclusive( _
            "Failed priming all tests; K2700 View Model should be connected.")
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

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.ViewModel.Connected, _
            Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & "; K2700 View Model is Connected."), _
            Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                "; K2700 View Model should be connected."))
    Else
        Set p_outcome = Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
   
    ' clear execution state before each test.
    
    If p_outcome.AssertSuccessful Then _
        This.ViewModel.Device.ClearExecutionState
   
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

    Dim p_outcome As cc_isr_Test_Fx.Assert
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
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("All tests cleaned up.")
    
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
' Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts that view model should initialize. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldInitialize() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldInitialize"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.ToggleConnectionExecutable, _
            "Toggle connection should be executable after initializing the View Model.")

    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldInitialize")
    
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
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldBeConnected() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldBeConnected"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.Connected, _
            "View model should be connected.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.CloseConnectionExecutable, _
            "View model close connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Observer.CloseConnectionExecutable, _
            "View model Observer close connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Observer.ToggleConnectionExecutable, _
            "View model Observer toggle connection executable should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsFalse(This.ViewModel.OpenConnectionExecutable, _
            "View model open connection executable should be disabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.ImmediateTriggerOptionExecutable, _
            "Immediate trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.ExternalTriggerOptionExecutable, _
            "External trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.Observer.ExternalTriggerOptionExecutable, _
            "Observer External trigger option button should be enabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsFalse(This.ViewModel.StartMonitoringExecutable, _
            "Start monitoring command should be disabled.")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "Stop monitoning command should be disabled.")
        
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldBeConnected")
    
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
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.TopCard, This.ViewModel.TopCard, _
            "View Model should be read the top card")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.BottomCard, This.ViewModel.BottomCard, _
            "View Model should be read the bottom card")

    ' the view module initializes in continuous mode.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.ContinuousSenseFunctionName, _
            This.ViewModel.SenseFunctionName, _
            "View Model should set the sense function name")

    ' the cards are set for immeidate mode.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.TopCardFunctionScanList, _
            This.ViewModel.TopCardFunctionScanList, _
            "View Model should be read the top card function scan list")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.BottomCardFunctionScanList, _
            This.ViewModel.BottomCardFunctionScanList, _
            "View Model should be read the top card function scan list")
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

    Set AssertShouldReadCards = p_outcome
End Function

''' <summary>   Unit test. Asserts that view model should read cards. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldReadCards() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldReadCards"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldReadCards")
    
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

''' <summary>   Unit test. Asserts that view model should restore known state. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRestoreKnownState() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldRestoreKnownState"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.

    Dim p_details As String: p_details = VBA.vbNullString

    ' check if we need to restore the GPIB-Lan know state.
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(This.ViewModel.ShouldRestoreKnownState(p_details), _
            "The GPIB-Lan should be in its expected known state upon connecting; " & p_details)
    End If

    Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = "VOLT:DC"
    If p_outcome.AssertSuccessful Then
        
        ' change function mode to voltage
        This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionSetter p_expectedSenseFunctionName
        Set p_outcome = Assert.AreEqual(p_expectedSenseFunctionName, _
            This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunction, _
            "Sense function should be set to the expected value.")
            
    End If
    
    Dim p_actualSenseFunctionName As String
    If p_outcome.AssertSuccessful Then
        
        ' validate the actual function
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = Assert.AreEqual(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            "Actual sense function should be set to the expected value.")
            
    End If
    
    ' now that the function was changed, a resore should be required
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, p_details), _
            "Restore should be required after setting the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then
        
        This.ViewModel.RestoreKnownState
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        
        ' once restore, restore should no longer be required
        Set p_outcome = Assert.IsFalse(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, p_details), _
            "Restore should not be required after restoring the function to: '" & p_actualSenseFunctionName & "'; " & _
            p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(This.ViewModel.ShouldRestoreKnownState(p_details), _
            "The View Model should be in its expected known state after restoring state #1; " & p_details)
    End If
    
    If p_outcome.AssertSuccessful Then
        This.ViewModel.ViSession.GpibLan.AssertTalkAfterWriteAllowed = True
        Set p_outcome = Assert.IsTrue(This.ViewModel.ShouldRestoreKnownState(p_details), _
            "The View Model should not be in its expected known state after setting allowing to assert talk after write ('" & _
            VBA.CStr(This.ViewModel.ViSession.GpibLan.AssertTalkAfterWriteAllowed) & "').")
    End If
    
    ' if restore is required we should restore
    If p_outcome.AssertSuccessful Then
        
        This.ViewModel.RestoreKnownState
        
        ' once restored, restore should no longer be required
        Set p_outcome = Assert.IsFalse(This.ViewModel.ShouldRestoreSenseFunction(p_actualSenseFunctionName, p_details), _
            "Restore should not be required after restoring the assert talk after write to: '" & _
            VBA.CStr(This.ViewModel.ViSession.GpibLan.AssertTalkAfterWriteAllowed) & "'; " & _
            p_details)
        
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(This.ViewModel.ShouldRestoreKnownState(p_details), _
            "The View Model should be in its expected known state after restoring know state #2; " & p_details)
    End If
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldRestoreKnownState")
    
    Set TestShouldRestoreKnownState = p_outcome
    
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
Public Function TestShouldRecoverFromSyntaxError() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromSyntaxError"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.ClearExecutionStateCommand()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should clear execution state and query operation completion #1.")
    End If

    If p_outcome.AssertSuccessful Then
        
        ' issue a bad command
        On Error Resume Next
        This.ViewModel.ViSession.WriteLine ("**OPC")
        On Error GoTo 0
        
        ' clear the error state
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
        
        DoEvents
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 100
        
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.ClearExecutionStateCommand()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should clear ewxecution state and query operation completion #2.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldRecoverFromSyntaxError")
    
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

''' <summary>   Unit test. Asserts that view model should recover from read after write true condition. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldRecoverFromReadAfterWriteTrue() As Assert

    Const p_procedureName As String = "TestShouldRecoverFromReadAfterWriteTrue"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    Dim p_errorNumber As String
    Dim p_errorMessage As String
    Dim p_success As Boolean
    Dim p_actualReply As String
    Dim p_expectedReply As String
    
    If p_outcome.AssertSuccessful And This.ViewModel.ViSession.UsingGpibLan Then
        ' turn on read after write condition.
        This.ViewModel.ViSession.GpibLan.ReadAfterWriteEnabledSetter True
        Set p_outcome = Assert.IsTrue(This.ViewModel.ViSession.GpibLan.ReadAfterWriteEnabledGetter, _
            "Read after write should be true.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.ViewModel.Device.CloseConnection
        Set p_outcome = Assert.IsFalse(This.ViewModel.Device.Connected, _
            "View Model should be disconnected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        This.ViewModel.RestoreKnownState
        Set p_outcome = Assert.IsTrue(This.ViewModel.Device.Connected, _
            "View Model should be connected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_expectedReply = "1"
        p_actualReply = This.ViewModel.Device.QueryOperationCompleted()
        Set p_outcome = Assert.AreEqual(p_expectedReply, p_actualReply, _
            "View Model should query operation completion.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = AssertShouldReadCards()
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldRecoverFromReadAfterWriteTrue")
    
    Set TestShouldRecoverFromReadAfterWriteTrue = p_outcome
    
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
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConfigureImmediateMode() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConfigureImmediateMode"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' configure immediate mode with front switch.
        This.ViewModel.FrontInputsRequired = True
        This.ViewModel.ConfigureImmediateTriggerReadingsCommand
        
        Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = This.ImmediateSenseFunctionName
        Dim p_actualSenseFunctionName As String:
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = Assert.AreEqualString(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            VBA.VbCompareMethod.vbTextCompare, _
            "Immediate mode sense function name should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementMode
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementMode.Immediate
        Set p_outcome = Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "Immediate measurement mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSource.Immediate, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "Immediate trigger source should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' turn off auto increment so that we can take a single reading.
        Dim p_autoIncrementChannelNumber  As Boolean: p_autoIncrementChannelNumber = False
        This.ViewModel.AutoIncrementChannelNoEnabled = p_autoIncrementChannelNumber
        Set p_outcome = Assert.AreEqual(p_autoIncrementChannelNumber, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        ' take a reading
        This.ViewModel.MeasureImmediatelyCommand
        
        ' wait for the reading event to take shape.
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        
        ' get the reading from the observer.
        Dim p_reading As String
        p_reading = This.Observer.MeasuredReading
        
        Dim p_channelNumber As Integer
        p_channelNumber = This.Observer.MeasuredChannelNumber

        Dim p_readingValue As Double
        p_readingValue = This.Observer.MeasuredValue
        
        Set p_outcome = Assert.AreEqual(This.ViewModel.SelectedChannelNumber, p_channelNumber, _
            "Reading should come from the expected channel number.")
            
        Set p_outcome = Assert.IsFalse(VBA.vbNullString = p_reading, _
            "Reading should not be empty.")
            
        Set p_outcome = Assert.IsTrue(p_readingValue > 0, _
            "Reading value should be positive.")
            
    End If
    

    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldConfigureImmediateMode")
    
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
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConfigureExternalMode() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConfigureExternalMode"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' configure external mode with front switch.
        This.ViewModel.FrontInputsRequired = True
        This.ViewModel.ConfigureExternalTriggerReadingCommand
        
        Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = This.ExternalSenseFunctionName
        Dim p_actualSenseFunctionName As String:
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = Assert.AreEqualString(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            VBA.VbCompareMethod.vbTextCompare, _
            "External mode sense function name should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Dim p_expectedMeasurementMode As cc_isr_Tcp_Scpi.MeasurementMode
        p_expectedMeasurementMode = cc_isr_Tcp_Scpi.MeasurementMode.External
        Set p_outcome = Assert.AreEqual(p_expectedMeasurementMode, This.ViewModel.CurrentMeasurementMode, _
            "External measurement mode should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(cc_isr_Tcp_Scpi.TriggerSource.External, _
            This.ViewModel.K2700.TriggerSystem.SourceGetter(), _
            "External trigger source should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.IsFalse(This.ViewModel.K2700.TriggerSystem.ContinuousEnabledGetter, _
            "Continuous trigger should be disabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.SampleCountGetter, _
            "Sample count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(1, This.ViewModel.K2700.TriggerSystem.TriggerCountGetter, _
            "Trigger count should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.IsTrue(This.ViewModel.K2700.SenseSystem.SenseSystem.AutoRangeEnabledGetter(), _
            "Auto range should be enabled.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual(1#, _
            This.ViewModel.K2700.SenseSystem.SenseSystem.PowerLineCyclesGetter(), _
            "The integration rate in power line cycles should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.AreEqual("READ,,,,,", This.ViewModel.K2700.FormatSystem.ElementsGetter, _
            "Format elements should be as expected.")
    End If
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldConfigureExternalMode")
    
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
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldMonitorTriggering() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldMonitorTriggering"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        
        ' configure external mode with front switch.
        This.ViewModel.FrontInputsRequired = True
        This.ViewModel.ConfigureExternalTriggerReadingCommand
        
        Dim p_expectedSenseFunctionName As String: p_expectedSenseFunctionName = This.ExternalSenseFunctionName
        Dim p_actualSenseFunctionName As String:
        p_actualSenseFunctionName = This.ViewModel.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        Set p_outcome = Assert.AreEqualString(p_expectedSenseFunctionName, p_actualSenseFunctionName, _
            VBA.VbCompareMethod.vbTextCompare, _
            "External mode sense function name should be as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_minimumTimerInterval As Integer: p_minimumTimerInterval = 100
        Set p_outcome = Assert.IsTrue(This.ViewModel.TimerInterval >= p_minimumTimerInterval, _
            "Timer interval should exceed" & VBA.CStr(p_minimumTimerInterval - 1) & ".")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger initiated should be off before starting the monitoring timer.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_autoIncrementChannelNumber  As Boolean: p_autoIncrementChannelNumber = False
        This.ViewModel.AutoIncrementChannelNoEnabled = p_autoIncrementChannelNumber
        Set p_outcome = Assert.AreEqual(p_autoIncrementChannelNumber, _
            This.ViewModel.AutoIncrementChannelNoEnabled, _
            "Auto increment channel number should be as expected.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_selectedChannelNumber  As Integer: p_selectedChannelNumber = 3
        This.ViewModel.SelectedChannelNumber = p_selectedChannelNumber
        Set p_outcome = Assert.AreEqual(p_selectedChannelNumber, This.ViewModel.SelectedChannelNumber, _
            "Selected channel number should be as expected.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.StartMonitoringExternalTriggersCommand
        Set p_outcome = Assert.IsFalse(This.ViewModel.PauseRequested, _
            "Pause Requested should be off after starting the monitoring timer.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        Set p_outcome = Assert.IsTrue(This.ViewModel.K2700.ExtTrigInitiated, _
            "External trigger should get initiated after starting the monitoring timer.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsTrue(This.ViewModel.StopMonitoringExecutable, _
            "The stop monitoring executable to should enabeld.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.AreEqual(p_selectedChannelNumber, This.ViewModel.CurrentChannelNumber, _
            "Current channel number should be set to the selected channel number.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        This.ViewModel.StopMonitoringExternalTriggersCommand
        Set p_outcome = Assert.IsTrue(This.ViewModel.StopRequested, _
            "Stop Requested should be on off after stopping monitoring.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        cc_isr_Core_IO.Factory.NewStopwatch().Wait 10
        Set p_outcome = Assert.IsTrue(This.ViewModel.PauseRequested, _
            "Pause should be requested after stopping monitoring.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(This.ViewModel.StopMonitoringExecutable, _
            "The stop monitoring executable to should disabled after stopping monitoring.")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(This.ViewModel.ExternalTrigMonitoringEnabled, _
            "External monitoring enabled should be off after stopping monitoring.")
        
    End If
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldMonitorTriggering")
    
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
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestViewModelTestTemplate() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestViewModelTestTemplate"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    
    
    
    
    ' Finally, verify that no error message was recorded.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Last error message should be empty but found: '" & This.ViewModel.LastErrorMessage & "'.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestViewModelTestTemplate")
    
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


