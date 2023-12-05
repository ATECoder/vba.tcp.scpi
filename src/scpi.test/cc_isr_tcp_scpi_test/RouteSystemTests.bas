Attribute VB_Name = "RouteSystemTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Route System Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_debugPrintPrefix As String = "''' "

''' <summary>   This class properties. </summary>
Private Type this_
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
            Set p_outcome = Test7700CardsShouldBePopulated
        Case 2
            Set p_outcome = Test7700CardsShouldSelected
        Case 3
            Set p_outcome = Test7700CardsShouldBuildScanLists
        Case 4
            Set p_outcome = Test7700CardsShouldBuild4WireScanLists
        Case Else
    End Select
    AfterEach
    Set RunTest = p_outcome
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
''' 12/4/2023 3:17:16 PM
''' Test 01 Test7700CardsShouldBePopulated passed. Elapsed time: 0.1 ms.
''' Test 02 Test7700CardsShouldSelected passed. Elapsed time: 0.1 ms.
''' Test 03 Test7700CardsShouldBuildScanLists passed. Elapsed time: 0.3 ms.
''' Test 04 Test7700CardsShouldBuild4WireScanLists passed. Elapsed time: 0.6 ms.
''' Ran 4 out of 4 tests.
''' Passed: 4; Failed: 0; Inconclusive: 0.
''' </code>
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 4
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
    Debug.Print m_debugPrintPrefix; _
        "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print m_debugPrintPrefix; _
        "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
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
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    ' initialize the current and previous test numbers.
    This.TestNumber = 0
    This.PreviousTestNumber = 0

    Debug.Print m_debugPrintPrefix; Date; Time
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")

    This.Name = "RouteSystemTests"
    
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' Prime all tests
    
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
         Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
   
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
  
    ' set the previous test number to the current test number.
    This.PreviousTestNumber = This.TestNumber
    
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
    End If
    
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
'  Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts populating the multimplexer card 7700 cards. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function Test7700CardsShouldBePopulated() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "Test7700CardsShouldBePopulated"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    
    If p_outcome.AssertSuccessful Then
        Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewTcpSession())
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_routeSystem, TypeName(p_routeSystem) & " should be instantiated.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_routeSystem.InstrumentFamilyCards, "Instrument family card collection should be instantiated.")
    End If
    
    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.Populate7700Cards
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    p_cardName = "7700"
    Dim p_card As MultiplexerCard
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.InstrumentFamilyCards(p_cardName)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, "The expected cad should be selected from the Instrument family card collection.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print m_debugPrintPrefix; "Test " & Format(This.TestNumber, "00") & " " & _
        p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set Test7700CardsShouldBePopulated = p_outcome
    
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

''' <summary>   Unit test. Asserts populating the multimplexer card 7700 cards. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function Test7700CardsShouldSelected() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "Test7700CardsShouldSelected"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    If p_outcome.AssertSuccessful Then
        Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewTcpSession())
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_routeSystem, _
            TypeName(p_routeSystem) & " should be instantiated.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_routeSystem.InstrumentFamilyCards, _
            "Instrument family card collection should be instantiated.")
    End If
    
    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.Populate7700Cards
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, _
            "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    Dim p_card As MultiplexerCard
    
    If p_outcome.AssertSuccessful Then
        p_cardName = "7700"
        Set p_card = p_routeSystem.InstrumentFamilyCards(p_cardName)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the Instrument family card collection.")
    End If
    
    Dim p_options As String: p_options = "7700,7702"
    p_cardName = "7700"
    p_expectedCount = 2
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.PopulateCards(p_options)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, p_actualCount, _
            "Installed card collection should have the expected number of cards.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_cardName = "7700"
        Dim p_cards As Collection
        Set p_cards = p_routeSystem.InstalledCards
        Set p_card = p_cards(1)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the installed card collection.")
    End If
    
    Dim p_expectedCapacity As Integer
    Dim p_expectedFirstChannel As Integer
    Dim p_expectedLastChannel As Integer
    Dim p_expectedSlotNumber As Integer
    
    p_cardName = "7700"
    p_expectedFirstChannel = 1
    p_expectedLastChannel = 20
    p_expectedCapacity = 20
    p_expectedSlotNumber = 1
    
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.InstalledCards(CStr(p_expectedSlotNumber))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the installed card collection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, p_card.Capacity, _
            "Card '" & p_cardName & "' should have the expected capacity.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFirstChannel, p_card.DeviceFirstChannel, _
            "Card '" & p_cardName & "' should have the expected first channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedLastChannel, p_card.DeviceLastChannel, _
            "Card '" & p_cardName & "' should have the expected last channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedSlotNumber, p_card.SlotNumber, _
            "Card '" & p_cardName & "' should have the expected slot number.")
    End If
    
    p_cardName = "7702"
    p_expectedFirstChannel = 21
    p_expectedLastChannel = 60
    p_expectedCapacity = 40
    p_expectedSlotNumber = 2
    
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.InstalledCards(CStr(p_expectedSlotNumber))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the installed card collection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, p_card.Capacity, _
            "Card '" & p_cardName & "' should have the expected capacity.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFirstChannel, p_card.DeviceFirstChannel, _
            "Card '" & p_cardName & "' should have the expected first channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedLastChannel, p_card.DeviceLastChannel, _
            "Card '" & p_cardName & "' should have the expected last channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedSlotNumber, p_card.SlotNumber, _
            "Card '" & p_cardName & "' should have the expected slot number.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print m_debugPrintPrefix; "Test " & Format(This.TestNumber, "00") & " " & _
        p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set Test7700CardsShouldSelected = p_outcome
    
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

''' <summary>   Asserts building scan lists. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function Assert7700CardsShouldBuildScanLists(ByVal a_senseFunctionName As String) As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "Assert7700CardsShouldBuildScanLists"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    If p_outcome.AssertSuccessful Then
        Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewTcpSession())
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_routeSystem, _
            TypeName(p_routeSystem) & " should be instantiated.")
    End If

    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.Populate7700Cards
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, _
            "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    Dim p_card As MultiplexerCard
    Dim p_deviceChannelNumber As Integer
    Dim p_options As String: p_options = "7700,7702"
    p_expectedCount = 2
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.PopulateCards(p_options)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, p_actualCount, _
            "Installed card collection should have the expected number of cards.")
    End If
    
    p_cardName = "7700"
    p_deviceChannelNumber = 1
    If p_outcome.AssertSuccessful Then
        
        ' build the scan lists here so as to set the channel numbers properly.
        p_routeSystem.BuildFunctionScanLists a_senseFunctionName
        
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    Dim p_expectedFunctionScanList As String
    p_expectedFunctionScanList = ":FUNC '" & a_senseFunctionName & "',(@101,1" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    Dim p_actualFunctionScanList As String
    
    If p_outcome.AssertSuccessful Then
        p_actualFunctionScanList = p_card.FunctionScanList
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFunctionScanList, p_actualFunctionScanList, _
            "The expected scan list should be built for card '" & p_cardName & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_deviceChannelNumber = p_deviceChannelNumber + p_card.FunctionalCapacity - 1
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    Dim p_expectedScanList As String
    p_expectedScanList = "(@1" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    Dim p_actualScanList As String
    If p_outcome.AssertSuccessful Then
        p_actualScanList = p_card.BuildChannelScanList(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedScanList, p_actualScanList, _
            "The expected scan list should be returned for card '" & p_cardName & "' and channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    Dim p_expectedRouteCommand As String
    p_expectedRouteCommand = ":ROUT:MULT:CLOS (@124,125)"
    Dim p_actualRouteCommand As String
    If p_outcome.AssertSuccessful Then
        p_actualRouteCommand = p_card.RouteMultipleCloseCommand()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedRouteCommand, p_actualRouteCommand, _
            "The expected route command should be returned for card '" & p_cardName & "'.")
    End If
    
    p_cardName = "7702"
    p_deviceChannelNumber = p_deviceChannelNumber + 1
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_actualFunctionScanList = p_card.FunctionScanList
        p_expectedFunctionScanList = ":FUNC '" & a_senseFunctionName & "',(@201,2" & VBA.CStr(p_card.FunctionalCapacity) & ")"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFunctionScanList, p_actualFunctionScanList, _
            "The expected scan list should be built for card '" & p_cardName & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_deviceChannelNumber = p_deviceChannelNumber + p_card.FunctionalCapacity - 1
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    p_expectedScanList = "(@2" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    If p_outcome.AssertSuccessful Then
        p_actualScanList = p_card.BuildChannelScanList(p_deviceChannelNumber)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedScanList, p_actualScanList, _
            "The expected scan list should be returned for card '" & p_cardName & "' and channel " & CStr(p_deviceChannelNumber) & ".")
    End If
    
    p_expectedRouteCommand = ":ROUT:MULT:CLOS (@244,245)"
    If p_outcome.AssertSuccessful Then
        p_actualRouteCommand = p_card.RouteMultipleCloseCommand()
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedRouteCommand, p_actualRouteCommand, _
            "The expected route command should be returned for card '" & p_cardName & "'.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Set Assert7700CardsShouldBuildScanLists = p_outcome
    
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


''' <summary>   Unit test. Asserts building scan lists. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function Test7700CardsShouldBuildScanLists() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "Test7700CardsShouldBuildScanLists"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert7700CardsShouldBuildScanLists("RES")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print m_debugPrintPrefix; "Test " & Format(This.TestNumber, "00") & " " & _
        p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set Test7700CardsShouldBuildScanLists = p_outcome
    
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

''' <summary>   Unit test. Asserts building 4-wire scan lists. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function Test7700CardsShouldBuild4WireScanLists() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "Test7700CardsShouldBuild4WireScanLists"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Entered the " & p_procedureName & " test.")
    End If
    
    ' proceed with test assertions.
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert7700CardsShouldBuildScanLists("FRES")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print m_debugPrintPrefix; "Test " & Format(This.TestNumber, "00") & " " & _
        p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set Test7700CardsShouldBuild4WireScanLists = p_outcome
    
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

