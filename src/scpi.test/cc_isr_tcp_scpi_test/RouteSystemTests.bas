Attribute VB_Name = "RouteSystemTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Route System Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ErrTracer As IErrTracer
End Type

Private This As this_

Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            Test7700CardsShouldBePopulated
        Case 2
            Test7700CardsShouldSelected
        Case 3
            Test7700CardsShouldBuildScanLists
        Case 4
            Test7700CardsShouldBuild4WireScanLists
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
    
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error object.
    
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    Set This.BeforeEachAssert = Assert.IsTrue(True, "initialize the pre-test assert.")
    
    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
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
    
End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    Set This.BeforeAllAssert = Nothing
End Sub

''' <summary>   Unit test. Asserts populating the multimplexer card 7700 cards. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Test7700CardsShouldBePopulated() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewViSession())
    
    Set p_outcome = Assert.IsNotNothing(p_routeSystem, TypeName(p_routeSystem) & " should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsNotNothing(p_routeSystem.InstrumentFamilyCards, "Instrument family cardcollection should be instantiated.")
    End If
    
    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer: p_actualCount = p_routeSystem.Populate7700Cards
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    p_cardName = "7700"
    Dim p_card As MultiplexerCard
    Set p_card = p_routeSystem.InstrumentFamilyCards(p_cardName)
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, "The expected cad should be selected from the Instrument family card collection.")
    End If
    
    Debug.Print p_outcome.BuildReport("Test7700CardsShouldBePopulated")
    
    Set Test7700CardsShouldBePopulated = p_outcome
    
End Function

''' <summary>   Unit test. Asserts populating the multimplexer card 7700 cards. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Test7700CardsShouldSelected() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewViSession())
    
    Set p_outcome = Assert.IsNotNothing(p_routeSystem, _
        TypeName(p_routeSystem) & " should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsNotNothing(p_routeSystem.InstrumentFamilyCards, _
            "Instrument family cardcollection should be instantiated.")
    End If
    
    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.Populate7700Cards
        Set p_outcome = Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, _
            "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    Dim p_card As MultiplexerCard
    
    If p_outcome.AssertSuccessful Then
        p_cardName = "7700"
        Set p_card = p_routeSystem.InstrumentFamilyCards(p_cardName)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the Instrument family card collection.")
    End If
    
    Dim p_options As String: p_options = "7700,7702"
    p_cardName = "7700"
    p_expectedCount = 2
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.PopulateCards(p_options)
        Set p_outcome = Assert.AreEqual(p_expectedCount, p_actualCount, _
            "Installed card collection should have the expected number of cards.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_cardName = "7700"
        Dim p_cards As Collection
        Set p_cards = p_routeSystem.InstalledCards
        Set p_card = p_cards(1)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
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
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the installed card collection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedCapacity, p_card.Capacity, _
            "Card '" & p_cardName & "' should have the expected capacity.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedFirstChannel, p_card.DeviceFirstChannel, _
            "Card '" & p_cardName & "' should have the expected first channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedLastChannel, p_card.DeviceLastChannel, _
            "Card '" & p_cardName & "' should have the expected last channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedSlotNumber, p_card.SlotNumber, _
            "Card '" & p_cardName & "' should have the expected slot number.")
    End If
    
    p_cardName = "7702"
    p_expectedFirstChannel = 21
    p_expectedLastChannel = 60
    p_expectedCapacity = 40
    p_expectedSlotNumber = 2
    
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.InstalledCards(CStr(p_expectedSlotNumber))
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected from the installed card collection.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedCapacity, p_card.Capacity, _
            "Card '" & p_cardName & "' should have the expected capacity.")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedFirstChannel, p_card.DeviceFirstChannel, _
            "Card '" & p_cardName & "' should have the expected first channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedLastChannel, p_card.DeviceLastChannel, _
            "Card '" & p_cardName & "' should have the expected last channel number.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedSlotNumber, p_card.SlotNumber, _
            "Card '" & p_cardName & "' should have the expected slot number.")
    End If
    
    Debug.Print p_outcome.BuildReport("Test7700CardsShouldSelected")
    
    Set Test7700CardsShouldSelected = p_outcome
    
    
End Function

''' <summary>   Asserts building scan lists. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Assert7700CardsShouldBuildScanLists(ByVal a_senseFunction As String) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewViSession())
    
    Set p_outcome = Assert.IsNotNothing(p_routeSystem, _
        TypeName(p_routeSystem) & " should be instantiated.")

    Dim p_expectedCount As Integer
    p_expectedCount = 4
    Dim p_actualCount As Integer
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.Populate7700Cards
        Set p_outcome = Assert.AreEqual(p_expectedCount, p_routeSystem.InstrumentFamilyCards.Count, _
            "Instrument family card collection should have the expected number of cards.")
    End If
    
    Dim p_cardName As String
    Dim p_card As MultiplexerCard
    Dim p_channelNumber As Integer
    Dim p_options As String: p_options = "7700,7702"
    p_expectedCount = 2
    If p_outcome.AssertSuccessful Then
        p_actualCount = p_routeSystem.PopulateCards(p_options)
        Set p_outcome = Assert.AreEqual(p_expectedCount, p_actualCount, _
            "Installed card collection should have the expected number of cards.")
    End If
    
    ' buid the scan lists here so as to set the channel numbers properly.
    p_routeSystem.BuildFunctionScanLists a_senseFunction
    
    p_cardName = "7700"
    p_channelNumber = 1
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
    
    Dim p_expectedFunctionScanList As String
    p_expectedFunctionScanList = ":FUNC '" & a_senseFunction & "',(@101,1" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    Dim p_actualFunctionScanList As String
    p_actualFunctionScanList = p_card.FunctionScanList
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(p_expectedFunctionScanList, p_actualFunctionScanList, _
            "The expected scan list should be built for card '" & p_cardName & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_channelNumber = p_channelNumber + p_card.FunctionalCapacity - 1
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
    
    Dim p_expectedScanList As String
    p_expectedScanList = "(@1" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    Dim p_actualScanList As String
    If p_outcome.AssertSuccessful Then
        p_actualScanList = p_card.BuildChannelScanList(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_expectedScanList, p_actualScanList, _
            "The expected scan list should be returned for card '" & p_cardName & "' and channel " & CStr(p_channelNumber) & ".")
    End If
    
    Dim p_expectedRouteCommand As String
    p_expectedRouteCommand = ":ROUT:MULT:CLOS (@124,125)"
    Dim p_actualRouteCommand As String
    If p_outcome.AssertSuccessful Then
        p_actualRouteCommand = p_card.RouteMultipleCloseCommand()
        Set p_outcome = Assert.AreEqual(p_expectedRouteCommand, p_actualRouteCommand, _
            "The expected route command should be returned for card '" & p_cardName & "'.")
    End If
    
    p_cardName = "7702"
    p_channelNumber = p_channelNumber + 1
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_actualFunctionScanList = p_card.FunctionScanList
        p_expectedFunctionScanList = ":FUNC '" & a_senseFunction & "',(@201,2" & VBA.CStr(p_card.FunctionalCapacity) & ")"
        Set p_outcome = Assert.AreEqual(p_expectedFunctionScanList, p_actualFunctionScanList, _
            "The expected scan list should be built for card '" & p_cardName & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_channelNumber = p_channelNumber + p_card.FunctionalCapacity - 1
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.IsNotNothing(p_card, _
            "A card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
        
    If p_outcome.AssertSuccessful Then
        Set p_card = p_routeSystem.SelectMultiplexerCard(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_cardName, p_card.Name, _
            "The expected card should be selected for channel " & CStr(p_channelNumber) & ".")
    End If
    
    p_expectedScanList = "(@2" & VBA.CStr(p_card.FunctionalCapacity) & ")"
    If p_outcome.AssertSuccessful Then
        p_actualScanList = p_card.BuildChannelScanList(p_channelNumber)
        Set p_outcome = Assert.AreEqual(p_expectedScanList, p_actualScanList, _
            "The expected scan list should be returned for card '" & p_cardName & "' and channel " & CStr(p_channelNumber) & ".")
    End If
    
    p_expectedRouteCommand = ":ROUT:MULT:CLOS (@244,245)"
    If p_outcome.AssertSuccessful Then
        p_actualRouteCommand = p_card.RouteMultipleCloseCommand()
        Set p_outcome = Assert.AreEqual(p_expectedRouteCommand, p_actualRouteCommand, _
            "The expected route command should be returned for card '" & p_cardName & "'.")
    End If
    
    Set Assert7700CardsShouldBuildScanLists = p_outcome
    
End Function


''' <summary>   Unit test. Asserts building scan lists. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Test7700CardsShouldBuildScanLists() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = Assert7700CardsShouldBuildScanLists("RES")

    Debug.Print p_outcome.BuildReport("Test7700CardsShouldBuildScanLists")
    
    Set Test7700CardsShouldBuildScanLists = p_outcome

End Function

''' <summary>   Unit test. Asserts building 4-wire scan lists. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Test7700CardsShouldBuild4WireScanLists() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = Assert7700CardsShouldBuildScanLists("FRES")

    Debug.Print p_outcome.BuildReport("Test7700CardsShouldBuild4WireScanLists")
    
    Set Test7700CardsShouldBuild4WireScanLists = p_outcome

End Function













