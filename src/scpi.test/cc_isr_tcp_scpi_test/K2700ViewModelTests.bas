Attribute VB_Name = "K2700ViewModelTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 View Model Tests. </summary>
''' <remarks>   Dependencies: cc_isr_Core_Tcp_Scpi.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    BeforeEachAssert As Assert
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    Host As String
    Port As Long
    ErrTracer As ErrTracer
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    SenseFunction As String
End Type

Private This As this_

Public Sub BeforeAll()

    ' initialize known data.
    This.TopCard = "7700"
    This.BottomCard = VBA.vbNullString
    This.SenseFunction = "FRES"
    This.TopCardFunctionScanList = ":FUNC 'FRES',(@101,120)"
    This.BottomCardFunctionScanList = VBA.vbNullString
    
    Set This.ViewModel = cc_isr_Tcp_Scpi.K2700ViewModel
    This.ViewModel.Host = "192.168.0.252"
    This.ViewModel.Port = 1234
    This.ViewModel.SocketReceiveTimeout = 100
    This.ViewModel.SenseFunctionName = This.SenseFunction
    Set This.ErrTracer = New ErrTracer
    
    ' initialize the view model.
    This.ViewModel.Initialize This.ErrTracer

    ' trap errors in case connection fails rendering all tests inconclusive.
    On Error Resume Next
    ' connect
    This.ViewModel.ToggleConnectionCommand True
    On Error GoTo 0

End Sub

Public Sub BeforeEach()

    Set This.BeforeEachAssert = IIf(This.ViewModel.Connected, Assert.IsTrue(True, "Connected"), _
                        Assert.Inconclusive("View Model should be connected"))
                        
    This.ViewModel.LastErrorMessage = VBA.vbNullString
    If This.BeforeEachAssert.AssertSuccessful Then _
        This.ViewModel.ClearExecutionStateCommand

End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    
    ' disconnect if connected
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.ToggleConnectionCommand False

    If Not This.ViewModel Is Nothing Then This.ViewModel.Dispose
    Set This.ViewModel = Nothing

End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Connection Tests
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Unit test. Asserts that view model should initialize. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestViewModelShouldInitialize() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.ToggleConnectionExecutable, _
            "Toggle connection should be executable after initializing the View Model")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Exception: " & This.ViewModel.LastErrorMessage)
        
    Debug.Print p_outcome.BuildReport("TestViewModelShouldInitialize")
    
    Set TestViewModelShouldInitialize = p_outcome

End Function

''' <summary>   Unit test. Asserts that view model should connect. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestViewModelShouldConnect() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(This.ViewModel.Connected, _
            "View model should connect")
        
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, This.ViewModel.LastErrorMessage, _
            "Exception: " & This.ViewModel.LastErrorMessage)
    
    Debug.Print p_outcome.BuildReport("TestViewModelShouldConnect")
    
    Set TestViewModelShouldConnect = p_outcome

End Function

''' <summary>   Unit test. Asserts that view model should read cards. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestViewModelShouldReadCards() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.TopCard, This.ViewModel.TopCard, _
            "View Model should be read the top card")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.BottomCard, This.ViewModel.BottomCard, _
            "View Model should be read the bottom card")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.SenseFunction, _
            This.ViewModel.SenseFunctionName, _
            "View Model should set the sense function name")

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.TopCardFunctionScanList, _
            This.ViewModel.TopCardFunctionScanList, _
            "View Model should be read the top card function scan list")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(This.BottomCardFunctionScanList, _
            This.ViewModel.BottomCardFunctionScanList, _
            "View Model should be read the top card function scan list")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(VBA.vbNullString, _
            This.ViewModel.LastErrorMessage, _
            "Exception: " & This.ViewModel.LastErrorMessage)

    Debug.Print p_outcome.BuildReport("TestViewModelShouldReadCards")
    
    Set TestViewModelShouldReadCards = p_outcome

End Function


Public Sub RunTests()
    BeforeAll
    BeforeEach
    TestParsingDeviceError
    AfterEach
    AfterAll
End Sub
