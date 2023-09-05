Attribute VB_Name = "ScpiSystemTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Scpi System Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    K2700 As cc_isr_Tcp_Scpi.K2700
    Host As String
    Port As Long
    SocketReceiveTimeout As Integer
    ErrTracer As IErrTracer
    DeviceErrorsTracer As IErrTracer
End Type

Private This As this_

Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestInputsShouldBeFront
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
    For p_testNumber = 1 To 1
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
    
    Dim p_deviceErrorsTracer As New DeviceErrorsTracer
    Set This.DeviceErrorsTracer = p_deviceErrorsTracer.Initialize(This.K2700)
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700().Initialize(This.ErrTracer)
    
    ' trap errors in case connection fails rendering all tests inconclusive.
    
    On Error Resume Next
    
    This.K2700.OpenConnection This.Host, This.Port, This.SocketReceiveTimeout
    
    Dim p_leftoverErrorMessage As String
    p_leftoverErrorMessage = VBA.vbNullString
    
    If Err.Number <> 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf This.K2700.Connected Then
        Set This.BeforeAllAssert = Assert.IsTrue(True, "Connected")
    Else
        Set This.BeforeAllAssert = Assert.Inconclusive("K2700 should be connected")
    End If
    
    This.ErrTracer.TraceError p_leftoverErrorMessage
    
    ' clear the error object.
    
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    Set This.BeforeEachAssert = Assert.IsTrue(True, "initialize the pre-test assert.")

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = IIf(This.K2700.Connected, _
            Assert.IsTrue(True, "Connected"), _
            Assert.Inconclusive("K2700 should be connected"))
    
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
        This.K2700.Device.ClearExecutionState
    
End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    
    Set This.ErrTracer = Nothing
    Set This.DeviceErrorsTracer = Nothing
    
    ' disconnect if connected
    If Not This.K2700 Is Nothing Then _
        This.K2700.CloseConnection

    If Not This.K2700 Is Nothing Then This.K2700.Dispose
    Set This.K2700 = Nothing

    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts inputs should be front. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInputsShouldBeFront() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNotNothing(This.K2700.ScpiSystem, _
            "Scpi System should be instantiated.")
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = Assert.IsTrue(This.K2700.ScpiSystem.QueryFrontSwitch(), _
            "Scpi System should query and report the correct state of the front switch.")

    End If
    
    Dim p_deviceErrorAssert As cc_isr_Test_Fx.Assert
    Set p_deviceErrorAssert = This.DeviceErrorsTracer.AssertLeftoverErrors

    If p_outcome.AssertSuccessful Then
        Set p_outcome = p_deviceErrorAssert
    ElseIf Not p_deviceErrorAssert.AssertSuccessful Then
        Set p_outcome = Assert.Fail(p_outcome.AssertMessage & VBA.vbCrLf & _
        "Device errors: " & VBA.vbCrLf & p_deviceErrorAssert.AssertMessage)
    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors

    Debug.Print p_outcome.BuildReport("TestInputsShouldBeFront")
    
    Set TestInputsShouldBeFront = p_outcome
    
End Function

