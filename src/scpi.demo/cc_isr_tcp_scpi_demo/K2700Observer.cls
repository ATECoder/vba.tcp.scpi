VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700Observer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The K2700 Observer for interfacing the K2700 view with the View Model. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_byRow As Boolean = True
Private Const m_displayIncrement As Integer = 1
Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 14
Private Const m_channelRow As Integer = 2
Private Const m_channelCol As Integer = 13

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_

    ' Observer properties
    K2700Sheet As K2700Sheet
    DesignMode As Boolean
        
    ' Initial User inputs
    SocketAddress As String
    ReadingOffset As Double
    GpibLanControllerPort As Long
    
    ' view model properties
    ViewModel As K2700ViewModel
    ElapsedTime As Double
    ElapsedTimeFormat As String
    ElapsedTimeCaption As String
    LastErrorMessage As String
    LastMessage As String
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    Measuring As Boolean
    MaximumChannelNumber As Integer

    ' Winsock properties
    Host As String
    Port As Long
    SocketId As Integer
    SessionTimeout As Integer
    Connected As Boolean

    ' ieee488 control properties
    MessageToSend As String
    SentMessage As String
    ReceivedMessage As String
    ReceivedMessageLength As Integer
    StatusByte As Integer
    StandardByte As Integer
    ReadStatusAfterWriteEnabled As Boolean
    ScpiCommands As Collection
    
    ' GPIB-Lan properties.
    ServiceRequested As Boolean
    SerialPollByte As Integer
    GpibAddress As Integer
    AutoAssertTalk As Boolean
    GpibLanReadTimeout As Integer
    
    ' 2700 properties
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    ReadingBeepEnabled As Boolean
    FrontInputsRequired As Boolean
    FrontInputsValue As Boolean
    FrontInputsCaption As String
    ChannelCount As Integer
    SelectedChannelNumber As Integer
    TargetChannelNumber As Integer
    MeasuredChannelNumber As Integer
    ContinuousSenseFunctionName As String
    ImmediateSenseFunctionName As String
    ExternalSenseFunctionName As String
    SenseFunctionName As String
    SenseFunctionCaption As String
    SenseFunctionUnit As String
    MeasuredReading As String
    MeasuredValueFormat As String
    MeasuredValue As Double
    SenseFunctions As Collection
    
    MeasurementRow As Integer
    MeasurementColumn As Integer
    ChannelRow As Integer
    ChannelColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String
    ChannelNumberCaptionPrefix As String
    
    ' service request monitoring timer control
    ExternalTrigMonitoringEnabled As Boolean
    StopRequested As Boolean
    PauseRequested As Boolean
    TimerInterval As Integer
    ' true when handling the timer event to prevent timer
    ' event handler overrun. This effectively synchronizes the timer
    ' event with the instrument.
    HandlingTimerEvent As Boolean

    ' Service request monitoring properties
    MonitoringState As String
    AutoIncrementChannelNoEnabled As Boolean
    SingleReadEnabled As Boolean

    ' Winsock connection executables
    OpenConnectionExecutable As Boolean
    CloseConnectionExecutable As Boolean
    ToggleConnectionExecutable As Boolean
    
    ' IEEE488 executables
    ClearExecutionStateExecutable As Boolean
    MessageToSendExecutable As Boolean
    QueryExecutable As Boolean
    ReadExecutable As Boolean
    QueryStatusByteExecutable As Boolean
    ReadStatusAfterWriteExecutable As Boolean
    ReadStandardEventStatusExecutable As Boolean
    ResetKnownStateExecutable As Boolean
    ServiceRequestedExecutable As Boolean
    WriteExecutable As Boolean
    
    ' 2700 executables
    ClearReadingsExecutable As Boolean
    FunctionNameExecutable As Boolean
    ImmediateTriggerOptionExecutable As Boolean
    MeasureExecutable As Boolean
    QueryInputsExecutable As Boolean
    
    ' Service request monitoring properties
    StopMonitoringExecutable As Boolean
    StartMonitoringExecutable As Boolean
    ExternalTriggerOptionExecutable As Boolean

    ' GPIB-Lan executables
    SelectiveDeviceClearExecutable As Boolean
    GoToLocalExecutable As Boolean
    LocalLockoutExecutable As Boolean
    AutoAssertTalkSetExecutable As Boolean
    AutoAssertTalkGetExecutable As Boolean
    ReadStatusByteExecutable As Boolean
    GpibAddressGetExecutable As Boolean
    GpibAddressSetExecutable As Boolean
    GpibLanReadTimeoutGetExecutable As Boolean
    GpibLanReadTimeoutSetExecutable As Boolean
        
    LastCommandDetails As String
    LastCommandOutcome As Boolean
    
End Type

Private This As this_


' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
    This.TopCard = VBA.vbNullString
    This.ChannelNumberCaptionPrefix = "R"
End Sub

''' <summary>   Handles the Class Terminate event. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>    Gets the Initialization condition of this object. </summary>
''' <value>   [Boolean] True if initialized. </value>
Public Property Get Initialized() As Boolean
    Initialized = Not This.ViewModel Is Nothing
End Property

''' <summary>   Initialize this worksheet if not initialized. </summary>
Friend Function Initialize(ByVal a_viewModel As K2700ViewModel) As K2700Observer

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.K2700Sheet = K2700Sheet.Initialize(a_viewModel)
    
    Set This.ViewModel = a_viewModel
    
    Me.Synchronize
    
    ' set reference to the view model event handler.
    
    Set m_viewModel = a_viewModel
    
    Set Initialize = Me

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.K2700Sheet = Nothing
    
    Set m_viewModel = Nothing
    
    Set This.ViewModel = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' User interface control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the design mode for unit testing and restoration of state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get DesignMode() As Boolean
    
    If This.K2700Sheet Is Nothing Then
        DesignMode = This.DesignMode
    Else
        DesignMode = This.K2700Sheet.DesignMode
    End If

End Property

''' <summary>   Sets the design mode for unit testing and restoration of state. </summary>
''' <param name="a_value">   [Boolean] true if in design mode. </param>
Public Property Let DesignMode(ByVal a_value As Boolean)
    If Me.DesignMode <> a_value Then
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.DesignMode = a_value
        This.DesignMode = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Toggle connection button click event. </summary>
Public Sub OnConnect()

    If This.ToggleConnectionExecutable Then
        If This.OpenConnectionExecutable Then
            This.ViewModel.OpenConnectionCommand
        ElseIf This.CloseConnectionExecutable Then
            This.ViewModel.CloseConnectionCommand
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnConnect", _
                " Both close and open connection are disabled."
        End If
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnConnect", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Open Connection button Click event. </summary>
Public Sub OnOpenConnectionButtonClick()

    Const p_procedureName As String = "OnOpenConnectionButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if already connected
    If This.ViewModel.Connected Then Exit Sub
        
    This.ViewModel.OpenConnectionCommand
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Open Connection command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    If This.K2700Sheet Is Nothing Then
        OpenConnectionExecutable = This.OpenConnectionExecutable
    Else
        OpenConnectionExecutable = This.K2700Sheet.OpenConnectionExecutable
    End If
End Property

''' <summary>   Gets the condition indicating that the toggle Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Connection command can be executed. </value>
Public Property Get ToggleConnectionExecutable() As Boolean
    ToggleConnectionExecutable = This.OpenConnectionExecutable Or This.CloseConnectionExecutable
End Property


''' <summary>   Sets the execution (Can Execute) status of the Open Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    If Me.OpenConnectionExecutable <> a_value Then
        
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.OpenConnectionExecutable = a_value
        
        This.OpenConnectionExecutable = a_value
    
    End If
End Property

''' <summary>   Handles the Close Connection button Click event. </summary>
Public Sub OnCloseConnectionButtonClick()

    Const p_procedureName As String = "OnCloseConnectionButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if already disconnected
    If Not This.ViewModel.Connected Then Exit Sub
        
    This.ViewModel.CloseConnectionCommand
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Gets the execution (Can Execute) status of the Close Connection command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    If This.K2700Sheet Is Nothing Then
        CloseConnectionExecutable = This.CloseConnectionExecutable
    Else
        CloseConnectionExecutable = This.K2700Sheet.CloseConnectionExecutable
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Close Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    If Me.CloseConnectionExecutable <> a_value Then
        
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.CloseConnectionExecutable = a_value
        
        This.CloseConnectionExecutable = a_value
    
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Public Sub OnClearExecutionState()
    If This.ClearExecutionStateExecutable Then
        Me.LastCommandOutcome = This.ViewModel.ClearExecutionStateCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnClearExecutionState", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Query button click event. </summary>
Public Sub OnQuery()
    If This.QueryExecutable Then
        This.ViewModel.QueryCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnQuery", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Read button click event. </summary>
Public Sub OnRead()
    If This.ReadExecutable Then
        This.ViewModel.ReadCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnRead", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the Query Status Byte button click event. </summary>
Public Sub OnQueryStatusByte()
    If This.QueryStatusByteExecutable Then
        This.ViewModel.QueryStatusByteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnQueryStatusByte", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Read Standard Event Status Byte button click event. </summary>
Public Sub OnReadStandardEventStatus()
    If This.ReadStandardEventStatusExecutable Then
        This.ViewModel.ReadStandardEventStatusCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnReadStandardEventStatus", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Public Sub OnResetKnownState()

    If This.ResetKnownStateExecutable Then
        Dim p_details As String
        If Not This.ViewModel.ResetKnownStateCommand(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnResetKnownState", _
                " Command failed; " & p_details
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnResetKnownState", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Public Sub OnWrite()
    If This.WriteExecutable Then
        This.ViewModel.WriteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnWrite", _
            " Command not enabled."
    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Go To Local button click event. </summary>
Public Sub OnGoToLocal()
    If This.GoToLocalExecutable Then
        This.ViewModel.GoToLocalCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnGoToLocal", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Public Sub OnGpibAddressGet()
    
    If This.GpibAddressGetExecutable Then
        This.ViewModel.GpibAddressGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnGpibAddressGet", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Public Sub OnGpibAddressSet()
    If This.GpibAddressSetExecutable Then
        This.ViewModel.GpibAddressSetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnGpibAddressSet", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Public Sub OnLocalLockout()
    If This.LocalLockoutExecutable Then
        This.ViewModel.LocalLockoutCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnLocalLockout", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Auto Asset TALK Set button click event. </summary>
Public Sub OnAutoAssertTalkSet()
    If This.AutoAssertTalkSetExecutable Then
        This.ViewModel.AutoAssertTalkSetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnAutoAssertTalkSet", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Auto Asset TALK get button click event. </summary>
Public Sub OnAutoAssertTalkGet()
    If This.AutoAssertTalkGetExecutable Then
        This.ViewModel.AutoAssertTalkGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnAutoAssertTalkGet", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Public Sub OnReadTimeoutGet()
    If This.GpibLanReadTimeoutGetExecutable Then
        This.ViewModel.GpibLanReadTimeoutGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnReadTimeoutGet", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Public Sub OnReadTimeoutSet()
    If This.GpibLanReadTimeoutGetExecutable Then
        This.ViewModel.GpibLanReadTimeoutGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnReadTimeoutSet", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Public Sub OnSDC()
    If This.SelectiveDeviceClearExecutable Then
        This.ViewModel.SelectiveDeviceClearCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnSDC", _
            " Command not enabled."
    End If
   
End Sub

''' <summary>   Handles the Read Status Byte (serial poll) button click event. </summary>
Public Sub OnReadStatusByte()
    If This.ReadStatusByteExecutable Then
        This.ViewModel.ReadStatusByteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnReadStatusByte", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Public Sub OnSRQ()
    If This.ServiceRequestedExecutable Then
        This.ViewModel.ServiceRequestedQueryCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnSRQ", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the power on reset button click event. </summary>
Public Sub OnPowerOneResetButtonClick()

    Const p_procedureName As String = "OnPowerOneResetButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' issue a power on reset.
    This.ViewModel.PowerOnResetCommand
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Measurments View
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Public Sub OnClearReadings()
    If This.ClearReadingsExecutable Then
        This.ViewModel.ClearReadingsCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnClearReadings", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Clear Messages Button Click Event. </summary>
Public Sub OnClearMessages()
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.LastErrorMessage = VBA.vbNullString
        This.ViewModel.LastMessage = VBA.vbNullString
    End If
End Sub

''' <summary>   Gets the Measurement Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementRow() As Integer
    MeasurementRow = This.MeasurementRow
End Property

''' <summary>   Sets the Measurement Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Row. </value>
Public Property Let MeasurementRow(ByVal a_value As Integer)
    If This.MeasurementRow <> a_value Then
        This.MeasurementRow = a_value
    End If
End Property

''' <summary>   Gets the Measurement Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementColumn() As Integer
    MeasurementColumn = This.MeasurementColumn
End Property

''' <summary>   Sets the Measurement Column. </summary>
''' <param name="a_value">   [Integer] the Measurement Column. </value>
Public Property Let MeasurementColumn(ByVal a_value As Integer)
    If This.MeasurementColumn <> a_value Then
        This.MeasurementColumn = a_value
    End If
End Property

''' <summary>   Gets the Channel Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelRow() As Integer
    ChannelRow = This.ChannelRow
End Property

''' <summary>   Sets the Channel Row. </summary>
''' <param name="a_value">   [Integer] the Channel Row. </value>
Public Property Let ChannelRow(ByVal a_value As Integer)
    If This.ChannelRow <> a_value Then
        This.ChannelRow = a_value
    End If
End Property

''' <summary>   Gets the Channel Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelColumn() As Integer
    ChannelColumn = This.ChannelColumn
End Property

''' <summary>   Sets the Channel Column. </summary>
''' <param name="a_value">   [Integer] the Channel Column. </value>
Public Property Let ChannelColumn(ByVal a_value As Integer)
    If This.ChannelColumn <> a_value Then
        This.ChannelColumn = a_value
    End If
End Property

''' <summary>   Gets the Measurement Cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementCellValue() As String
    If This.K2700Sheet Is Nothing Then
        MeasurementCellValue = This.MeasurementCellValue
    Else
        MeasurementCellValue = This.K2700Sheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
    End If
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If Me.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
    End If
End Property
        
''' <summary>   Gets the channel cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCellValue() As String
    If This.K2700Sheet Is Nothing Then
        ChannelCellValue = This.ChannelCellValue
    Else
        ChannelCellValue = This.K2700Sheet.Cells(Me.ChannelRow, Me.ChannelColumn).Value
    End If
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If Me.ChannelCellValue <> a_value Then
        
        This.ChannelCellValue = a_value
        
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.Cells(Me.ChannelRow, Me.ChannelColumn).Value = a_value
        
    End If
End Property

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Public Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)
    If Not This.K2700Sheet Is Nothing Then _
        SetReading_ a_channelNo, a_value, m_byRow, m_displayIncrement
End Sub

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_byRow">       [Boolean] true if the next reading goes to the next row. </param>
''' <param name="a_increment">   [Integer] the row or column increment. </param>
Private Sub SetReading_(ByVal a_channelNo As Integer, ByVal a_value As Double, _
    ByVal a_byRow As Boolean, ByVal a_increment As Integer)

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_channelNo - 1), _
        m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, _
        m_measurementCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelRow = IIf(a_byRow, m_channelRow + a_increment * (a_channelNo - 1), _
        m_channelRow)
    Me.ChannelColumn = IIf(a_byRow, m_channelCol, _
        m_channelCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelCellValue = This.ChannelNumberCaptionPrefix & VBA.CStr(a_channelNo)
    Me.MeasurementCellValue = VBA.CStr(a_value)
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the external trigger option click events. </summary>
Public Sub OnExternalTriggerOption()
    If This.ExternalTriggerOptionExecutable Then
        Me.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".OnExternalTriggerOptionChange", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
Public Sub OnImmediateTriggerOption()

    If This.ImmediateTriggerOptionExecutable Then
        Me.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnImmediateTriggerOptionChange", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Public Sub OnMeasure()
    If This.MeasureExecutable Then
        ' make an immediate measure of one or all channels.
        Me.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnMeasure", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the start monitoring external trigger button click. </summary>
Public Sub OnStartMonitorExtTrigButtonClick()
    This.ViewModel.StartMonitoringExternalTriggersCommand
End Sub

''' <summary>   Handles the stop monitoring external trigger button click. </summary>
Public Sub OnStopMonitorExtTrigButtonClick()
    This.ViewModel.StopMonitoringExternalTriggersCommand
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Public Sub OnMonitorExtTrigToggleButtonChange(ByVal a_value As Boolean)
    If This.StartMonitoringExecutable Then
        This.ViewModel.StartMonitoringExternalTriggersCommand
    ElseIf This.StopMonitoringExecutable Then
        This.ViewModel.StopMonitoringExternalTriggersCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnMonitorExtTrigToggleButtonChange", _
            " Command not enabled."
    End If
End Sub

''' <summary> Queries the Inputs status. </summary>
Public Sub OnQueryInputs()
    If This.QueryInputsExecutable Then
        This.ViewModel.QueryInputsCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnQueryInputs", _
            " Command not enabled."
    End If
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the INET Address dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Sets the Host. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let Host(ByVal a_value As String)
    If This.Host <> a_value Then
        This.Host = a_value
    End If
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Long
    Port = This.Port
End Property

''' <summary>   Sets the Port. </summary>
''' <param name="a_value">   [Long] the Port. </value>
Public Property Let Port(ByVal a_value As Long)
    If This.Port <> a_value Then
        This.Port = a_value
    End If
End Property

''' <summary>   Gets the Connected state. </summary>
''' <value>   [Integer]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property

''' <summary>   Sets the Connected state. </summary>
''' <param name="a_value">   [Boolean] the Connected state. </value>
Public Property Let Connected(ByVal a_value As Boolean)
    If This.Connected <> a_value Then
        This.Connected = a_value
    End If
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_value As String)
    If This.SocketAddress <> a_value Then
        This.SocketAddress = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SocketAddress = a_value
    End If
End Property

''' <summary>   Gets the socket id. </summary>
''' <value>   [Integer]. </value>
Public Property Get SocketId() As Integer
    SocketId = This.SocketId
End Property

''' <summary>   Sets the socket id. </summary>
''' <param name="a_value">   [Integer] the socket id. </value>
Public Property Let SocketId(ByVal a_value As Integer)
    If This.SocketId <> a_value Then
        This.SocketId = a_value
    End If
End Property

''' <summary>   Gets the socket and session timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket and session timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_value As Integer)
    If This.SessionTimeout <> a_value Then
        This.SessionTimeout = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SessionTimeout = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Message to Sent. </summary>
''' <value>   [String]. </value>
Public Property Get MessageToSend() As String
    MessageToSend = This.MessageToSend
End Property

''' <summary>   Sets the Message to Sent. </summary>
''' <param name="a_value">   [String] the Message to Sent. </value>
Public Property Let MessageToSend(ByVal a_value As String)
    If This.MessageToSend <> a_value Then
        This.MessageToSend = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MessageToSend = a_value
    End If
End Property

''' <summary>   Gets the Sent Message. </summary>
''' <value>   [String]. </value>
Public Property Get SentMessage() As String
    SentMessage = This.SentMessage
End Property

''' <summary>   Sets the Sent Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SentMessage(ByVal a_value As String)
    If This.SentMessage <> a_value Then
        This.SentMessage = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SentMessage = a_value
    End If
End Property

''' <summary>   Gets the Received Message. </summary>
''' <value>   [String]. </value>
Public Property Get ReceivedMessage() As String
    ReceivedMessage = This.ReceivedMessage
End Property

''' <summary>   Sets the Received Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let ReceivedMessage(ByVal a_value As String)
    If This.ReceivedMessage <> a_value Then
        This.ReceivedMessage = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReceivedMessage = a_value
    End If
End Property

''' <summary>   Gets the Received Message Length. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceivedMessageLength() As Integer
    ReceivedMessageLength = This.ReceivedMessageLength
End Property

''' <summary>   Sets the Received Message Length. </summary>
''' <param name="a_value">   [Integer] the Message Length value. </value>
Public Property Let ReceivedMessageLength(ByVal a_value As Integer)
    If This.ReceivedMessageLength <> a_value Then
        This.ReceivedMessageLength = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReceivedMessageLength = a_value
    End If
End Property

''' <summary>   Gets the Status Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StatusByte() As Integer
    StatusByte = This.StatusByte
End Property

''' <summary>   Sets the Status Byte. </summary>
''' <param name="a_value">   [Integer] the Status Byte value. </value>
Public Property Let StatusByte(ByVal a_value As Integer)
    If This.StatusByte <> a_value Then
        This.StatusByte = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.StatusByte = a_value
    End If
End Property

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    StandardByte = This.StandardByte
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Public Property Let StandardByte(ByVal a_value As Integer)
    If This.StandardByte <> a_value Then
        This.StandardByte = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.StandardByte = a_value
    End If
End Property

''' <summary>   Gets the Auto Status Read Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadStatusAfterWriteEnabled() As Boolean
    ReadStatusAfterWriteEnabled = This.ReadStatusAfterWriteEnabled
End Property

''' <summary>   Sets the Auto Status Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Status Read Enabled value. </value>
Public Property Let ReadStatusAfterWriteEnabled(ByVal a_value As Boolean)
    If This.ReadStatusAfterWriteEnabled <> a_value Then
        This.ReadStatusAfterWriteEnabled = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadStatusAfterWriteEnabled = a_value
    End If
End Property

''' <summary>   Gets the Scpi Commands. </summary>
''' <value>   [Collection]. </value>
Public Property Get ScpiCommands() As Collection
    Set ScpiCommands = This.ScpiCommands
End Property

''' <summary>   Sets the Scpi Commands. </summary>
''' <param name="a_value">   [Collection] the Scpi Commands value. </value>
Public Property Let ScpiCommands(ByVal a_value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.ScpiCommands, a_value) Then
        Set This.ScpiCommands = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ScpiCommands = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Serial Poll Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get SerialPollByte() As Integer
    SerialPollByte = This.SerialPollByte
End Property

''' <summary>   Sets the Serial Poll Byte. </summary>
''' <param name="a_value">   [Integer] the Serial Poll Byte value. </value>
Public Property Let SerialPollByte(ByVal a_value As Integer)
    If This.SerialPollByte <> a_value Then
        This.SerialPollByte = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SerialPollByte = a_value
    End If
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_value As Integer)
    If This.ServiceRequested <> a_value Then
        This.ServiceRequested = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ServiceRequested = a_value
    End If
End Property

''' <summary>   Gets the Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibAddress() As Integer
    GpibAddress = This.GpibAddress
End Property

''' <summary>   Sets the Gpib Address. </summary>
''' <param name="a_value">   [Integer] the Gpib Address value. </value>
Public Property Let GpibAddress(ByVal a_value As Integer)
    If This.GpibAddress <> a_value Then
        This.GpibAddress = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibAddress = a_value
    End If
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let AutoAssertTalk(ByVal a_value As Boolean)
    If This.AutoAssertTalk <> a_value Then
        This.AutoAssertTalk = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.AutoAssertTalk = a_value
    End If
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    If This.GpibLanControllerPort <> a_value Then
        This.GpibLanControllerPort = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibLanControllerPort = a_value
    End If
End Property

''' <summary>   Gets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanReadTimeout() As Integer
    GpibLanReadTimeout = This.GpibLanReadTimeout
End Property

''' <summary>   Sets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let GpibLanReadTimeout(ByVal a_value As Integer)
    If This.GpibLanReadTimeout <> a_value Then
        This.GpibLanReadTimeout = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibLanReadTimeout = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Top Card name. </summary>
''' <value>   [String]. </value>
Public Property Get TopCard() As String
    TopCard = This.TopCard
End Property

''' <summary>   Sets the Top Card name. </summary>
''' <param name="a_value">   [String] the Top Card name. </value>
Public Property Let TopCard(ByVal a_value As String)
    If This.TopCard <> a_value Then
        This.TopCard = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TopCard = a_value
    End If
End Property

''' <summary>   Gets the Top Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get TopCardFunctionScanList() As String
    TopCardFunctionScanList = This.TopCardFunctionScanList
End Property

''' <summary>   Sets the Top Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Top Card function scan list. </value>
Public Property Let TopCardFunctionScanList(ByVal a_value As String)
    If This.TopCardFunctionScanList <> a_value Then
        This.TopCardFunctionScanList = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TopCardFunctionScanList = a_value
    End If
End Property

''' <summary>   Gets the Bottom Card name. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCard() As String
    BottomCard = This.BottomCard
End Property

''' <summary>   Sets the Bottom Card name. </summary>
''' <param name="a_value">   [String] the Bottom Card name. </value>
Public Property Let BottomCard(ByVal a_value As String)
    If This.BottomCard <> a_value Then
        This.BottomCard = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.BottomCard = a_value
    End If
End Property

''' <summary>   Gets the Bottom Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCardFunctionScanList() As String
    BottomCardFunctionScanList = This.BottomCardFunctionScanList
End Property

''' <summary>   Sets the Bottom Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Bottom Card function scan list. </value>
Public Property Let BottomCardFunctionScanList(ByVal a_value As String)
    If This.BottomCardFunctionScanList <> a_value Then
        This.BottomCardFunctionScanList = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.BottomCardFunctionScanList = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Data Acquisition properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the maximum Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumChannelNumber() As Integer
    MaximumChannelNumber = This.MaximumChannelNumber
End Property

''' <summary>   Sets the maximum Channel Number. </summary>
''' <param name="a_value">   [Integer] the maximum Channel Number value. </value>
Public Property Let MaximumChannelNumber(ByVal a_value As Integer)
    If This.MaximumChannelNumber <> a_value Then
        This.MaximumChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MaximumChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the measuring state. True when measurements are collected. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Measuring() As Boolean
    Measuring = This.Measuring
End Property

''' <summary>   Sets the measuring state. True when measurements are collected. </summary>
''' <param name="a_value">   [Boolean] the measuring state; true when measuring is active. </value>
Public Property Let Measuring(ByVal a_value As Boolean)
    If This.Measuring <> a_value Then
        This.Measuring = a_value
    End If
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get MeasurementMode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    MeasurementMode = This.MeasurementMode
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Let MeasurementMode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_value Then
        This.MeasurementMode = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Required. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsRequired() As Boolean
    FrontInputsRequired = This.FrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are Required. </value>
Public Property Let FrontInputsRequired(ByVal a_value As Boolean)
    If This.FrontInputsRequired <> a_value Then
        This.FrontInputsRequired = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsRequired = a_value
    End If
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <remarks>   True if instrument beep sound on each reading is enabled. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ReadingBeepEnabled() As Boolean
    ReadingBeepEnabled = This.ReadingBeepEnabled
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <param name="a_value">   [Boolean] instrument beep sound on each reading is enabled. </value>
Public Property Let ReadingBeepEnabled(ByVal a_value As Boolean)
    If This.ReadingBeepEnabled <> a_value Then
        This.ReadingBeepEnabled = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadingBeepEnabled = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Value. </summary>
''' <remarks>   True if front inputs; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsValue() As Boolean
    FrontInputsValue = This.FrontInputsValue
End Property

''' <summary>   Sets the Front Inputs Value. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs. </value>
Public Property Let FrontInputsValue(ByVal a_value As Boolean)
    If This.FrontInputsValue <> a_value Then
        This.FrontInputsValue = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsValue = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Caption. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsCaption() As String
    FrontInputsCaption = This.FrontInputsCaption
End Property

''' <summary>   Sets the Front Inputs Caption. </summary>
''' <param name="a_value">   [String] the Front Inputs Caption. </value>
Public Property Let FrontInputsCaption(ByVal a_value As String)
    If This.FrontInputsCaption <> a_value Then
        This.FrontInputsCaption = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsCaption = a_value
    End If
End Property

''' <summary>   Gets the Channel Count. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelCount() As Integer
    ChannelCount = This.ChannelCount
End Property

''' <summary>   Sets the Channel Count. </summary>
''' <param name="a_value">   [Integer] the Channel Count value. </value>
Public Property Let ChannelCount(ByVal a_value As Integer)
    If This.ChannelCount <> a_value Then
        This.ChannelCount = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ChannelCount = a_value
    End If
End Property

''' <summary>   Gets the Selected Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedChannelNumber() As Integer
    SelectedChannelNumber = This.SelectedChannelNumber
End Property

''' <summary>   Sets the Selected Channel Number. </summary>
''' <param name="a_value">   [Integer] the Selected Channel Number value. </value>
Public Property Let SelectedChannelNumber(ByVal a_value As Integer)
    If This.SelectedChannelNumber <> a_value Then
        This.SelectedChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SelectedChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the Target Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get TargetChannelNumber() As Integer
    TargetChannelNumber = This.TargetChannelNumber
End Property

''' <summary>   Sets the Target Channel Number. </summary>
''' <param name="a_value">   [Integer] the Target Channel Number value. </value>
Public Property Let TargetChannelNumber(ByVal a_value As Integer)
    If This.TargetChannelNumber <> a_value Then
        This.TargetChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TargetChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the Measured Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredChannelNumber() As Integer
    MeasuredChannelNumber = This.MeasuredChannelNumber
End Property

''' <summary>   Sets the Measured Channel Number. </summary>
''' <param name="a_value">   [Integer] the Measured Channel Number value. </value>
Public Property Let MeasuredChannelNumber(ByVal a_value As Integer)
    If This.MeasuredChannelNumber <> a_value Then
        This.MeasuredChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MeasuredChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Name for continuous mode. </summary>
''' <value>   [String]. </value>
Public Property Get ContinuousSenseFunctionName() As String
    ContinuousSenseFunctionName = This.ContinuousSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for continuous mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ContinuousSenseFunctionName(ByVal a_value As String)
    If This.ContinuousSenseFunctionName <> a_value Then
        This.ContinuousSenseFunctionName = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ContinuousSenseFunctionName = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Name for Immediate trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get ImmediateSenseFunctionName() As String
    ImmediateSenseFunctionName = This.ImmediateSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for Immediate trigger  mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ImmediateSenseFunctionName(ByVal a_value As String)
    If This.ImmediateSenseFunctionName <> a_value Then
        This.ImmediateSenseFunctionName = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ImmediateSenseFunctionName = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Name for External trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get ExternalSenseFunctionName() As String
    ExternalSenseFunctionName = This.ExternalSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for External trigger mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ExternalSenseFunctionName(ByVal a_value As String)
    If This.ExternalSenseFunctionName <> a_value Then
        This.ExternalSenseFunctionName = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ExternalSenseFunctionName = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Internal. Sets the Sense Function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let SenseFunctionName(ByVal a_value As String)
    If This.SenseFunctionName <> a_value Then
        This.SenseFunctionName = a_value
    End If
End Property

''' <summary>   Gets the  Sense Function Caption. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionCaption() As String
    SenseFunctionCaption = This.SenseFunctionCaption
End Property

''' <summary>   Sets the  Sense Function Caption. </summary>
''' <param name="a_value">   [String] the  Sense Function Caption. </value>
Public Property Let SenseFunctionCaption(ByVal a_value As String)
    If This.SenseFunctionCaption <> a_value Then
        This.SenseFunctionCaption = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SenseFunctionCaption = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Unit. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionUnit() As String
    SenseFunctionUnit = This.SenseFunctionUnit
End Property

''' <summary>   Sets the Sense Function Unit. </summary>
''' <param name="a_value">   [String] the Sense Function Unit value. </value>
Public Property Let SenseFunctionUnit(ByVal a_value As String)
    If This.SenseFunctionUnit <> a_value Then
        This.SenseFunctionUnit = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SenseFunctionUnit = a_value
    End If
End Property

''' <summary>   Gets the Measured Reading. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = This.MeasuredReading
End Property

''' <summary>   Sets the Measured Reading. </summary>
''' <param name="a_value">   [String] the Measured Reading value. </value>
Public Property Let MeasuredReading(ByVal a_value As String)
    If This.MeasuredReading <> a_value Then
        This.MeasuredReading = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MeasuredReading = a_value
    End If
End Property

''' <summary>   Gets the Measured Value Format. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredValueFormat() As String
    MeasuredValueFormat = This.MeasuredValueFormat
End Property

''' <summary>   Sets the Measured Value Format. </summary>
''' <param name="a_value">   [String] the Measured Value Format value. </value>
Public Property Let MeasuredValueFormat(ByVal a_value As String)
    If This.MeasuredValueFormat <> a_value Then
        This.MeasuredValueFormat = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MeasuredValueFormat = a_value
    End If
End Property

''' <summary>   Gets the Measured Value. </summary>
''' <value>   [Double]. </value>
Public Property Get MeasuredValue() As Double
    MeasuredValue = This.MeasuredValue
End Property

''' <summary>   Sets the Measured Value. </summary>
''' <param name="a_value">   [Double] the Measured Value. </value>
Public Property Let MeasuredValue(ByVal a_value As Double)
    If This.MeasuredValue <> a_value Then
        This.MeasuredValue = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MeasuredValue = a_value
    End If
End Property

''' <summary>   Gets the Sense Functions. </summary>
''' <value>   [Collection]. </value>
Public Property Get SenseFunctions() As Collection
    Set SenseFunctions = This.SenseFunctions
End Property

''' <summary>   Sets the Sense Functions. </summary>
''' <param name="a_value">   [Collection] the Sense Functions value. </value>
Public Property Let SenseFunctions(ByVal a_value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.SenseFunctions, a_value) Then
        Set This.SenseFunctions = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SenseFunctions = a_value
    End If
End Property

''' <summary>   Gets the Reading Offset. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadingOffset() As Double
    ReadingOffset = This.ReadingOffset
End Property

''' <summary>   Sets the Reading Offset. </summary>
''' <param name="a_value">   [Double] the Reading Offset value. </value>
Public Property Let ReadingOffset(ByVal a_value As Double)
    If This.ReadingOffset <> a_value Then
        This.ReadingOffset = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadingOffset = a_value
    End If
End Property

''' <summary>   Gets the Monitoring State. </summary>
''' <value>   [String]. </value>
Public Property Get MonitoringState() As String
    MonitoringState = This.MonitoringState
End Property

''' <summary>   Sets the Monitoring State. </summary>
''' <param name="a_value">   [String] the Monitoring State value. </value>
Public Property Let MonitoringState(ByVal a_value As String)
    If This.MonitoringState <> a_value Then
        This.MonitoringState = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MonitoringState = a_value
    End If
End Property

''' <summary>   Gets the External Trig Monitoring Enabled condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopd on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = This.ExternalTrigMonitoringEnabled
End Property

''' <summary>   Sets the External Trig Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trig Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If This.ExternalTrigMonitoringEnabled <> a_value Then
        This.ExternalTrigMonitoringEnabled = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ExternalTrigMonitoringEnabled = a_value
    End If
End Property

''' <summary>   Gets the Auto Increment Channel Number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementChannelNoEnabled() As Boolean
    AutoIncrementChannelNoEnabled = This.AutoIncrementChannelNoEnabled
End Property

''' <summary>   Sets the Auto Increment Channel Number Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment Channel Number Enabled value. </value>
Public Property Let AutoIncrementChannelNoEnabled(ByVal a_value As Boolean)
    If This.AutoIncrementChannelNoEnabled <> a_value Then
        This.AutoIncrementChannelNoEnabled = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.AutoIncrementChannelNoEnabled = This.AutoIncrementChannelNoEnabled
    End If
End Property

''' <summary>   Gets the Single Read Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get SingleReadEnabled() As Boolean
    SingleReadEnabled = This.SingleReadEnabled
End Property

''' <summary>   Sets the Single Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Single Read Enabled value. </value>
Public Property Let SingleReadEnabled(ByVal a_value As Boolean)
    If This.SingleReadEnabled <> a_value Then
        This.SingleReadEnabled = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SingleReadEnabled = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Pause Requested condition. </summary>
''' <remarks>   When true, a pause request is sent to the timer so that the timer
'''             event is not handled and the timer is paused on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get PauseRequested() As Boolean
    PauseRequested = This.PauseRequested
End Property

''' <summary>   Sets the Pause Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Pause Requested value. </value>
Friend Property Let PauseRequested(ByVal a_value As Boolean)
    If This.PauseRequested <> a_value Then
        This.PauseRequested = a_value
    End If
End Property

''' <summary>   Gets the Stop Requested condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopd on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get StopRequested() As Boolean
    StopRequested = This.StopRequested
End Property

''' <summary>   Sets the Stop Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Stop Requested value. </value>
Friend Property Let StopRequested(ByVal a_value As Boolean)
    If This.StopRequested <> a_value Then
        This.StopRequested = a_value
    End If
End Property

''' <summary>   Gets the Timer Interval in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Timer Interval in milliseconds value. </value>
Public Property Let TimerInterval(ByVal a_value As Integer)
    If This.TimerInterval <> a_value Then
        This.TimerInterval = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TimerInterval = a_value
    End If
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Double]. </value>
Public Property Get ElapsedTime() As Double
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Double] the Elapsed Time in milliseconds. </value>
Public Property Let ElapsedTime(ByVal a_value As Double)
    If This.ElapsedTime <> a_value Then
        This.ElapsedTime = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ElapsedTime = a_value
    End If
End Property

''' <summary>   Gets the Elapsed Time Format. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeFormat() As String
    ElapsedTimeFormat = This.ElapsedTimeFormat
End Property

''' <summary>   Sets the Elapsed Time Format. </summary>
''' <param name="a_value">   [String] the elapsed time format. </value>
Public Property Let ElapsedTimeFormat(ByVal a_value As String)
    If This.ElapsedTimeFormat <> a_value Then
        This.ElapsedTimeFormat = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ElapsedTimeFormat = a_value
    End If
End Property

''' <summary>   Gets the Elapsed Time Caption. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeCaption() As String
    ElapsedTimeCaption = This.ElapsedTimeCaption
End Property

''' <summary>   Sets the Elapsed Time Caption. </summary>
''' <param name="a_value">   [String] the elapsed time Caption. </value>
Public Property Let ElapsedTimeCaption(ByVal a_value As String)
    If This.ElapsedTimeCaption <> a_value Then
        This.ElapsedTimeCaption = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ElapsedTimeCaption = a_value
    End If
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = This.LastErrorMessage
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the last error message value. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    If This.LastErrorMessage <> a_value Then
        
        This.LastErrorMessage = a_value
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.LastErrorMessage = a_value
            
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.LastErrorMessage = a_value
            
    End If
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = This.LastMessage
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the last message value. </value>
Public Property Let LastMessage(ByVal a_value As String)
    If This.LastMessage <> a_value Then
        This.LastMessage = a_value
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.LastMessage = a_value
    
        If Not This.K2700Sheet Is Nothing Then _
            This.K2700Sheet.LastMessage = a_value
            
    End If
End Property

''' <summary>   Gets the Last Command Details. </summary>
''' <value>   [String]. </value>
Public Property Get LastCommandDetails() As String
    LastCommandDetails = This.LastCommandDetails
End Property

''' <summary>   Sets the Last Command Details. </summary>
''' <param name="a_value">   [String] the last Command Details value. </value>
Public Property Let LastCommandDetails(ByVal a_value As String)
    If This.LastCommandDetails <> a_value Then
        This.LastCommandDetails = a_value
    End If
End Property

''' <summary>   Gets the Last Command Outcome. </summary>
''' <value>   [Boolean]. </value>
Public Property Get LastCommandOutcome() As Boolean
    LastCommandOutcome = This.LastCommandOutcome
End Property

''' <summary>   Sets the Last Command Outcome. </summary>
''' <param name="a_value">   [Boolean] the last Command Outcome value. </value>
Public Property Let LastCommandOutcome(ByVal a_value As Boolean)
    If This.LastCommandOutcome <> a_value Then
        This.LastCommandOutcome = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE488 Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Reset Known State command can be executed. </value>
Public Property Get ResetKnownStateExecutable() As Boolean
    ResetKnownStateExecutable = This.ResetKnownStateExecutable
End Property

''' <summary>   Gets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Clear Execution State Command can be executed. </value>
Public Property Get ClearExecutionStateExecutable() As Boolean
    ClearExecutionStateExecutable = This.ClearExecutionStateExecutable
End Property

''' <summary>   Gets the condition indicating that the Query command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query command can be executed. </value>
Public Property Get QueryExecutable() As Boolean
    QueryExecutable = This.QueryExecutable
End Property

''' <summary>   Gets the condition indicating that the Read command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read command can be executed. </value>
Public Property Get ReadExecutable() As Boolean
    ReadExecutable = This.ReadExecutable
End Property

''' <summary>   Gets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status Byte command can be executed. </value>
Public Property Get QueryStatusByteExecutable() As Boolean
    QueryStatusByteExecutable = This.QueryStatusByteExecutable
End Property

''' <summary>   Gets the condition indicating that the Read Status After Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status After Write command can be executed. </value>
Public Property Get ReadStatusAfterWriteExecutable() As Boolean
    ReadStatusAfterWriteExecutable = This.ReadStatusAfterWriteExecutable
End Property

''' <summary>   Gets the condition indicating that the Read Standard Event Status command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Standard Event Status command can be executed. </value>
Public Property Get ReadStandardEventStatusExecutable() As Boolean
    ReadStandardEventStatusExecutable = This.ReadStandardEventStatusExecutable
End Property

''' <summary>   Gets the condition indicating that the Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Write command can be executed. </value>
Public Property Get WriteExecutable() As Boolean
    WriteExecutable = This.WriteExecutable
End Property

''' <summary>   Gets the condition indicating that Message To Send command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Message to Send command can be executed. </value>
Public Property Get MessageToSendExecutable() As Boolean
    MessageToSendExecutable = This.MessageToSendExecutable
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Executables
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Selective Device Clear command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Selective Device Clear command can be executed. </value>
Public Property Get SelectiveDeviceClearExecutable() As Boolean
    SelectiveDeviceClearExecutable = This.SelectiveDeviceClearExecutable
End Property

''' <summary>   Gets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Go To Local command can be executed. </value>
Public Property Get GoToLocalExecutable() As Boolean
    GoToLocalExecutable = This.GoToLocalExecutable
End Property


''' <summary>   Gets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Local Lockout command can be executed. </value>
Public Property Get LocalLockoutExecutable() As Boolean
    LocalLockoutExecutable = This.LocalLockoutExecutable
End Property

''' <summary>   Gets the condition indicating that the Auto Asset TALK get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK get command can be executed. </value>
Public Property Get AutoAssertTalkGetExecutable() As Boolean
    AutoAssertTalkGetExecutable = This.AutoAssertTalkGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Auto Asset TALK Setcommand
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK Setcommand can be executed. </value>
Public Property Get AutoAssertTalkSetExecutable() As Boolean
    AutoAssertTalkSetExecutable = This.AutoAssertTalkSetExecutable
End Property

''' <summary>   Gets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get ReadStatusByteExecutable() As Boolean
    ReadStatusByteExecutable = This.ReadStatusByteExecutable
End Property

''' <summary>   Gets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Service Requested command can be executed. </value>
Public Property Get ServiceRequestedExecutable() As Boolean
    ServiceRequestedExecutable = This.ServiceRequestedExecutable
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Get Command can be executed. </value>
Public Property Get GpibAddressGetExecutable() As Boolean
    GpibAddressGetExecutable = This.GpibAddressGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Set Command can be executed. </value>
Public Property Get GpibAddressSetExecutable() As Boolean
    GpibAddressSetExecutable = This.GpibAddressSetExecutable
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Get command can be executed. </value>
Public Property Get GpibLanReadTimeoutGetExecutable() As Boolean
    GpibLanReadTimeoutGetExecutable = This.GpibLanReadTimeoutGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Set command can be executed. </value>
Public Property Get GpibLanReadTimeoutSetExecutable() As Boolean
    GpibLanReadTimeoutSetExecutable = This.GpibLanReadTimeoutSetExecutable
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Executables
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <value>   [Boolean] True if the Clear Readings command can be executed. </value>
Public Property Get ClearReadingsExecutable() As Boolean
    If This.K2700Sheet.ClearReadingsButton Is Nothing Then
        ClearReadingsExecutable = This.ClearReadingsExecutable
    Else
        ClearReadingsExecutable = This.K2700Sheet.ClearReadingsButton.Enabled
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If Not This.K2700Sheet.ClearReadingsButton Is Nothing Then
        This.K2700Sheet.ClearReadingsButton.Enabled = a_value
    End If
    This.ClearReadingsExecutable = a_value
End Property

''' <summary>   Gets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Inputs command can be executed. </value>
Public Property Get QueryInputsExecutable() As Boolean
    QueryInputsExecutable = This.QueryInputsExecutable
End Property

''' <summary>   Gets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Measure command can be executed. </value>
Public Property Get MeasureExecutable() As Boolean
    MeasureExecutable = This.MeasureExecutable
End Property

''' <summary>   Gets the condition indicating that the Immediate Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Immediate Trigger Option command  can be executed. </value>
Public Property Get ImmediateTriggerOptionExecutable() As Boolean
    ImmediateTriggerOptionExecutable = This.ImmediateTriggerOptionExecutable
End Property

''' <summary>   Gets the condition indicating that the External Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the External Trigger Option command  can be executed. </value>
Public Property Get ExternalTriggerOptionExecutable() As Boolean
    ExternalTriggerOptionExecutable = This.ExternalTriggerOptionExecutable
End Property

''' <summary>   Gets the condition indicating that the function name command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the function name command can be executed. </value>
Public Property Get FunctionNameExecutable() As Boolean
    FunctionNameExecutable = This.FunctionNameExecutable
End Property

''' <summary>   Gets the condition indicating that the Start Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Start Monitoring External Trigger command can be executed. </value>
Public Property Get StartMonitoringExecutable() As Boolean
    StartMonitoringExecutable = This.StartMonitoringExecutable
End Property

''' <summary>   Gets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </value>
Public Property Get StopMonitoringExecutable() As Boolean
    StopMonitoringExecutable = This.StopMonitoringExecutable
End Property

''' <summary>   Gets the condition indicating that the Toggle Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Monitoring External Trigger command can be executed. </value>
Public Property Get ToggleMonitoringExecutable() As Boolean
    ToggleMonitoringExecutable = This.StartMonitoringExecutable Or This.StopMonitoringExecutable
End Property

''' <summary>   Handles the <see cref=cc_isr_Tcp_Scpi.K2700ViewModel"/>.<see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_viewModel_ChannelMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.MeasuredReading = a_eventArgs.Reading
    Me.MeasuredChannelNumber = a_eventArgs.ChannelNumber
        
    Dim p_value As Double
    Dim p_details As String
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        p_value = -1#
    Else
        If Not cc_isr_Core.StringExtensions.TryParseDouble(a_eventArgs.Reading, p_value, p_details) Then
            p_value = -1#
        End If
    End If
    Me.MeasuredValue = p_value
    
    Me.SetReading This.MeasuredChannelNumber, This.MeasuredValue
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Friend Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "HandleViewModelPropertyChanged"

    Select Case a_propertyName

        Case "Host"
            Me.Host = This.ViewModel.Host
        Case "Port"
            Me.Port = This.ViewModel.Port
        Case "SocketAddress"
            Me.SocketAddress = This.ViewModel.SocketAddress
        Case "SocketId"
            Me.SocketId = This.ViewModel.SocketId
        Case "SessionTimeout"
            Me.SessionTimeout = This.ViewModel.SessionTimeout
        Case "PauseRequested"
            Me.PauseRequested = This.ViewModel.PauseRequested
        Case "StopRequested"
            Me.StopRequested = This.ViewModel.StopRequested
        Case "TimerInterval"
            Me.TimerInterval = This.ViewModel.TimerInterval
        Case "Connected"
            Me.Connected = This.ViewModel.Connected
        Case "MessageToSend"
            Me.MessageToSend = This.ViewModel.MessageToSend
        Case "SentMessage"
            Me.SentMessage = This.ViewModel.SentMessage
        Case "ReceivedMessage"
            Me.ReceivedMessage = This.ViewModel.ReceivedMessage
        Case "ReceivedMessageLength"
            Me.ReceivedMessageLength = This.ViewModel.ReceivedMessageLength
        Case "ElapsedTime"
            Me.ElapsedTime = This.ViewModel.ElapsedTime
        Case "ElapsedTimeFormat"
            Me.ElapsedTimeFormat = This.ViewModel.ElapsedTimeFormat
        Case "ElapsedTimeCaption"
            Me.ElapsedTimeCaption = This.ViewModel.ElapsedTimeCaption
        Case "LastErrorMessage"
            Me.LastErrorMessage = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            Me.LastMessage = This.ViewModel.LastMessage
        Case "SerialPollByte"
            Me.SerialPollByte = This.ViewModel.SerialPollByte
        Case "StatusByte"
            Me.StatusByte = This.ViewModel.StatusByte
        Case "StandardByte"
            Me.StandardByte = This.ViewModel.StandardByte
        Case "ServiceRequested"
            Me.ServiceRequested = This.ViewModel.ServiceRequested
        Case "GpibAddress"
            Me.GpibAddress = This.ViewModel.GpibAddress
        Case "AutoAssertTalk"
            Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
        Case "GpibLanReadTimeout"
            Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
            
        Case "TopCard"
            Me.TopCard = This.ViewModel.TopCard
        Case "BottomCard"
            Me.BottomCard = This.ViewModel.BottomCard
        Case "TopCardFunctionScanList"
            Me.TopCardFunctionScanList = This.ViewModel.TopCardFunctionScanList
        Case "BottomCardFunctionScanList"
            Me.BottomCardFunctionScanList = This.ViewModel.BottomCardFunctionScanList
        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "FrontInputsValue"
        Case "FrontInputsCaption"
            Me.FrontInputsCaption = This.ViewModel.FrontInputsCaption
        Case "SelectedChannelNumber"
            Me.SelectedChannelNumber = This.ViewModel.SelectedChannelNumber
        Case "ChannelCount"
            Me.ChannelCount = This.ViewModel.ChannelCount
        Case "TargetChannelNumber"
            Me.TargetChannelNumber = This.ViewModel.TargetChannelNumber
        Case "MeasuredChannelNumber"
            Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
        Case "SenseFunctionName"
            Me.SenseFunctionName = This.ViewModel.SenseFunctionName
        Case "SenseFunctionCaption"
            Me.SenseFunctionCaption = This.ViewModel.SenseFunctionCaption
        Case "SenseFunctionUnit"
            Me.SenseFunctionUnit = This.ViewModel.SenseFunctionUnit
        Case "MeasuredReading"
            Me.MeasuredReading = This.ViewModel.MeasuredReading
        Case "MeasuredValueFormat"
            Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
        
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "ExternalTrigMonitoringEnabled"
            Me.ExternalTrigMonitoringEnabled = This.ViewModel.ExternalTrigMonitoringEnabled
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        Case "AutoIncrementChannelNoEnabled"
            Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
        Case "SenseFunctions"
            Me.SenseFunctions = This.ViewModel.SenseFunctions
        Case "ScpiCommands"
            Me.ScpiCommands = This.ViewModel.ScpiCommands
        Case "Measuring"
            Me.Measuring = This.ViewModel.Measuring
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "ReadStatusAfterWriteEnabled"
            Me.ReadStatusAfterWriteEnabled = This.ViewModel.ReadStatusAfterWriteEnabled
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        Case "DesignMode"
            Me.DesignMode = This.ViewModel.DesignMode
        
        Case "OpenConnectionExecutable"
            This.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
        Case "CloseConnectionExecutable"
            This.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
        Case "ToggleConnectionExecutable"
        Case "ResetKnownStateExecutable"
            This.ResetKnownStateExecutable = This.ViewModel.ResetKnownStateExecutable
        Case "ClearExecutionStateExecutable"
            This.ClearExecutionStateExecutable = This.ViewModel.ClearExecutionStateExecutable
        Case "QueryExecutable"
            This.QueryExecutable = This.ViewModel.QueryExecutable
        Case "ReadExecutable"
            This.ReadExecutable = This.ViewModel.ReadExecutable
        Case "QueryStatusByteExecutable"
            This.QueryStatusByteExecutable = This.ViewModel.QueryStatusByteExecutable
        Case "ReadStandardEventStatusExecutable"
            This.ReadStandardEventStatusExecutable = This.ViewModel.ReadStandardEventStatusExecutable
        Case "ReadStatusAfterWriteExecutable"
            This.ReadStatusAfterWriteExecutable = This.ViewModel.ReadStatusAfterWriteExecutable
        Case "WriteExecutable"
            This.WriteExecutable = This.ViewModel.WriteExecutable
        Case "QueryInputsExecutable"
            This.QueryInputsExecutable = This.ViewModel.QueryInputsExecutable
        Case "ClearReadingsExecutable"
            This.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
        Case "MeasureExecutable"
            This.MeasureExecutable = This.ViewModel.MeasureExecutable
        Case "StartMonitoringExecutable"
            This.StartMonitoringExecutable = This.ViewModel.StartMonitoringExecutable
        Case "StopMonitoringExecutable"
            This.StopMonitoringExecutable = This.ViewModel.StopMonitoringExecutable
        Case "ImmediateTriggerOptionExecutable"
            This.ImmediateTriggerOptionExecutable = This.ViewModel.ImmediateTriggerOptionExecutable
        Case "ExternalTriggerOptionExecutable"
            This.ExternalTriggerOptionExecutable = This.ViewModel.ExternalTriggerOptionExecutable
        Case "SelectiveDeviceClearExecutable"
            This.SelectiveDeviceClearExecutable = This.ViewModel.SelectiveDeviceClearExecutable
        Case "GoToLocalExecutable"
            This.GoToLocalExecutable = This.ViewModel.GoToLocalExecutable
        Case "LocalLockoutExecutable"
            This.LocalLockoutExecutable = This.ViewModel.LocalLockoutExecutable
        Case "AutoAssertTalkSetExecutable"
            This.AutoAssertTalkSetExecutable = This.ViewModel.AutoAssertTalkSetExecutable
        Case "AutoAssertTalkGetExecutable"
            This.AutoAssertTalkGetExecutable = This.ViewModel.AutoAssertTalkGetExecutable
        Case "ReadStatusByteExecutable"
            This.ReadStatusByteExecutable = This.ViewModel.ReadStatusByteExecutable
        Case "ServiceRequestedExecutable"
            This.ServiceRequestedExecutable = This.ViewModel.ServiceRequestedExecutable
        Case "GpibAddressGetExecutable"
            This.GpibAddressGetExecutable = This.ViewModel.GpibAddressGetExecutable
        Case "GpibAddressSetExecutable"
            This.GpibAddressSetExecutable = This.ViewModel.GpibAddressSetExecutable
        Case "GpibLanReadTimeoutGetExecutable"
            This.GpibLanReadTimeoutGetExecutable = This.ViewModel.GpibLanReadTimeoutGetExecutable
        Case "GpibLanReadTimeoutSetExecutable"
            This.GpibLanReadTimeoutSetExecutable = This.ViewModel.GpibLanReadTimeoutSetExecutable
        Case "MessageToSendExecutable"
            This.MessageToSendExecutable = This.ViewModel.MessageToSendExecutable
        Case "FunctionNameExecutable"
            This.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable

    End Select
    
End Sub

''' <summary>   Synchronizes the view mode with the Observer. </summary>
''' <remarks>   This is required becuase the view (observer) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a resule, the observer retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' set the view model settings that come from the ViewModel user interface
    
    ' set user interface settings
    Me.GpibLanControllerPort = This.K2700Sheet.GpibLanControllerPort
    Me.GpibAddress = This.K2700Sheet.PrimaryGpibAddress
    Me.MaximumChannelNumber = This.K2700Sheet.MaximumChannelNumber
    Me.SocketAddress = This.K2700Sheet.SocketAddress
    Me.SessionTimeout = This.K2700Sheet.SessionTimeout
    Me.TimerInterval = This.K2700Sheet.TimerInterval
    
    This.ViewModel.SocketAddress = This.SocketAddress
    This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort
    This.ViewModel.SessionTimeout = Me.SessionTimeout
    This.ViewModel.TimerInterval = Me.TimerInterval
    This.ViewModel.MaximumChannelNumber = Me.MaximumChannelNumber

    Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
    Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
    Me.BottomCard = This.ViewModel.BottomCard
    Me.BottomCardFunctionScanList = This.ViewModel.BottomCardFunctionScanList
    Me.ChannelCount = This.ViewModel.ChannelCount
    Me.Connected = This.ViewModel.Connected
    Me.ElapsedTime = This.ViewModel.ElapsedTime
    Me.ElapsedTimeFormat = This.ViewModel.ElapsedTimeFormat
    Me.ElapsedTimeCaption = This.ViewModel.ElapsedTimeCaption
    Me.ExternalTrigMonitoringEnabled = This.ViewModel.ExternalTrigMonitoringEnabled
    Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
    Me.FrontInputsCaption = This.ViewModel.FrontInputsCaption
    Me.GpibAddress = This.ViewModel.GpibAddress
    Me.Host = This.ViewModel.Host
    Me.LastErrorMessage = This.ViewModel.LastErrorMessage
    Me.LastMessage = This.ViewModel.LastMessage
    Me.MessageToSend = This.ViewModel.MessageToSend
    Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.MonitoringState = This.ViewModel.MonitoringState
    Me.PauseRequested = This.ViewModel.PauseRequested
    Me.Port = This.ViewModel.Port
    Me.ReadingOffset = This.ViewModel.ReadingOffset
    Me.ReadStatusAfterWriteEnabled = This.ViewModel.ReadStatusAfterWriteEnabled
    Me.ReceivedMessage = This.ViewModel.ReceivedMessage
    Me.ReceivedMessageLength = This.ViewModel.ReceivedMessageLength
    Me.ScpiCommands = This.ViewModel.ScpiCommands
    Me.SelectedChannelNumber = This.ViewModel.SelectedChannelNumber
    Me.SenseFunctionName = This.ViewModel.SenseFunctionName
    Me.SenseFunctionCaption = This.ViewModel.SenseFunctionCaption
    Me.SenseFunctions = This.ViewModel.SenseFunctions
    Me.SenseFunctionUnit = This.ViewModel.SenseFunctionUnit
    Me.SentMessage = This.ViewModel.SentMessage
    Me.SerialPollByte = This.ViewModel.SerialPollByte
    Me.ServiceRequested = This.ViewModel.ServiceRequested
    Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
    Me.SocketAddress = This.ViewModel.SocketAddress
    Me.SocketId = This.ViewModel.SocketId
    Me.StandardByte = This.ViewModel.StandardByte
    Me.StatusByte = This.ViewModel.StatusByte
    Me.StopRequested = This.ViewModel.StopRequested
    Me.TargetChannelNumber = This.ViewModel.TargetChannelNumber
    Me.TopCard = This.ViewModel.TopCard
    Me.TopCardFunctionScanList = This.ViewModel.TopCardFunctionScanList
    
    This.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
    This.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
    This.ResetKnownStateExecutable = This.ViewModel.ResetKnownStateExecutable
    This.ClearExecutionStateExecutable = This.ViewModel.ClearExecutionStateExecutable
    This.QueryExecutable = This.ViewModel.QueryExecutable
    This.ReadExecutable = This.ViewModel.ReadExecutable
    This.QueryStatusByteExecutable = This.ViewModel.QueryStatusByteExecutable
    This.ReadStandardEventStatusExecutable = This.ViewModel.ReadStandardEventStatusExecutable
    This.ReadStatusAfterWriteExecutable = This.ViewModel.ReadStatusAfterWriteExecutable
    This.WriteExecutable = This.ViewModel.WriteExecutable
    This.QueryInputsExecutable = This.ViewModel.QueryInputsExecutable
    This.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
    This.MeasureExecutable = This.ViewModel.MeasureExecutable
    This.StartMonitoringExecutable = This.ViewModel.StartMonitoringExecutable
    This.StopMonitoringExecutable = This.ViewModel.StopMonitoringExecutable
    This.ImmediateTriggerOptionExecutable = This.ViewModel.ImmediateTriggerOptionExecutable
    This.ExternalTriggerOptionExecutable = This.ViewModel.ExternalTriggerOptionExecutable
    This.SelectiveDeviceClearExecutable = This.ViewModel.SelectiveDeviceClearExecutable
    This.GoToLocalExecutable = This.ViewModel.GoToLocalExecutable
    This.LocalLockoutExecutable = This.ViewModel.LocalLockoutExecutable
    This.AutoAssertTalkSetExecutable = This.ViewModel.AutoAssertTalkSetExecutable
    This.AutoAssertTalkGetExecutable = This.ViewModel.AutoAssertTalkGetExecutable
    This.ReadStatusByteExecutable = This.ViewModel.ReadStatusByteExecutable
    This.ServiceRequestedExecutable = This.ViewModel.ServiceRequestedExecutable
    This.GpibAddressGetExecutable = This.ViewModel.GpibAddressGetExecutable
    This.GpibAddressSetExecutable = This.ViewModel.GpibAddressSetExecutable
    This.GpibLanReadTimeoutGetExecutable = This.ViewModel.GpibLanReadTimeoutGetExecutable
    This.GpibLanReadTimeoutSetExecutable = This.ViewModel.GpibLanReadTimeoutSetExecutable
    This.MessageToSendExecutable = This.ViewModel.MessageToSendExecutable
    This.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_viewModel_PropertyChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not a_eventArgs Is Nothing Then
    
        Dim p_propertyName As String
        p_propertyName = a_eventArgs.PropertyName
        HandleViewModelPropertyChanged p_propertyName
    
    End If

    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' append a description
    Err.Description = Err.Description & ". Property name: '" & p_propertyName & "'."
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

