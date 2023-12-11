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

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_

    ' Observer properties
    Initialized As Boolean
    FocusedControlName As String
        
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
    MaximumDutNumber As Integer
    

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
    DutCount As Integer
    SelectedDutNumber As Integer
    TargetDutNumber As Integer
    MeasuredDutNumber As Integer
    RearInputsSenseFunctionName As String
    FrontInputsSenseFunctionName As String
    SenseFunctionName As String
    SenseFunctionCaption As String
    SenseFunctionUnit As String
    MeasuredReading As String
    MeasuredValueFormat As String
    MeasuredValue As Double
    SenseFunctions As Collection
    
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
    AutoIncrementDutNumberEnabled As Boolean
    SingleReadEnabled As Boolean

    ' Winsock connection executables
    OpenConnectionExecutable As Boolean
    CloseConnectionExecutable As Boolean
    ToggleConnectionExecutable As Boolean
    
    ' IEEE488 executables
    ClearExecutionStateExecutable As Boolean
    QueryExecutable As Boolean
    ReadExecutable As Boolean
    QueryStatusByteExecutable As Boolean
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Toggle connection button click event. </summary>
Public Sub OnConnect()
    If This.ToggleConnectionExecutable Then
        If This.OpenConnectionExecutable Then
            This.ViewModel.OpenConnectionCommand This.SocketAddress, This.SessionTimeout
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE488 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
''' <param name="a_value">   [String] the Message to Send. </value>
Public Sub OnQuery(ByVal a_value As String)
    If This.QueryExecutable Then
        This.ViewModel.QueryCommand a_value
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
''' <param name="a_value">        [String] the Message to Send. </value>
''' <param name="a_readStatus">   [String] reads server request and standard event
'''                                statuses depending on this being a query message. </value>
Public Sub OnWrite(ByVal a_value As String, ByVal a_readStatus As Boolean)
    If This.WriteExecutable Then
        This.ViewModel.WriteCommand a_value, a_readStatus
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnWrite", _
            " Command not enabled."
    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
''' <param name="a_value">   [Integer] the Primary Gpib Address value. </value>
Public Sub OnGpibAddressSet(ByVal a_value As Integer)
    If This.GpibAddressSetExecutable Then
        This.ViewModel.GpibAddressSetCommand a_value
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
''' <param name="a_value">   [Boolean] True to Auto Assert Talk upon write. </value>
Public Sub OnAutoAssertTalkSet(ByVal a_value As Boolean)
    If This.AutoAssertTalkSetExecutable Then
        This.ViewModel.AutoAssertTalkSetCommand a_value
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition View
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Public Sub OnClearReadings()
    If This.ClearReadingsExecutable Then
        This.ViewModel.ClearReadingsCommand This.ReadingOffset, 1, This.MaximumDutNumber
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & "OnClearReadings", _
            " Command not enabled."
    End If
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the external trigger option button change and click events. </summary>
Public Sub OnExternalTriggerOptionChange(ByVal a_value As Boolean)
    If This.ExternalTriggerOptionExecutable Then
        If a_value Then
        
            Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
            Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
            p_mode.BeepEnabled = True
            p_mode.AutoIncrement = False
            p_mode.FrontInputs = DataView.ExternalFrontInputsRequired
            p_mode.MaximumDutCount = DataView.MaximumDutNumber
            p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
            
            p_mode.DutNumber = This.SelectedDutNumber
            
            p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
            p_mode.ReadingOffset = UserView.ReadingOffset
            p_mode.SenseFunction = IIf(p_mode.FrontInputs, DataView.FrontInputsSenseFunctionName, _
                DataView.RearInputsSenseFunctionName)
            p_mode.SingleRead = True
            p_mode.TimerInterval = DataView.TimerInterval
        
            Me.LastCommandOutcome = This.ViewModel.ConfigureMeasureCommand(p_mode, Me.LastCommandDetails)
        End If
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".OnExternalTriggerOptionChange", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
''' <remarks>   The check box change event also causes a click event.
'''             This method issues the click event on actual click.
'''             Note: Make sure to not implement the click event,
Public Sub OnImmediateTriggerOptionChange(ByVal a_value As Boolean)

    If This.ImmediateTriggerOptionExecutable Then
        If a_value Then
            
            Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
            Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
            p_mode.AutoIncrement = False
            p_mode.BeepEnabled = True
            p_mode.FrontInputs = DataView.ImmediateFrontInputsRequired
            p_mode.MaximumDutCount = DataView.MaximumDutNumber
            p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
            
            p_mode.DutNumber = This.SelectedDutNumber
            
            p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
            p_mode.ReadingOffset = UserView.ReadingOffset
            p_mode.SenseFunction = IIf(p_mode.FrontInputs, DataView.FrontInputsSenseFunctionName, _
                DataView.RearInputsSenseFunctionName)
            p_mode.SingleRead = True
            p_mode.TimerInterval = DataView.TimerInterval
            
            Me.LastCommandOutcome = This.ViewModel.ConfigureMeasureCommand(p_mode, Me.LastCommandDetails)
        End If
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the INET Address dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Sets the Host. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let Host(ByVal a_value As String)
    This.Host = a_value
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Long
    Port = This.Port
End Property

''' <summary>   Sets the Port. </summary>
''' <param name="a_value">   [Long] the Port. </value>
Public Property Let Port(ByVal a_value As Long)
    This.Port = a_value
End Property

''' <summary>   Gets the Connected state. </summary>
''' <value>   [Integer]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property

''' <summary>   Sets the Connected state. </summary>
''' <param name="a_value">   [Boolean] the Connected state. </value>
Public Property Let Connected(ByVal a_value As Boolean)
    This.Connected = a_value
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_value As String)
    This.SocketAddress = a_value
End Property

''' <summary>   Gets the socket id. </summary>
''' <value>   [Integer]. </value>
Public Property Get SocketId() As Integer
    SocketId = This.SocketId
End Property

''' <summary>   Sets the socket id. </summary>
''' <param name="a_value">   [Integer] the socket id. </value>
Public Property Let SocketId(ByVal a_value As Integer)
    This.SocketId = a_value
End Property

''' <summary>   Gets the socket and session timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket and session timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_value As Integer)
    This.SessionTimeout = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Message to Sent. </summary>
''' <value>   [String]. </value>
Public Property Get MessageToSend() As String
    MessageToSend = This.MessageToSend
End Property

''' <summary>   Sets the Message to Sent. </summary>
''' <param name="a_value">   [String] the Message to Sent. </value>
Public Property Let MessageToSend(ByVal a_value As String)
    This.MessageToSend = a_value
End Property

''' <summary>   Gets the Sent Message. </summary>
''' <value>   [String]. </value>
Public Property Get SentMessage() As String
    SentMessage = This.SentMessage
End Property

''' <summary>   Sets the Sent Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SentMessage(ByVal a_value As String)
    This.SentMessage = a_value
End Property

''' <summary>   Gets the Received Message. </summary>
''' <value>   [String]. </value>
Public Property Get ReceivedMessage() As String
    ReceivedMessage = This.ReceivedMessage
End Property

''' <summary>   Sets the Received Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let ReceivedMessage(ByVal a_value As String)
    This.ReceivedMessage = a_value
End Property

''' <summary>   Gets the Received Message Length. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceivedMessageLength() As Integer
    ReceivedMessageLength = This.ReceivedMessageLength
End Property

''' <summary>   Sets the Received Message Length. </summary>
''' <param name="a_value">   [Integer] the Message Length value. </value>
Public Property Let ReceivedMessageLength(ByVal a_value As Integer)
    This.ReceivedMessageLength = a_value
End Property

''' <summary>   Gets the Status Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StatusByte() As Integer
    StatusByte = This.StatusByte
End Property

''' <summary>   Sets the Status Byte. </summary>
''' <param name="a_value">   [Integer] the Status Byte value. </value>
Public Property Let StatusByte(ByVal a_value As Integer)
    This.StatusByte = a_value
End Property

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    StandardByte = This.StandardByte
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Public Property Let StandardByte(ByVal a_value As Integer)
    This.StandardByte = a_value
End Property

''' <summary>   Gets the Auto Status Read Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadStatusAfterWriteEnabled() As Boolean
    ReadStatusAfterWriteEnabled = This.ReadStatusAfterWriteEnabled
End Property

''' <summary>   Sets the Auto Status Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Status Read Enabled value. </value>
Public Property Let ReadStatusAfterWriteEnabled(ByVal a_value As Boolean)
    This.ReadStatusAfterWriteEnabled = a_value
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
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Serial Poll Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get SerialPollByte() As Integer
    SerialPollByte = This.SerialPollByte
End Property

''' <summary>   Sets the Serial Poll Byte. </summary>
''' <param name="a_value">   [Integer] the Serial Poll Byte value. </value>
Public Property Let SerialPollByte(ByVal a_value As Integer)
    This.SerialPollByte = a_value
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_value As Integer)
    This.ServiceRequested = a_value
End Property

''' <summary>   Gets the Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibAddress() As Integer
    GpibAddress = This.GpibAddress
End Property

''' <summary>   Sets the Gpib Address. </summary>
''' <param name="a_value">   [Integer] the Gpib Address value. </value>
Public Property Let GpibAddress(ByVal a_value As Integer)
    This.GpibAddress = a_value
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let AutoAssertTalk(ByVal a_value As Boolean)
    This.AutoAssertTalk = a_value
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
    This.GpibLanReadTimeout = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Top Card name. </summary>
''' <value>   [String]. </value>
Public Property Get TopCard() As String
    TopCard = This.TopCard
End Property

''' <summary>   Sets the Top Card name. </summary>
''' <param name="a_value">   [String] the Top Card name. </value>
Public Property Let TopCard(ByVal a_value As String)
    This.TopCard = a_value
End Property

''' <summary>   Gets the Top Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get TopCardFunctionScanList() As String
    TopCardFunctionScanList = This.TopCardFunctionScanList
End Property

''' <summary>   Sets the Top Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Top Card function scan list. </value>
Public Property Let TopCardFunctionScanList(ByVal a_value As String)
    This.TopCardFunctionScanList = a_value
End Property

''' <summary>   Gets the Bottom Card name. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCard() As String
    BottomCard = This.BottomCard
End Property

''' <summary>   Sets the Bottom Card name. </summary>
''' <param name="a_value">   [String] the Bottom Card name. </value>
Public Property Let BottomCard(ByVal a_value As String)
    This.BottomCard = a_value
End Property

''' <summary>   Gets the Bottom Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCardFunctionScanList() As String
    BottomCardFunctionScanList = This.BottomCardFunctionScanList
End Property

''' <summary>   Sets the Bottom Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Bottom Card function scan list. </value>
Public Property Let BottomCardFunctionScanList(ByVal a_value As String)
    This.BottomCardFunctionScanList = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the maximum DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutNumber() As Integer
    MaximumDutNumber = This.MaximumDutNumber
End Property

''' <summary>   Sets the maximum DUT number. </summary>
''' <param name="a_value">   [Integer] the maximum DUT number value. </value>
Public Property Let MaximumDutNumber(ByVal a_value As Integer)
    This.MaximumDutNumber = a_value
End Property

''' <summary>   Gets the measuring state. True when measurements are collected. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Measuring() As Boolean
    Measuring = This.Measuring
End Property

''' <summary>   Sets the measuring state. True when measurements are collected. </summary>
''' <param name="a_value">   [Boolean] the measuring state; true when measuring is active. </value>
Public Property Let Measuring(ByVal a_value As Boolean)
    This.Measuring = a_value
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
    This.MeasurementMode = a_value
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
    This.FrontInputsRequired = a_value
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
    This.ReadingBeepEnabled = a_value
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
    This.FrontInputsValue = a_value
End Property

''' <summary>   Gets the Front Inputs Caption. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsCaption() As String
    FrontInputsCaption = This.FrontInputsCaption
End Property

''' <summary>   Sets the Front Inputs Caption. </summary>
''' <param name="a_value">   [String] the Front Inputs Caption. </value>
Public Property Let FrontInputsCaption(ByVal a_value As String)
    This.FrontInputsCaption = a_value
End Property

''' <summary>   Gets the DUT count. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutCount() As Integer
    DutCount = This.DutCount
End Property

''' <summary>   Sets the DUT count. </summary>
''' <param name="a_value">   [Integer] the DUT count value. </value>
Public Property Let DutCount(ByVal a_value As Integer)
    This.DutCount = a_value
End Property

''' <summary>   Gets the Selected DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedDutNumber() As Integer
    SelectedDutNumber = This.SelectedDutNumber
End Property

''' <summary>   Sets the Selected DUT number. </summary>
''' <param name="a_value">   [Integer] the Selected DUT number value. </value>
Public Property Let SelectedDutNumber(ByVal a_value As Integer)
    This.SelectedDutNumber = a_value
End Property

''' <summary>   Gets the Target DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get TargetDutNumber() As Integer
    TargetDutNumber = This.TargetDutNumber
End Property

''' <summary>   Sets the Target DUT number. </summary>
''' <param name="a_value">   [Integer] the Target DUT number value. </value>
Public Property Let TargetDutNumber(ByVal a_value As Integer)
    This.TargetDutNumber = a_value
End Property

''' <summary>   Gets the Measured DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredDutNumber() As Integer
    MeasuredDutNumber = This.MeasuredDutNumber
End Property

''' <summary>   Sets the Measured DUT number. </summary>
''' <param name="a_value">   [Integer] the Measured DUT number value. </value>
Public Property Let MeasuredDutNumber(ByVal a_value As Integer)
    This.MeasuredDutNumber = a_value
End Property

''' <summary>   Gets the Sense Function Name for Immediate trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get RearInputsSenseFunctionName() As String
    RearInputsSenseFunctionName = This.RearInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for rear inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let RearInputsSenseFunctionName(ByVal a_value As String)
    This.RearInputsSenseFunctionName = a_value
End Property

''' <summary>   Gets the Sense Function Name for front inputs. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsSenseFunctionName() As String
    FrontInputsSenseFunctionName = This.FrontInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for front inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let FrontInputsSenseFunctionName(ByVal a_value As String)
    This.FrontInputsSenseFunctionName = a_value
End Property

''' <summary>   Gets the Sense Function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Internal. Sets the Sense Function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let SenseFunctionName(ByVal a_value As String)
    This.SenseFunctionName = a_value
End Property

''' <summary>   Gets the  Sense Function Caption. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionCaption() As String
    SenseFunctionCaption = This.SenseFunctionCaption
End Property

''' <summary>   Sets the  Sense Function Caption. </summary>
''' <param name="a_value">   [String] the  Sense Function Caption. </value>
Public Property Let SenseFunctionCaption(ByVal a_value As String)
    This.SenseFunctionCaption = a_value
End Property

''' <summary>   Gets the Sense Function Unit. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionUnit() As String
    SenseFunctionUnit = This.SenseFunctionUnit
End Property

''' <summary>   Sets the Sense Function Unit. </summary>
''' <param name="a_value">   [String] the Sense Function Unit value. </value>
Public Property Let SenseFunctionUnit(ByVal a_value As String)
    This.SenseFunctionUnit = a_value
End Property

''' <summary>   Gets the Measured Reading. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = This.MeasuredReading
End Property

''' <summary>   Sets the Measured Reading. </summary>
''' <param name="a_value">   [String] the Measured Reading value. </value>
Public Property Let MeasuredReading(ByVal a_value As String)
    This.MeasuredReading = a_value
End Property

''' <summary>   Gets the Measured Value Format. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredValueFormat() As String
    MeasuredValueFormat = This.MeasuredValueFormat
End Property

''' <summary>   Sets the Measured Value Format. </summary>
''' <param name="a_value">   [String] the Measured Value Format value. </value>
Public Property Let MeasuredValueFormat(ByVal a_value As String)
    This.MeasuredValueFormat = a_value
End Property

''' <summary>   Gets the Measured Value. </summary>
''' <value>   [Double]. </value>
Public Property Get MeasuredValue() As Double
    MeasuredValue = This.MeasuredValue
End Property

''' <summary>   Sets the Measured Value. </summary>
''' <param name="a_value">   [Double] the Measured Value. </value>
Public Property Let MeasuredValue(ByVal a_value As Double)
    This.MeasuredValue = a_value
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
    This.ReadingOffset = a_value
End Property

''' <summary>   Gets the Monitoring State. </summary>
''' <value>   [String]. </value>
Public Property Get MonitoringState() As String
    MonitoringState = This.MonitoringState
End Property

''' <summary>   Sets the Monitoring State. </summary>
''' <param name="a_value">   [String] the Monitoring State value. </value>
Public Property Let MonitoringState(ByVal a_value As String)
    This.MonitoringState = a_value
End Property

''' <summary>   Gets the External Trig Monitoring Enabled condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopped on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = This.ExternalTrigMonitoringEnabled
End Property

''' <summary>   Sets the External Trig Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trig Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    This.ExternalTrigMonitoringEnabled = a_value
End Property

''' <summary>   Gets the Auto Increment DUT number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementDutNumberEnabled() As Boolean
    AutoIncrementDutNumberEnabled = This.AutoIncrementDutNumberEnabled
End Property

''' <summary>   Sets the Auto Increment DUT number Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment DUT number Enabled value. </value>
Public Property Let AutoIncrementDutNumberEnabled(ByVal a_value As Boolean)
    This.AutoIncrementDutNumberEnabled = a_value
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
    This.SingleReadEnabled = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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
    This.PauseRequested = a_value
End Property

''' <summary>   Gets the Stop Requested condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopped on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get StopRequested() As Boolean
    StopRequested = This.StopRequested
End Property

''' <summary>   Sets the Stop Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Stop Requested value. </value>
Friend Property Let StopRequested(ByVal a_value As Boolean)
    This.StopRequested = a_value
End Property

''' <summary>   Gets the Timer Interval in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Timer Interval in milliseconds value. </value>
Public Property Let TimerInterval(ByVal a_value As Integer)
    This.TimerInterval = a_value
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Double]. </value>
Public Property Get ElapsedTime() As Double
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Double] the Elapsed Time in milliseconds. </value>
Public Property Let ElapsedTime(ByVal a_value As Double)
    This.ElapsedTime = a_value
End Property

''' <summary>   Gets the Elapsed Time Format. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeFormat() As String
    ElapsedTimeFormat = This.ElapsedTimeFormat
End Property

''' <summary>   Sets the Elapsed Time Format. </summary>
''' <param name="a_value">   [String] the elapsed time format. </value>
Public Property Let ElapsedTimeFormat(ByVal a_value As String)
    This.ElapsedTimeFormat = a_value
End Property

''' <summary>   Gets the Elapsed Time Caption. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeCaption() As String
    ElapsedTimeCaption = This.ElapsedTimeCaption
End Property

''' <summary>   Sets the Elapsed Time Caption. </summary>
''' <param name="a_value">   [String] the elapsed time Caption. </value>
Public Property Let ElapsedTimeCaption(ByVal a_value As String)
    This.ElapsedTimeCaption = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = This.LastErrorMessage
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the last error message value. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    This.LastErrorMessage = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = This.LastMessage
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the last message value. </value>
Public Property Let LastMessage(ByVal a_value As String)
    This.LastMessage = a_value
End Property

''' <summary>   Gets the Last Command Details. </summary>
''' <value>   [String]. </value>
Public Property Get LastCommandDetails() As String
    LastCommandDetails = This.LastCommandDetails
End Property

''' <summary>   Sets the Last Command Details. </summary>
''' <param name="a_value">   [String] the last Command Details value. </value>
Public Property Let LastCommandDetails(ByVal a_value As String)
    This.LastCommandDetails = a_value
End Property

''' <summary>   Gets the Last Command Outcome. </summary>
''' <value>   [Boolean]. </value>
Public Property Get LastCommandOutcome() As Boolean
    LastCommandOutcome = This.LastCommandOutcome
End Property

''' <summary>   Sets the Last Command Outcome. </summary>
''' <param name="a_value">   [Boolean] the last Command Outcome value. </value>
Public Property Let LastCommandOutcome(ByVal a_value As Boolean)
    This.LastCommandOutcome = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Observable executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Open Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Open Connection command can be executed. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = This.OpenConnectionExecutable
End Property

''' <summary>   Gets the condition indicating that the Close Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Close Connection command can be executed. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    CloseConnectionExecutable = This.CloseConnectionExecutable
End Property

''' <summary>   Gets the condition indicating that the toggle Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Connection command can be executed. </value>
Public Property Get ToggleConnectionExecutable() As Boolean
    ToggleConnectionExecutable = This.OpenConnectionExecutable Or This.CloseConnectionExecutable
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE488 Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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

''' <summary>   Gets the condition indicating that the Auto Asset TALK Set Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK Set Command can be executed. </value>
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Clear Readings command
'''              can be executed. </summary>
''' <value>   [Boolean] True if the Clear Readings command can be executed. </value>
Public Property Get ClearReadingsExecutable() As Boolean
    ClearReadingsExecutable = This.ClearReadingsExecutable
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

''' <summary>   Handles the <see cref=cc_isr_Tcp_Scpi.K2700ViewModel"/>.<see cref="cc_isr_Tcp_Scpi.DutMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.DutReadingEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_DutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

    Const p_procedureName = "m_viewModel_DutMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' other measured values are set by property changes.
    Me.MeasuredDutNumber = a_eventArgs.DutNumber
    
    ' Debug.Print "DUT #" * VBA.CStr(Me.MeasuredDutNumber) & " reading '" & Me.MeasuredReading & " value is " & VBA.CStr(Me.MeasuredValue)
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the property change event of the view model. </summary>
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
            Me.FrontInputsValue = This.ViewModel.FrontInputsValue
        Case "FrontInputsCaption"
            Me.FrontInputsCaption = This.ViewModel.FrontInputsCaption
        Case "SelectedDutNumber"
            Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
        Case "DutCount"
            Me.DutCount = This.ViewModel.DutCount
        Case "TargetDutNumber"
            Me.TargetDutNumber = This.ViewModel.TargetDutNumber
        Case "MeasuredDutNumber"
            Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
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
        Case "AutoIncrementDutNumberEnabled"
            Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
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
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        
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
        Case "FunctionNameExecutable"
            This.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable

    End Select
    
End Sub

''' <summary>   Synchronizes the view model with the Observer. </summary>
''' <remarks>   This is required because the view (observer) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a result, the observer retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' set the view model settings that come from the Observer user interface
    ' This.ViewModel.SocketAddress = This.SocketAddress
    ' This.ViewModel.ReadingOffset = This.ReadingOffset
    This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort
    ' This.ViewModel.SessionTimeout = Me.SessionTimeout
    ' This.ViewModel.TimerInterval = Me.TimerInterval
    ' This.ViewModel.MaximumDutNumber = Me.MaximumDutNumber

    ' The initial values must be set for property change to work.
    Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
    Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
    Me.BottomCard = This.ViewModel.BottomCard
    Me.BottomCardFunctionScanList = This.ViewModel.BottomCardFunctionScanList
    Me.DutCount = This.ViewModel.DutCount
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
    Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.MonitoringState = This.ViewModel.MonitoringState
    Me.PauseRequested = This.ViewModel.PauseRequested
    Me.Port = This.ViewModel.Port
    Me.ReadingOffset = This.ViewModel.ReadingOffset
    Me.ReceivedMessage = This.ViewModel.ReceivedMessage
    Me.ReceivedMessageLength = This.ViewModel.ReceivedMessageLength
    Me.ScpiCommands = This.ViewModel.ScpiCommands
    Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
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
    Me.TargetDutNumber = This.ViewModel.TargetDutNumber
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
    This.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable
    
End Sub

''' <summary>   Handles the property change event of the view model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_K2700ViewModel_PropertyChanged"
    
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
    This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>    Gets the Initialization condition of this object. </summary>
''' <value>   [Boolean] True if initialized. </value>
Public Property Get Initialized() As Boolean
    Initialized = Not This.ViewModel Is Nothing
End Property

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
    This.GpibLanControllerPort = DataSheet.GpibLanControllerPort
    This.TopCard = VBA.vbNullString
End Sub

''' <summary>   Initialize this worksheet if not initialized. </summary>
Friend Function Initialize(ByVal a_viewModel As K2700ViewModel) As K2700Observer

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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
    This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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
    This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub
