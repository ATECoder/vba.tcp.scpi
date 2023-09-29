VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700Observer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_sessionTimeout As String = "B1"
Private Const m_inputAddress As String = "B2"
Private Const m_inputPort As String = "B3"
Private Const m_socketAddress As String = "C2"
Private Const m_socketId As String = "D2"
Private Const m_connected As String = "E2"
Private Const m_sentMessage As String = "F2"
Private Const m_receivedMessageLength As String = "J2"
Private Const m_receivedMessage As String = "K2"
Private Const m_elapsedTime As String = "L2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastMessage As String = "B4"
Private Const m_lastErrorMessage As String = "B33"

Private Const m_serialPoll As String = "E11"
Private Const m_statusByte As String = "F11"
Private Const m_standardByte As String = "H11"

Private Const m_srq As String = "C10"
Private Const m_gpibAddress As String = "C12"
Private Const m_readTimeout As String = "C15"
Private Const m_AutoAssertTalk As String = "C18"

Private Const m_readingFormat As String = "C21"
Private Const m_timerInterval As String = "C22"
Private Const m_card1 As String = "C23"
Private Const m_card2 As String = "C24"
Private Const m_card1ScanList As String = "C25"
Private Const m_card2ScanList As String = "C26"
Private Const m_inputs As String = "C27"
Private Const m_channelNumber As String = "C28"
Private Const m_functionName As String = "B29"
Private Const m_functionValue As String = "C29"
Private Const m_functionUnit As String = "D29"
Private Const m_channelCount As String = "C30"
Private Const m_externalTriggerStatus As String = "C31"

Private Const m_measurementTopRow As Integer = 2
Private Const m_measurementNoCol As String = "M"
Private Const m_measurementValueCol As String = "N"

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
    CurrentMeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption

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
    CurrentChannelNumber As Integer
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
    ReadStatsuByteExecutable As Boolean
    GpibAddressGetExecutable As Boolean
    GpibAddressSetExecutable As Boolean
    GpibLanReadTimeoutGetExecutable As Boolean
    GpibLanReadTimeoutSetExecutable As Boolean
        
    MeasurementTopRow As Integer
    MeasurementRow As Integer
    MeasurementColumn As String
    ChannelColumn As String
    ChannelCell As String
    MeasurementCell As String
    ChannelValue As String
    MeasurementValue As String
    
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
Friend Sub ConnectToggleButton_Click()
    If This.ToggleConnectionExecutable Then
        If This.OpenConnectionExecutable Then
            This.ViewModel.OpenConnectionCommand
        ElseIf This.CloseConnectionExecutable Then
            This.ViewModel.CloseConnectionCommand
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
                " Both close and open connection are disabled."
        End If
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE488 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Friend Sub CLSButton_Click()
    If This.ClearExecutionStateExecutable Then
        Me.LastCommandOutcome = This.ViewModel.ClearExecutionStateCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
        
    
End Sub

''' <summary>   Handles the Commands Combo Box change event. </summary>
Friend Sub CommandsComboBox_Change()
    Me.MessageToSend = This.MessageToSend
End Sub

''' <summary>   Handles the Query button click event. </summary>
Friend Sub QueryButton_Click()
    If This.QueryExecutable Then
        This.ViewModel.QueryCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Read button click event. </summary>
Friend Sub ReadButton_Click()
    If This.ReadExecutable Then
        This.ViewModel.ReadCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the reading status after write changed event. </summary>
Friend Sub ReadStatusAfterWriteCheckBox_Change()
    If This.AutoAssertTalkGetExecutable Then _
        This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteEnabled
End Sub

''' <summary>   Handles the Query Status Byte button click event. </summary>
Public Sub QueryStatusByteButton_Click()
    If This.QueryStatusByteExecutable Then
        This.ViewModel.QueryStatusByteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Read Standard Event Status Byte button click event. </summary>
Friend Sub ReadStandardEventStatusButton_Click()
    If This.ReadStandardEventStatusExecutable Then
        This.ViewModel.ReadStandardEventStatusCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Friend Sub RSTButton_Click()
    If This.ResetKnownStateExecutable Then
        This.ViewModel.ResetKnownStateCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Friend Sub WriteButton_Click()
    If This.WriteExecutable Then
        This.ViewModel.WriteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Go To Local button click event. </summary>
Friend Sub GoToLocalButton_Click()
    If This.GoToLocalExecutable Then
        This.ViewModel.GoToLocalCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
    
End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Friend Sub GpibAddressGetButton_Click()
    If This.GpibAddressGetExecutable Then
        This.ViewModel.GpibAddressGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Friend Sub GpibAddressSetButton_Click()
    If This.GpibAddressSetExecutable Then
        This.ViewModel.GpibAddressSetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Friend Sub LocalLockoutButton_Click()
    If This.LocalLockoutExecutable Then
        This.ViewModel.LocalLockoutCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), ".LocalLockoutButton_Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Auto Asset TALK Set button click event. </summary>
Friend Sub AutoAssertTalkSetButton_Click()
    If This.AutoAssertTalkSetExecutable Then
        This.ViewModel.AutoAssertTalkSetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Auto Asset TALK get button click event. </summary>
Friend Sub AutoAssertTalkGetButton_Click()
    If This.AutoAssertTalkGetExecutable Then
        This.ViewModel.AutoAssertTalkGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Friend Sub ReadTimeoutGetButton_Click()
    If This.GpibLanReadTimeoutGetExecutable Then
        This.ViewModel.GpibLanReadTimeoutGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Friend Sub ReadTimeoutSetButton_Click()
    If This.GpibLanReadTimeoutGetExecutable Then
        This.ViewModel.GpibLanReadTimeoutGetCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Friend Sub SDCButton_Click()
    If This.SelectiveDeviceClearExecutable Then
        This.ViewModel.SelectiveDeviceClearCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
   
End Sub

''' <summary>   Handles the serial poll button click event. </summary>
Friend Sub ReadStatsuByteButton_Click()
    If This.ReadStatsuByteExecutable Then
        This.ViewModel.ReadStatsuByteCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Friend Sub SRQButton_Click()
    If This.ServiceRequestedExecutable Then
        This.ViewModel.ServiceRequestedQueryCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the auto increment check box change and click events. </summary>
Friend Sub AutoIncrementCheckBox_Change()
    Me.AutoIncrementChannelNoEnabled = This.AutoIncrementChannelNoEnabled
End Sub

Friend Sub ClearReadingsButton_Click()
    If This.ClearReadingsExecutable Then
        This.ViewModel.ClearReadingsCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the external trigger option button change and click events. </summary>
Friend Sub ExternalTriggerOptionButton_Change()
    If This.ExternalTriggerOptionExecutable Then
         Me.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Selects the measurement function. </summary>
Friend Sub FunctionComboBox_Change()
    
    ' Commands for click
    Me.SenseFunctionName = This.SenseFunctionName
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
''' <remarks>   The check box change event also causes a click event.
'''             This method issues the click event on actual click.
'''             Note: Make sure to not implement the click event,
Friend Sub ImmediateTriggerOptionButton_Change()

    If This.ImmediateTriggerOptionExecutable Then
        Me.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the Inputs check box change event. </summary>
Friend Sub InputsCheckBox_Change()
    This.ViewModel.FrontInputsRequired = Me.FrontInputsRequired
End Sub

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Friend Sub MeasureButton_Click()
    If This.MeasureExecutable Then
        ' make an immediate measure of one or all channels.
        Me.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(Me.LastCommandDetails)
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Friend Sub MonitorExtTrigToggleButton_Change()
    If This.StartMonitoringExecutable Then
        This.ViewModel.StartMonitoringExternalTriggersCommand
    ElseIf This.StopMonitoringExecutable Then
        This.ViewModel.StopMonitoringExternalTriggersCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

''' <summary> Queries the Inputs status. </summary>
Friend Sub QueryInputsButton_Click()
    If This.QueryInputsExecutable Then
        This.ViewModel.QueryInputsCommand
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me), "._Click", _
            " Command not enabled."
    End If
End Sub

Friend Sub SingleMultipleCheckBox_Change()
    Me.SingleReadEnabled = Me.SingleReadEnabled
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
Friend Property Let SocketAddress(ByVal a_value As String)
    If This.SocketAddress <> a_value Then
        This.SocketAddress = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SocketAddress = This.SocketAddress
    End If
End Property

''' <summary>   Gets the socket id. </summary>
''' <value>   [Integer]. </value>
Public Property Get SocketId() As Integer
    SocketId = This.SocketId
End Property

''' <summary>   Sets the socket id. </summary>
''' <param name="a_value">   [Integer] the socket id. </value>
Friend Property Let SocketId(ByVal a_value As Integer)
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
            This.ViewModel.SessionTimeout = This.SessionTimeout
    End If
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
    If This.MessageToSend <> a_value Then
        This.MessageToSend = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MessageToSend = This.MessageToSend
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
            This.ViewModel.SentMessage = This.SentMessage
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
            This.ViewModel.ReceivedMessage = This.ReceivedMessage
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
            This.ViewModel.ReceivedMessageLength = This.ReceivedMessageLength
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
            This.ViewModel.StatusByte = This.StatusByte
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
            This.ViewModel.StandardByte = This.StandardByte
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
            This.ViewModel.ReadStatusAfterWriteEnabled = This.ReadStatusAfterWriteEnabled
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
            This.ViewModel.ScpiCommands = This.ScpiCommands
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
    If This.SerialPollByte <> a_value Then
        This.SerialPollByte = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SerialPollByte = This.SerialPollByte
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
            This.ViewModel.ServiceRequested = This.ServiceRequested
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
            This.ViewModel.GpibAddress = This.GpibAddress
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
            This.ViewModel.AutoAssertTalk = This.AutoAssertTalk
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
                This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort
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
            This.ViewModel.GpibLanReadTimeout = This.GpibLanReadTimeout
    End If
End Property

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
            This.ViewModel.TopCard = This.TopCard
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
            This.ViewModel.TopCardFunctionScanList = This.TopCardFunctionScanList
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
            This.ViewModel.BottomCard = This.BottomCard
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
            This.ViewModel.BottomCardFunctionScanList = This.BottomCardFunctionScanList
    End If
End Property

''' <summary>   Gets the current measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get CurrentMeasurementMode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    CurrentMeasurementMode = This.CurrentMeasurementMode
End Property

''' <summary>   Gets the current measurement mode. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Friend Property Let CurrentMeasurementMode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.CurrentMeasurementMode <> a_value Then
        This.CurrentMeasurementMode = a_value
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
            This.ViewModel.FrontInputsRequired = This.FrontInputsRequired
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
            This.ViewModel.ReadingBeepEnabled = This.ReadingBeepEnabled
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
            This.ViewModel.FrontInputsValue = This.FrontInputsValue
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
            This.ViewModel.FrontInputsCaption = This.FrontInputsCaption
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
            This.ViewModel.ChannelCount = This.ChannelCount
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
            This.ViewModel.SelectedChannelNumber = This.SelectedChannelNumber
    End If
End Property

''' <summary>   Gets the Current Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get CurrentChannelNumber() As Integer
    CurrentChannelNumber = This.CurrentChannelNumber
End Property

''' <summary>   Sets the Current Channel Number. </summary>
''' <param name="a_value">   [Integer] the Current Channel Number value. </value>
Public Property Let CurrentChannelNumber(ByVal a_value As Integer)
    If This.CurrentChannelNumber <> a_value Then
        This.CurrentChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.CurrentChannelNumber = This.CurrentChannelNumber
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
            This.ViewModel.MeasuredChannelNumber = This.MeasuredChannelNumber
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
            This.ViewModel.ContinuousSenseFunctionName = This.ContinuousSenseFunctionName
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
            This.ViewModel.ImmediateSenseFunctionName = This.ImmediateSenseFunctionName
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
            This.ViewModel.ExternalSenseFunctionName = This.ExternalSenseFunctionName
    End If
End Property

''' <summary>   Gets the Sense Function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Internal. Sets the Sense Function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Friend Property Let SenseFunctionName(ByVal a_value As String)
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
            This.ViewModel.SenseFunctionCaption = This.SenseFunctionCaption
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
            This.ViewModel.SenseFunctionUnit = This.SenseFunctionUnit
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
            This.ViewModel.MeasuredReading = This.MeasuredReading
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
            This.ViewModel.MeasuredValueFormat = This.MeasuredValueFormat
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
            This.ViewModel.MeasuredValue = This.MeasuredValue
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
            This.ViewModel.SenseFunctions = This.SenseFunctions
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
                This.ViewModel.ReadingOffset = This.ReadingOffset
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
            This.ViewModel.MonitoringState = This.MonitoringState
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
Friend Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If This.ExternalTrigMonitoringEnabled <> a_value Then
        This.ExternalTrigMonitoringEnabled = a_value
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
            This.ViewModel.SingleReadEnabled = This.SingleReadEnabled
    End If
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
            This.ViewModel.TimerInterval = This.TimerInterval
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
            This.ViewModel.ElapsedTime = This.ElapsedTime
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
            This.ViewModel.ElapsedTimeFormat = This.ElapsedTimeFormat
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
            This.ViewModel.ElapsedTimeCaption = This.ElapsedTimeCaption
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
            This.ViewModel.LastErrorMessage = This.LastErrorMessage
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
            This.ViewModel.LastMessage = This.LastMessage
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Observalbe executables
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

''' <summary>   Gets the condition indicating that the Auto Asset TALK Setcommand
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK Setcommand can be executed. </value>
Public Property Get AutoAssertTalkSetExecutable() As Boolean
    AutoAssertTalkSetExecutable = This.AutoAssertTalkSetExecutable
End Property

''' <summary>   Gets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get ReadStatsuByteExecutable() As Boolean
    ReadStatsuByteExecutable = This.ReadStatsuByteExecutable
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


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Current Measurement Top Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementTopRow() As Integer
    MeasurementTopRow = This.MeasurementTopRow
End Property

''' <summary>   Sets the Current Measurement Top Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Top Row. </value>
Public Property Let MeasurementTopRow(ByVal a_value As Integer)
    If This.MeasurementTopRow <> a_value Then
        This.MeasurementTopRow = a_value
    End If
End Property

''' <summary>   Gets the Current Measurement Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementRow() As Integer
    MeasurementRow = This.MeasurementRow
End Property

''' <summary>   Sets the Current Measurement Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Row. </value>
Public Property Let MeasurementRow(ByVal a_value As Integer)
    If This.MeasurementRow <> a_value Then
        This.MeasurementRow = a_value
    End If
End Property

''' <summary>   Gets the Current Measurement Column. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementColumn() As String
    MeasurementColumn = This.MeasurementColumn
End Property

''' <summary>   Sets the Current Measurement Column. </summary>
''' <param name="a_value">   [String] the Measurement Column. </value>
Public Property Let MeasurementColumn(ByVal a_value As String)
    If This.MeasurementColumn <> a_value Then
        This.MeasurementColumn = a_value
    End If
End Property

''' <summary>   Gets the Current Channel Column. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelColumn() As String
    ChannelColumn = This.ChannelColumn
End Property

''' <summary>   Sets the Current Channel Column. </summary>
''' <param name="a_value">   [String] the Channel Column. </value>
Public Property Let ChannelColumn(ByVal a_value As String)
    If This.ChannelColumn <> a_value Then
        This.ChannelColumn = a_value
    End If
End Property

''' <summary>   Gets the Current Channel Cell. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCell() As String
    ChannelCell = This.ChannelCell
End Property

''' <summary>   Sets the Current Channel Cell. </summary>
''' <param name="a_value">   [String] the Current Channel Cell. </value>
Public Property Let ChannelCell(ByVal a_value As String)
    If This.ChannelCell <> a_value Then
        This.ChannelCell = a_value
    End If
End Property
        
''' <summary>   Gets the Current Measurement Cell. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementCell() As String
    MeasurementCell = This.MeasurementCell
End Property

''' <summary>   Sets the Current Measurement Cell. </summary>
''' <param name="a_value">   [String] the Current Measurement Cell. </value>
Public Property Let MeasurementCell(ByVal a_value As String)
    If This.MeasurementCell <> a_value Then
        This.MeasurementCell = a_value
    End If
End Property

''' <summary>   Gets the Current Measurement Value. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementValue() As String
    MeasurementValue = This.MeasurementValue
End Property

''' <summary>   Sets the Current Measurement Value. </summary>
''' <param name="a_value">   [String] the last CurrentMeasurementValue. </value>
Public Property Let MeasurementValue(ByVal a_value As String)
    If This.MeasurementValue <> a_value Then
        This.MeasurementValue = a_value
        ' Range(Me.MeasurementCell).Value = a_value
    End If
End Property
        
''' <summary>   Gets the Current Channel Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelValue() As String
    ChannelValue = This.ChannelValue
End Property

''' <summary>   Sets the Current Channel Value. </summary>
''' <param name="a_value">   [String] the last CurrentChannelValue. </value>
Public Property Let ChannelValue(ByVal a_value As String)
    If This.ChannelValue <> a_value Then
        This.ChannelValue = a_value
        ' Range(Me.ChannelCell).Value = a_value
    End If
End Property


''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Friend Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)

    Me.MeasurementRow = Me.MeasurementTopRow + a_channelNo - 1
    
    Me.ChannelCell = Me.ChannelColumn & VBA.CStr(Me.MeasurementRow)
    Me.ChannelValue = VBA.CStr(a_channelNo)

    Me.MeasurementCell = Me.MeasurementColumn & VBA.CStr(Me.MeasurementRow)
    Me.MeasurementValue = VBA.Format$(a_value, This.ViewModel.MeasuredValueFormat)
    
End Sub

Private Sub Class_Initialize()
    This.TopCard = VBA.vbNullString
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
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
    
    SetReading Me.MeasuredChannelNumber, Me.MeasuredValue
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the <see cref="cc_isr_Tcp_Scpi.Demo.K2700ViewModel.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Handles the <see cref="cc_isr_Tcp_Scpi.Demo.K2700ViewModel.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event arguments. </param>
Friend Sub m_ViewModel_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Friend Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "m_K2700ViewModel_PropertyChanged"

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
        Case "CurrentChannelNumber"
            Me.CurrentChannelNumber = This.ViewModel.CurrentChannelNumber
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
        Case "CurrentMeasurementMode"
            Me.CurrentMeasurementMode = This.ViewModel.CurrentMeasurementMode
        Case "ReadStatusAfterWriteEnabled"
            Me.ReadStatusAfterWriteEnabled = This.ViewModel.ReadStatusAfterWriteEnabled
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        
        Case "OpenConnectionExecutable"
            This.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
        Case "CloseConnectionExecutable"
            This.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
        Case "ToggleConnectionExecutable"
            ' Me.ToggleConnectionButton.Enabled = This.ViewModel.ToggleConnectionExecutable
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
            'WriteButton.Enabled = This.ViewModel.WriteExecutable
        Case "QueryInputsExecutable"
            This.QueryInputsExecutable = This.ViewModel.QueryInputsExecutable
            'QueryInputsButton.Enabled = This.ViewModel.QueryInputsExecutable
        Case "ClearReadingsExecutable"
            This.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
            'ClearReadingsButton.Enabled = This.ViewModel.ClearReadingsExecutable
        Case "MeasureExecutable"
            This.MeasureExecutable = This.ViewModel.MeasureExecutable
            'MeasureButton.Enabled = This.ViewModel.MeasureExecutable
        Case "StartMonitoringExecutable"
            This.StartMonitoringExecutable = This.ViewModel.StartMonitoringExecutable
            'MonitorExtTrigButton.Enabled = This.ViewModel.MonitorExtTrigExecutable
        Case "StopMonitoringExecutable"
            This.StopMonitoringExecutable = This.ViewModel.StopMonitoringExecutable
            'MonitorExtTrigButton.Enabled = This.ViewModel.MonitorExtTrigExecutable
        Case "ImmediateTriggerOptionExecutable"
            This.ImmediateTriggerOptionExecutable = This.ViewModel.ImmediateTriggerOptionExecutable
            'ImmediateTriggerOptionButton.Enabled = This.ViewModel.ImmediateTriggerOptionExecutable
        Case "ExternalTriggerOptionExecutable"
            This.ExternalTriggerOptionExecutable = This.ViewModel.ExternalTriggerOptionExecutable
            'ExternalTriggerOptionButton.Enabled = This.ViewModel.ExternalTriggerOptionExecutable

        Case "SelectiveDeviceClearExecutable"
            This.SelectiveDeviceClearExecutable = This.ViewModel.SelectiveDeviceClearExecutable
            'SDCButton.Enabled = This.ViewModel.SelectiveDeviceClearExecutable
        Case "GoToLocalExecutable"
            This.GoToLocalExecutable = This.ViewModel.GoToLocalExecutable
            'Button.Enabled = This.ViewModel.GoToLocalExecutable
        Case "LocalLockoutExecutable"
            This.LocalLockoutExecutable = This.ViewModel.LocalLockoutExecutable
            'LocalLockoutButton.Enabled = This.ViewModel.LocalLockoutExecutable
        Case "AutoAssertTalkSetExecutable"
            This.AutoAssertTalkSetExecutable = This.ViewModel.AutoAssertTalkSetExecutable
            'AutoAssertTalkSetButton.Enabled = This.ViewModel.AutoAssertTalkSetExecutable
        Case "AutoAssertTalkGetExecutable"
            This.AutoAssertTalkGetExecutable = This.ViewModel.AutoAssertTalkGetExecutable
            'AutoAssertTalkGetButton.Enabled = This.ViewModel.AutoAssertTalkGetExecutable
        Case "ReadStatsuByteExecutable"
            This.ReadStatsuByteExecutable = This.ViewModel.ReadStatsuByteExecutable
            'ReadStatsuByteButton.Enabled = This.ViewModel.ReadStatsuByteExecutable
        Case "ServiceRequestedExecutable"
            This.ServiceRequestedExecutable = This.ViewModel.ServiceRequestedExecutable
            'SRQButton.Enabled = This.ViewModel.ServiceRequestedExecutable
        Case "GpibAddressGetExecutable"
            This.GpibAddressGetExecutable = This.ViewModel.GpibAddressGetExecutable
            'GpibAddressGetButton.Enabled = This.ViewModel.GpibAddressGetExecutable
        Case "GpibAddressSetExecutable"
            This.GpibAddressSetExecutable = This.ViewModel.GpibAddressSetExecutable
            'GpibAddressSetButton.Enabled = This.ViewModel.GpibAddressSetExecutable
        Case "GpibLanReadTimeoutGetExecutable"
            This.GpibLanReadTimeoutGetExecutable = This.ViewModel.GpibLanReadTimeoutGetExecutable
            'ReadTimeoutGetButton.Enabled = This.ViewModel.GpibLanReadTimeoutGetExecutable
        Case "GpibLanReadTimeoutSetExecutable"
            This.GpibLanReadTimeoutSetExecutable = This.ViewModel.GpibLanReadTimeoutSetExecutable
            'ReadTimeoutSetButton.Enabled = This.ViewModel.GpibLanReadTimeoutSetExecutable
        
        Case "MessageToSendExecutable"
            This.MessageToSendExecutable = This.ViewModel.MessageToSendExecutable
            'CommandsComboBox.Enabled = This.ViewModel.MessageToSendExecutable
        Case "FunctionNameExecutable"
            This.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable
            'FunctionComboBox.Enabled = This.ViewModel.FunctionNameExecutable

    End Select
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
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
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Initialize this worksheet if not initialized. </summary>
Friend Function Initialize(ByVal a_viewModel As K2700ViewModel) As K2700Observer

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.ViewModel = a_viewModel
    
    ' set the view model settings that come from the Observer user interface
    This.ViewModel.SocketAddress = This.SocketAddress
    This.ViewModel.ReadingOffset = This.ReadingOffset
    This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort

    Set m_viewModel = a_viewModel
    
    If Not This.Initialized Then
    End If
    
    Me.MeasurementTopRow = m_measurementTopRow
    Me.ChannelColumn = m_measurementNoCol
    Me.MeasurementColumn = m_measurementValueCol

    This.Initialized = True
        
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
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

