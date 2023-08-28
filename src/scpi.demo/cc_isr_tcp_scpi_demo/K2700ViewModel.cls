VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700ViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A public creatable singleton View Model for the
'''             2700 instrument control. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a property value changes. </summary>
Implements cc_isr_Core.INotifyPropertyChanged

''' <summary>   Notifies all registered handlers of a property value change. </summary>
Public Event PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

''' <summary>   The event that is raised upon completion of a channel measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="ChannelReadingEventArgs"/> event args. </param>
Public Event ChannelMeasured(ByVal a_eventArgs As ChannelReadingEventArgs)

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a connection is changing and was changed. </summary>
Implements cc_isr_Winsock.IConnectable

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a measurements are completed. </summary>
Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private Type this_

    Stopwatch As Stopwatch

    K2700 As K2700

    Device As cc_isr_Ieee488.Device

    ViSession As cc_isr_Ieee488.ViSession

    ' view model properties
    ElapsedTime As Integer
    ElapsedTimeFormat As String
    ElapsedTimeCaption As String
    LastErrorMessage As String
    LastMessage As String

    ' Winsock properties
    Host As String
    Port As Integer
    SocketAddress As String
    SocketId As Integer
    ReceiveTimeout As Integer

    ' ieee488 control properties
    SessionReadTimeout As Integer
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
    ReadAfterWriteEnabled As Boolean
    ReadTimeout As Integer
    
    ' 2700 properties
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    FrontInputsRequired As Boolean
    FrontInputsValue As Boolean
    FrontInputsCaption As String
    ChannelCount As Integer
    SelectedChannelNumber As Integer
    CurrentChannelNumber As Integer
    MeasuredChannelNumber As Integer
    SenseFunctionName As String
    SenseFunctionCaption As String
    SenseFunctionUnit As String
    MeasuredReading As String
    MeasuredValueFormat As String
    MeasuredValue As Double
    SenseFunctions As Collection
    ReadingOffset As Double
    
    ' service request monitoring timer control
    PauseRequested As Boolean
    TimerInterval As Integer
    ' true when handling the timer event to prevent timer
    ' event handler overrun. This effectively synchronizes the timer
    ' event with the instrument.
    HandlingTimerEvent As Boolean

    ' Service request monitoring properties
    ' this must be enable before stating monitoring.
    ExternalTrigMonitoringEnabled As Boolean
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
    ReadStatusByteExecutable As Boolean
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
    MonitorExtTrigExecutable As Boolean
    ExternalTriggerOptionExecutable As Boolean

    ' GPIB-Lan executables
    SelectiveDeviceClearExecutable As Boolean
    GoToLocalExecutable As Boolean
    LocalLockoutExecutable As Boolean
    ReadAfterWriteSetExecutable As Boolean
    ReadAfterWriteGetExecutable As Boolean
    SerialPollExecutable As Boolean
    GpibAddressGetExecutable As Boolean
    GpibAddressSetExecutable As Boolean
    ReadTimeoutGetExecutable As Boolean
    ReadTimeoutSetExecutable As Boolean


End Type

Private This As this_

''' <summary>   Initialize this worksheet. </summary>
Public Sub Initialize()

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' populate the combo box commands
    PopulateScpiCommands

    ' populate the function box commands
    PopulateSenseFunctions

    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize(New ErrTracer)
    Set m_k2700 = This.K2700
    Set This.Device = This.K2700.Device
    Set This.ViSession = This.Device.ViSession
    Set This.Stopwatch = cc_isr_Core_IO.Factory.NewStopwatch()
    
    m_k2700_ConnectionChanged cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize()

    Me.ToggleConnectionExecutableSetter
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the host dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Sets the Host. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let Host(ByVal a_value As String)
    If This.Host <> a_value Then
        This.Host = a_value
        Me.OnPropertyChanged "Host"
        Me.ToggleConnectionExecutableSetter
    End If
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Integer
    Port = This.Port
End Property

''' <summary>   Sets the Port. </summary>
''' <param name="a_value">   [Integer] the Port. </value>
Public Property Let Port(ByVal a_value As Integer)
    If This.Port <> a_value Then
        This.Port = a_value
        This.Port = a_value
        Me.OnPropertyChanged "Port"
        Me.ToggleConnectionExecutableSetter
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
        Me.OnPropertyChanged "SocketAddress"
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
        This.SocketId = a_value
        Me.OnPropertyChanged "SocketId"
    End If
End Property

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceiveTimeout() As Integer
    ReceiveTimeout = This.ReceiveTimeout
End Property

''' <summary>   Sets the receive timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let ReceiveTimeout(ByVal a_value As Integer)
    If This.ReceiveTimeout <> a_value Then
        This.ReceiveTimeout = a_value
        Me.OnPropertyChanged "ReceiveTimeout"
        Me.ToggleConnectionExecutableSetter
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Open Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Open Connection command can be executed. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = This.OpenConnectionExecutable
End Property

''' <summary>   Sets the condition indicating that the Open Connection command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Open Connection command can be executed. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    If This.OpenConnectionExecutable <> a_value Then
        This.OpenConnectionExecutable = a_value
        Me.OnPropertyChanged "OpenConnectionExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Close Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Close Connection command can be executed. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    CloseConnectionExecutable = This.CloseConnectionExecutable
End Property

''' <summary>   Sets the condition indicating that the Close Connection command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Close Connection command can be executed. </param>
Public Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    If This.CloseConnectionExecutable <> a_value Then
        This.CloseConnectionExecutable = a_value
        Me.OnPropertyChanged "CloseConnectionExecutable"
    End If
End Property

''' <summary>   Sets the condition indicating that the Toggle Connection command
'''             can be executed. </summary>
Friend Sub ToggleConnectionExecutableSetter()
    Me.ToggleConnectionExecutable = VBA.vbNullString <> Me.Host _
        And (5025 = Me.Port Or 1234 = Me.Port) _
        And 0 < Me.ReceiveTimeout
End Sub

''' <summary>   Gets the condition indicating that the Toggle Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Connection command can be executed. </value>
Public Property Get ToggleConnectionExecutable() As Boolean
    ToggleConnectionExecutable = This.ToggleConnectionExecutable
End Property

''' <summary>   Sets the condition indicating that the Toggle Connection command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Toggle Connection command can be executed. </param>
Public Property Let ToggleConnectionExecutable(ByVal a_value As Boolean)
    If This.ToggleConnectionExecutable <> a_value Then
        This.ToggleConnectionExecutable = a_value
        Me.OnPropertyChanged "ToggleConnectionExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Toggle connection button click event. </summary>
''' <param name=a_connect>   [Boolean] true to connect; otherwise, disconnect. </param>
Public Sub ToggleConnectionCommand(ByVal a_connect As Boolean)

    Const p_procedureName = "ToggleConnectionCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    If Not a_connect And Me.Connected Then
    
        This.K2700.CloseConnection
    
    ElseIf a_connect And Not Me.Connected Then

        This.K2700.OpenConnection This.Host, This.Port, This.ReceiveTimeout
    
        This.ViSession.SessionReadTimeout = This.SessionReadTimeout
    
    End If

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 properties
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.Device"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.Device"/>. </value>
Public Property Get Device() As cc_isr_Ieee488.Device
    Set Device = This.Device
End Property

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.ViSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.ViSession"/>. </value>
Public Property Get ViSession() As cc_isr_Ieee488.ViSession
    Set ViSession = This.ViSession
End Property

''' <summary>   Gets the session read timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = This.SessionReadTimeout
End Property

''' <summary>   Sets the session read timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionReadTimeout(ByVal a_value As Integer)
    If This.SessionReadTimeout <> a_value Then
        This.SessionReadTimeout = a_value
        Me.OnPropertyChanged "SessionReadTimeout"
    End If
End Property

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
        Me.OnPropertyChanged "MessageToSend"
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
        Me.OnPropertyChanged "SentMessage"
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
        Me.OnPropertyChanged "ReceivedMessage"
        Me.ReceivedMessageLength = VBA.Len(a_value)
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
        Me.OnPropertyChanged "ReceivedMessageLength"
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
        Me.OnPropertyChanged "StatusByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequests.RequestingService = _
                a_value And cc_isr_Ieee488.ServiceRequests.RequestingService
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
        Me.OnPropertyChanged "StandardByte"
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
        Me.OnPropertyChanged "ReadStatusAfterWriteEnabled"
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
        Me.OnPropertyChanged "ScpiCommands"
    End If
End Property

''' <summary>   Populates the Commands collection. </summary>
Public Sub PopulateScpiCommands()

    Dim p_scpiCommands As New Collection
    p_scpiCommands.Add cc_isr_Ieee488.ClearExecutionStateCommand
    p_scpiCommands.Add cc_isr_Ieee488.ClearExecutionStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.IdentityQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.OperationCompleteCommand
    p_scpiCommands.Add cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.OptionsQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.ResetKnownStateCommand
    p_scpiCommands.Add cc_isr_Ieee488.ResetKnownStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardEventEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.StandardEventEnableQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.StandardEventStatusQueryCommand
    p_scpiCommands.Add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.ServiceRequestEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.ServiceRequestQueryCommand
    p_scpiCommands.Add cc_isr_Ieee488.WaitCommand
    p_scpiCommands.Add cc_isr_Ieee488.WaitCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
   
    p_scpiCommands.Add "SYST:BEEP OFF"
    p_scpiCommands.Add "SYST:CLE"
    p_scpiCommands.Add "SYST:ERR?"
    p_scpiCommands.Add "SYST:FRSW?"
    p_scpiCommands.Add "FUNC 'RES',(@101:120)"
    p_scpiCommands.Add "FUNC 'RES',(@201:220)"
    p_scpiCommands.Add "FORM:ELEM READ"
    p_scpiCommands.Add "STAT:PRES"
    p_scpiCommands.Add "STAT:MEAS:ENAB 32"
    p_scpiCommands.Add "FETCH?"
    p_scpiCommands.Add "READ?"

    Me.ScpiCommands = p_scpiCommands

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Reset Known State command can be executed. </value>
Public Property Get ResetKnownStateExecutable() As Boolean
    ResetKnownStateExecutable = This.ResetKnownStateExecutable
End Property

''' <summary>   Sets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Reset Known State command can be executed. </param>
Public Property Let ResetKnownStateExecutable(ByVal a_value As Boolean)
    If This.ResetKnownStateExecutable <> a_value Then
        This.ResetKnownStateExecutable = a_value
        Me.OnPropertyChanged "ResetKnownStateExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Clear Execution State Command can be executed. </value>
Public Property Get ClearExecutionStateExecutable() As Boolean
    ClearExecutionStateExecutable = This.ClearExecutionStateExecutable
End Property

''' <summary>   Sets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Clear Execution State Command can be executed. </param>
Public Property Let ClearExecutionStateExecutable(ByVal a_value As Boolean)
    If This.ClearExecutionStateExecutable <> a_value Then
        This.ClearExecutionStateExecutable = a_value
        Me.OnPropertyChanged "ClearExecutionStateExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Query command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query command can be executed. </value>
Public Property Get QueryExecutable() As Boolean
    QueryExecutable = This.QueryExecutable
End Property

''' <summary>   Sets the condition indicating that the Query command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query command can be executed. </param>
Public Property Let QueryExecutable(ByVal a_value As Boolean)
    If This.QueryExecutable <> a_value Then
        This.QueryExecutable = a_value
        Me.OnPropertyChanged "QueryExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read command can be executed. </value>
Public Property Get ReadExecutable() As Boolean
    ReadExecutable = This.ReadExecutable
End Property

''' <summary>   Sets the condition indicating that the Read command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read command can be executed. </param>
Public Property Let ReadExecutable(ByVal a_value As Boolean)
    If This.ReadExecutable <> a_value Then
        This.ReadExecutable = a_value
        Me.OnPropertyChanged "ReadExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status Byte command can be executed. </value>
Public Property Get ReadStatusByteExecutable() As Boolean
    ReadStatusByteExecutable = This.ReadStatusByteExecutable
End Property

''' <summary>   Sets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read Status Byte command can be executed. </param>
Public Property Let ReadStatusByteExecutable(ByVal a_value As Boolean)
    If This.ReadStatusByteExecutable <> a_value Then
        This.ReadStatusByteExecutable = a_value
        Me.OnPropertyChanged "ReadStatusByteExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Standard Event Status command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Standard Event Status command can be executed. </value>
Public Property Get ReadStandardEventStatusExecutable() As Boolean
    ReadStandardEventStatusExecutable = This.ReadStandardEventStatusExecutable
End Property

''' <summary>   Sets the condition indicating that the Read Standard Event Status command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read Standard Event Status command can be executed. </param>
Public Property Let ReadStandardEventStatusExecutable(ByVal a_value As Boolean)
    If This.ReadStandardEventStatusExecutable <> a_value Then
        This.ReadStandardEventStatusExecutable = a_value
        Me.OnPropertyChanged "ReadStandardEventStatusExecutable"
    End If
End Property
''' <summary>   Gets the condition indicating that the Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Write command can be executed. </value>
Public Property Get WriteExecutable() As Boolean
    WriteExecutable = This.WriteExecutable
End Property

''' <summary>   Sets the condition indicating that the Write command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Write command can be executed. </param>
Public Property Let WriteExecutable(ByVal a_value As Boolean)
    If This.WriteExecutable <> a_value Then
        This.WriteExecutable = a_value
        Me.OnPropertyChanged "WriteExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that Message To Send command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Message to Send command can be executed. </value>
Public Property Get MessageToSendExecutable() As Boolean
    MessageToSendExecutable = This.MessageToSendExecutable
End Property

''' <summary>   Sets the Message to send command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Message to send command can be executed. </param>
Public Property Let MessageToSendExecutable(ByVal a_value As Boolean)
    If This.MessageToSendExecutable <> a_value Then
        This.MessageToSendExecutable = a_value
        Me.OnPropertyChanged "MessageToSendExecutable"
    End If
End Property


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Clear Known State button click event. </summary>
Public Sub ClearExecutionStateCommand()

    Const p_procedureName = "ClearExecutionStateCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.LastMessage = "Clearing execution state."

    Me.RestartStopwatchCommand

    This.Device.ClearExecutionState

    Me.ReadStopwatchCommand

    Me.LastMessage = "Done clearing execution state."

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

''' <summary>   Writes the <see cref="MessageToSend"/> and update elapsed times. </summary>
Public Sub WriteCommand()

    Const p_procedureName = "WriteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.WriteLine This.MessageToSend

    Me.ReadStopwatchCommand

    Me.SentMessage = This.MessageToSend

    If Me.ReadStatusAfterWriteEnabled Then

        If This.ViSession.UsingGpibLan Then _
            Me.SerialPollByte = This.ViSession.GpibLan.SerialPoll

        Dim p_isQuery As Boolean
        p_isQuery = cc_isr_Core.StringExtensions.EndsWith(This.MessageToSend, "?")
        If Not p_isQuery Then
            Me.StatusByte = This.Device.QueryServiceRequestStatus
            Me.StandardByte = This.Device.QueryStandardEventsStatus
        End If

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Query the instrument using the message to send and update then elapsed tiMe. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Public Sub QueryCommand()

    Const p_procedureName = "QueryCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.ReceivedMessage = This.ViSession.QueryLine(Me.MessageToSend)

    Me.ReadStopwatchCommand

    Me.SentMessage = Me.MessageToSend
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

''' <summary>   Read from the instrument updating the elap[es times. </summary>
Public Sub ReadCommand()

    Const p_procedureName = "ReadCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.ReceivedMessage = This.ViSession.Read()

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStandardEventStatusCommand()

    Const p_procedureName = "ReadStandardEventStatusCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.StandardByte = This.Device.QueryStandardEventsStatus

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Public Sub ResetKnownStateCommand()

    Const p_procedureName = "ResetKnownStateCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.Device.ResetKnownState

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query Status Byte button click event. </summary>
Public Sub QueryStatusByteCommand()

    Const p_procedureName = "QueryStatusByteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.StatusByte = This.Device.QueryServiceRequestStatus()

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

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
        Me.OnPropertyChanged "SerialPollByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequests.RequestingService = _
                a_value And cc_isr_Ieee488.ServiceRequests.RequestingService
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
        Me.OnPropertyChanged "ServiceRequested"
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
        Me.OnPropertyChanged "GpibAddress"
    End If
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = This.ReadAfterWriteEnabled
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let ReadAfterWriteEnabled(ByVal a_value As Boolean)
    If This.ReadAfterWriteEnabled <> a_value Then
        This.ReadAfterWriteEnabled = a_value
        Me.OnPropertyChanged "ReadAfterWriteEnabled"
    End If
End Property

''' <summary>   Gets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadTimeout() As Integer
    ReadTimeout = This.ReadTimeout
End Property

''' <summary>   Sets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let ReadTimeout(ByVal a_value As Integer)
    If This.ReadTimeout <> a_value Then
        This.ReadTimeout = a_value
        Me.OnPropertyChanged "ReadTimeout"
        Me.ToggleConnectionExecutableSetter
    End If
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

''' <summary>   Sets the condition indicating that the Selective Device Clear command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Selective Device Clear command can be executed. </param>
Public Property Let SelectiveDeviceClearExecutable(ByVal a_value As Boolean)
    If This.SelectiveDeviceClearExecutable <> a_value Then
        This.SelectiveDeviceClearExecutable = a_value
        Me.OnPropertyChanged "SelectiveDeviceClearExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Go To Local command can be executed. </value>
Public Property Get GoToLocalExecutable() As Boolean
    GoToLocalExecutable = This.GoToLocalExecutable
End Property

''' <summary>   Sets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Go To Local command can be executed. </param>
Public Property Let GoToLocalExecutable(ByVal a_value As Boolean)
    If This.GoToLocalExecutable <> a_value Then
        This.GoToLocalExecutable = a_value
        Me.OnPropertyChanged "GoToLocalExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Local Lockout command can be executed. </value>
Public Property Get LocalLockoutExecutable() As Boolean
    LocalLockoutExecutable = This.LocalLockoutExecutable
End Property

''' <summary>   Sets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Local Lockout command can be executed. </param>
Public Property Let LocalLockoutExecutable(ByVal a_value As Boolean)
    If This.LocalLockoutExecutable <> a_value Then
        This.LocalLockoutExecutable = a_value
        Me.OnPropertyChanged "LocalLockoutExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read After Write Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read After Write Get command can be executed. </value>
Public Property Get ReadAfterWriteGetExecutable() As Boolean
    ReadAfterWriteGetExecutable = This.ReadAfterWriteGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Read After Write Get command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read After Write Get command can be executed. </param>
Public Property Let ReadAfterWriteGetExecutable(ByVal a_value As Boolean)
    If This.ReadAfterWriteGetExecutable <> a_value Then
        This.ReadAfterWriteGetExecutable = a_value
        Me.OnPropertyChanged "ReadAfterWriteGetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read After Write Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read After Write Set command can be executed. </value>
Public Property Get ReadAfterWriteSetExecutable() As Boolean
    ReadAfterWriteSetExecutable = This.ReadAfterWriteSetExecutable
End Property

''' <summary>   Sets the condition indicating that the Read After Write Set command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read After Write Set command can be executed. </param>
Public Property Let ReadAfterWriteSetExecutable(ByVal a_value As Boolean)
    If This.ReadAfterWriteSetExecutable <> a_value Then
        This.ReadAfterWriteSetExecutable = a_value
        Me.OnPropertyChanged "ReadAfterWriteSetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get SerialPollExecutable() As Boolean
    SerialPollExecutable = This.SerialPollExecutable
End Property

''' <summary>   Sets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Serial Poll command can be executed. </param>
Public Property Let SerialPollExecutable(ByVal a_value As Boolean)
    If This.SerialPollExecutable <> a_value Then
        This.SerialPollExecutable = a_value
        Me.OnPropertyChanged "SerialPollExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Service Requested command can be executed. </value>
Public Property Get ServiceRequestedExecutable() As Boolean
    ServiceRequestedExecutable = This.ServiceRequestedExecutable
End Property

''' <summary>   Sets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Service Requested command can be executed. </param>
Public Property Let ServiceRequestedExecutable(ByVal a_value As Boolean)
    If This.ServiceRequestedExecutable <> a_value Then
        This.ServiceRequestedExecutable = a_value
        Me.OnPropertyChanged "ServiceRequestedExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Get Command can be executed. </value>
Public Property Get GpibAddressGetExecutable() As Boolean
    GpibAddressGetExecutable = This.GpibAddressGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Gpib Address Get Command can be executed. </param>
Public Property Let GpibAddressGetExecutable(ByVal a_value As Boolean)
    If This.GpibAddressGetExecutable <> a_value Then
        This.GpibAddressGetExecutable = a_value
        Me.OnPropertyChanged "GpibAddressGetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Set Command can be executed. </value>
Public Property Get GpibAddressSetExecutable() As Boolean
    GpibAddressSetExecutable = This.GpibAddressSetExecutable
End Property

''' <summary>   Sets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Gpib Address Set Command can be executed. </param>
Public Property Let GpibAddressSetExecutable(ByVal a_value As Boolean)
    If This.GpibAddressSetExecutable <> a_value Then
        This.GpibAddressSetExecutable = a_value
        Me.OnPropertyChanged "GpibAddressSetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Get command can be executed. </value>
Public Property Get ReadTimeoutGetExecutable() As Boolean
    ReadTimeoutGetExecutable = This.ReadTimeoutGetExecutable
End Property

''' <summary>   Gets the Read Timeout Get command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Get command can be executed. </param>
Public Property Let ReadTimeoutGetExecutable(ByVal a_value As Boolean)
    If This.ReadTimeoutGetExecutable <> a_value Then
        This.ReadTimeoutGetExecutable = a_value
        Me.OnPropertyChanged "ReadTimeoutGetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Set command can be executed. </value>
Public Property Get ReadTimeoutSetExecutable() As Boolean
    ReadTimeoutSetExecutable = This.ReadTimeoutSetExecutable
End Property

''' <summary>   Sets the Read Timeout Set command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Set command can be executed. </param>
Public Property Let ReadTimeoutSetExecutable(ByVal a_value As Boolean)
    If This.ReadTimeoutSetExecutable <> a_value Then
        This.ReadTimeoutSetExecutable = a_value
        Me.OnPropertyChanged "ReadTimeoutSetExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Selective Device Clear button click event. </summary>
Public Sub SelectiveDeviceClearCommand()

    Const p_procedureName = "SelectiveDeviceClearCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.SelectiveDeviceClear

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Serial Poll button click event. </summary>
Public Sub SerialPollCommand()

    Const p_procedureName = "SerialPollCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.StatusByte = This.Device.ReadStatusByte(True)

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Write a message to the device via the GPIB Lan. </summary>
Public Sub QueryGpibLanCommand()

    Const p_procedureName = "QueryGpibLanCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    If 0 < This.ViSession.GpibLan.SendToController(Me.MessageToSend) Then
        Me.ReceivedMessage = This.ViSession.GpibLan.ReceiveFromDevice()
    End If

    Me.ReadStopwatchCommand

    Me.SentMessage = Me.MessageToSend

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Read from the device using the GPIB LAN. </summary>
Public Sub ReadGpibLanCommand()

    Const p_procedureName = "ReadGpibLanCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.ReceivedMessage = This.ViSession.GpibLan.ReceiveFromDevice()

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Write a message to the device via the GPIB Lan. </summary>
Public Sub WriteGpibLanCommand()

    Const p_procedureName = "WriteGpibLanCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.SendToController (Me.MessageToSend)

    Me.ReadStopwatchCommand

    Me.SentMessage = Me.MessageToSend

    If Me.ReadStatusAfterWriteEnabled Then

        If This.ViSession.UsingGpibLan Then _
            Me.SerialPollByte = This.ViSession.GpibLan.SerialPoll

    End If


' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Go To Local button click event. </summary>
Public Sub GoToLocalCommand()

    Const p_procedureName = "GoToLocalCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.GoToLocal

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Public Sub LocalLockoutCommand()

    Const p_procedureName = "LocalLockoutCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.LocalLockout

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write set button click event. </summary>
Public Sub ReadAfterWriteSetCommand()

    Const p_procedureName = "ReadAfterWriteSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.ReadAfterWriteEnabledSetter Me.ReadAfterWriteEnabled

    Me.ReadStopwatchCommand

    This.ViSession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write get button click event. </summary>
Public Sub ReadAfterWriteGetCommand()

    Const p_procedureName = "ReadAfterWriteGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.ReadAfterWriteEnabled = This.ViSession.GpibLan.ReadAfterWriteEnabledGetter

    Me.ReadStopwatchCommand

    This.ViSession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Public Sub ServiceRequestedQueryCommand()

    ' this does nothing
    ' Me.ServiceRequested = This.VISession.GpibLan.ServiceRequested

    ' do this instead.
    Me.SerialPollCommand

End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Public Sub GpibAddressGetCommand()

    Const p_procedureName = "GpibAddressGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.GpibAddress = This.ViSession.GpibLan.GpibAddressGetter

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Public Sub GpibAddressSetCommand()

    Const p_procedureName = "GpibAddressSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.GpibAddressSetter Me.GpibAddress

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Public Sub ReadTimeoutGetCommand()

    Const p_procedureName = "ReadTimeoutGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.ReadTimeout = This.ViSession.GpibLan.ReadTimeoutGetter

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Public Sub ReadTimeoutSetCommand()

    Const p_procedureName = "ReadTimeoutSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    This.ViSession.GpibLan.ReadTimeoutSetter Me.ReadTimeout

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Tcp_Scpi.K2700"/>. </summary>
''' <value>   An <see cref="cc_isr_Tcp_Scpi.K2700"/>. </value>
Public Property Get K2700() As cc_isr_Tcp_Scpi.K2700
    Set K2700 = This.K2700
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
        Me.OnPropertyChanged "TopCard"
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
        Me.OnPropertyChanged "TopCardFunctionScanList"
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
        Me.OnPropertyChanged "BottomCard"
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
        Me.OnPropertyChanged "BottomCardFunctionScanList"
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
        Me.OnPropertyChanged "FrontInputsRequired"
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
        Me.OnPropertyChanged "FrontInputsValue"
        Me.FrontInputsCaption = IIf(a_value, "Front", "Rear")
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
        Me.OnPropertyChanged "FrontInputsCaption"
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
        Me.OnPropertyChanged "ChannelCount"
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
        Me.OnPropertyChanged "SelectedChannelNumber"
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
        Me.OnPropertyChanged "CurrentChannelNumber"
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
        Me.OnPropertyChanged "MeasuredChannelNumber"
    End If
End Property

''' <summary>   Gets the  Sense Function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Sets the  Sense Function Name. </summary>
''' <param name="a_value">   [String] the  Sense Function Name. </value>
Public Property Let SenseFunctionName(ByVal a_value As String)
    If This.SenseFunctionName <> a_value Then
        This.SenseFunctionName = a_value
        Me.OnPropertyChanged "SenseFunctionName"
        Me.OnSenseFunctionChanged
    End If
End Property

''' <summary>   Handles the sense function change. </summary>
Public Sub OnSenseFunctionChanged()

    Select Case Me.SenseFunctionName

        Case "CONT"
    
            Me.SenseFunctionCaption = "Continuity"
            Me.SenseFunctionUnit = "T/F"
    
        Case "CURR:AC"
    
            Me.SenseFunctionCaption = "Current AC"
            Me.SenseFunctionUnit = "A"
    
        Case "CURR:DC"
    
            Me.SenseFunctionCaption = "Current DC"
            Me.SenseFunctionUnit = "A"
    
        Case "FREQ"
    
            Me.SenseFunctionCaption = "Frequency"
            Me.SenseFunctionUnit = "Hz"
    
        Case "FRES"
    
            Me.SenseFunctionCaption = "4-Wire Res."
            Me.SenseFunctionUnit = VBA.Chr$(234)
    
        Case "PER"
    
            Me.SenseFunctionCaption = "Period"
            Me.SenseFunctionUnit = "s"
    
        Case "RES"
    
            Me.SenseFunctionCaption = "2-Wire Res."
            Me.SenseFunctionUnit = VBA.Chr$(234)
    
        Case "TEMP"
    
            Me.SenseFunctionCaption = "Temperature"
            Me.SenseFunctionUnit = "K"
    
        Case "VOLT:AC"
    
            Me.SenseFunctionCaption = "Voltage AC"
            Me.SenseFunctionUnit = "v"
    
        Case "VOLT:DC"
    
            Me.SenseFunctionCaption = "Voltage DC"
            Me.SenseFunctionUnit = "v"
    
        Case Else
    
            Me.SenseFunctionCaption = "unknown"
            Me.SenseFunctionUnit = "?"

    End Select

End Sub

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
        Me.OnPropertyChanged "SenseFunctionCaption"
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
        Me.OnPropertyChanged "SenseFunctionUnit"
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
        Me.OnPropertyChanged "MeasuredReading"
        Me.MeasuredValue = CDbl(a_value)
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
        Me.OnPropertyChanged "MeasuredValueFormat"
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
        Me.OnPropertyChanged "MeasuredValue"
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
        Me.OnPropertyChanged "SenseFunctions"
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
        Me.OnPropertyChanged "ReadingOffset"
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
        Me.OnPropertyChanged "MonitoringState"
    End If
End Property

''' <summary>   Gets the External Trigger Monitoring Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = This.ExternalTrigMonitoringEnabled
End Property

''' <summary>   Sets the External Trigger Monitoring Enabled. </summary>
''' <param name="a_value">   [Boolean] the External Trigger Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If This.ExternalTrigMonitoringEnabled <> a_value Then
        This.ExternalTrigMonitoringEnabled = a_value
        Me.OnPropertyChanged "ExternalTrigMonitoringEnabled"
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
        Me.OnPropertyChanged "AutoIncrementChannelNoEnabled"
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
        Me.OnPropertyChanged "SingleReadEnabled"
    End If
End Property

''' <summary>   Populates the collection of sense functions. </summary>
Public Sub PopulateSenseFunctions()

    Dim p_senseFunctions As New Collection
    p_senseFunctions.Add "CONT"
    p_senseFunctions.Add "CURR:AC"
    p_senseFunctions.Add "CURR:DC"
    p_senseFunctions.Add "FREQ"
    p_senseFunctions.Add "FRES"
    p_senseFunctions.Add "PER"
    p_senseFunctions.Add "RES"
    p_senseFunctions.Add "TEMP"
    p_senseFunctions.Add "VOLT:AC"
    p_senseFunctions.Add "VOLT:DC"

    Me.SenseFunctions = p_senseFunctions

End Sub

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

''' <summary>   Sets the condition indicating that the Clear Readings command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Clear Readings command can be executed. </param>
Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If This.ClearReadingsExecutable <> a_value Then
        This.ClearReadingsExecutable = a_value
        Me.OnPropertyChanged "ClearReadingsExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Inputs command can be executed. </value>
Public Property Get QueryInputsExecutable() As Boolean
    QueryInputsExecutable = This.QueryInputsExecutable
End Property

''' <summary>   Sets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query Inputs command can be executed. </param>
Public Property Let QueryInputsExecutable(ByVal a_value As Boolean)
    If This.QueryInputsExecutable <> a_value Then
        This.QueryInputsExecutable = a_value
        Me.OnPropertyChanged "QueryInputsExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Measure command can be executed. </value>
Public Property Get MeasureExecutable() As Boolean
    MeasureExecutable = This.MeasureExecutable
End Property

''' <summary>   Sets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Measure command can be executed. </param>
Public Property Let MeasureExecutable(ByVal a_value As Boolean)
    If This.MeasureExecutable <> a_value Then
        This.MeasureExecutable = a_value
        Me.OnPropertyChanged "MeasureExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Monitor External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Monitor External Trigger command can be executed. </value>
Public Property Get MonitorExtTrigExecutable() As Boolean
    MonitorExtTrigExecutable = This.MonitorExtTrigExecutable
End Property

''' <summary>   Sets the condition indicating that the Monitor External Trigger command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Monitor External Trigger command can be executed. </param>
Public Property Let MonitorExtTrigExecutable(ByVal a_value As Boolean)
    If This.MonitorExtTrigExecutable <> a_value Then
        This.MonitorExtTrigExecutable = a_value
        Me.OnPropertyChanged "MonitorExtTrigExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Immediate Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Immediate Trigger Option command  can be executed. </value>
Public Property Get ImmediateTriggerOptionExecutable() As Boolean
    ImmediateTriggerOptionExecutable = This.ImmediateTriggerOptionExecutable
End Property

''' <summary>   Sets the condition indicating that the Immediate Trigger Option command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Immediate Trigger Option command can be executed. </param>
Public Property Let ImmediateTriggerOptionExecutable(ByVal a_value As Boolean)
    If This.ImmediateTriggerOptionExecutable <> a_value Then
        This.ImmediateTriggerOptionExecutable = a_value
        Me.OnPropertyChanged "ImmediateTriggerOptionExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the External Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the External Trigger Option command  can be executed. </value>
Public Property Get ExternalTriggerOptionExecutable() As Boolean
    ExternalTriggerOptionExecutable = This.ExternalTriggerOptionExecutable
End Property

''' <summary>   Sets the condition indicating that the External Trigger Option command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the External Trigger Option command can be executed. </param>
Public Property Let ExternalTriggerOptionExecutable(ByVal a_value As Boolean)
    If This.ExternalTriggerOptionExecutable <> a_value Then
        This.ExternalTriggerOptionExecutable = a_value
        Me.OnPropertyChanged "ExternalTriggerOptionExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the function name command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the function name command can be executed. </value>
Public Property Get FunctionNameExecutable() As Boolean
    FunctionNameExecutable = This.FunctionNameExecutable
End Property

''' <summary>   Sets the function name condition. </summary>
''' <param name="a_value">   [Boolean] True if the function name command can be executed. </param>
Public Property Let FunctionNameExecutable(ByVal a_value As Boolean)
    If This.FunctionNameExecutable <> a_value Then
        This.FunctionNameExecutable = a_value
        Me.OnPropertyChanged "FunctionNameExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns true if the inputs are valid for front or rear settings. </summary>
Public Function ValidateInputs() As Boolean

    ' gets the front input condition: True is Front.
    This.FrontInputsValue = This.K2700.IsFrontInputs
    
    ' check if front or rear
    If This.FrontInputsRequired And Not This.FrontInputsValue Then

        ' display the proper message
        Me.LastMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."

        ValidateInputs = False
    
    ElseIf Not This.FrontInputsRequired And This.FrontInputsValue Then

        ' display the proper message
        Me.LastMessage = "Please depress the front panel INPUTS button to its REAR (R) position and try again."
    
        ValidateInputs = False

    Else

        ValidateInputs = True

    End If

End Function

''' <summary>   Turns on external trigger mode. </summary>
Public Sub ConfigureExternalTriggerReadingCommand()

    Const p_procedureName = "ConfigureExternalTriggerReadingCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' exit if inputs are not in the correct settings

    If Me.ValidateInputs() Then
   
        ' clear execution state when switching mode.

        This.K2700.Device.ClearExecutionState
   
        ' Configure external triggering reading requiring the front panel

        This.K2700.ConfigureExternalTriggerReading This.SenseFunctionName, _
                This.FrontInputsValue

        ' this is done in the previous command.
        ' This.VISession.WriteLine ":TRIG:SOUR IMM"
        ' This.VISession.WriteLine ":TRIG:SOUR EXT"
    
        ' clear the serial poll register
        ' these are also part of the previous command.
        ' This.VISession.WriteLine "*ESE 0"
        ' This.VISession.WriteLine "*SRE 63"
    
        ' clear the measurement event enable register
        ' This.VISession.WriteLine "STAT:PRES"
    
        ' enable the reading available bit of the measurement event register.
        ' This.VISession.WriteLine "STAT:MEAS:ENAB 32"
    
        ' This.Device.QueryServiceRequestStatus
    
        ' move this to starting the monitoring.
        ' initiate awaiting for the external trigger
        ' This.VISession.WriteLine ":INIT; *CLS"
   
        ' disable the immediate reading mode buttons
        Me.MeasureExecutable = False

        ' enable toggling the external trigger monitoring.
        Me.MonitorExtTrigExecutable = True

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
Public Sub ConfigureImmediateTriggerReadingsCommand()

    Const p_procedureName = "ConfigureImmediateTriggerReadingsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' exit if inputs are not in the correct settings

    If Me.ValidateInputs() Then

        ' clear execution state when switching mode.

        This.K2700.Device.ClearExecutionState

        ' configure measuring a single value on each read.
        This.K2700.ConfigureSingleFunctionReading Me.SenseFunctionName
    
        ' Enable making immediate measurements.
        Me.MeasureExecutable = True
    
        ' disable monitoring the external trigger.
        Me.MonitorExtTrigExecutable = False

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsCommand()

    Const p_procedureName = "QueryInputsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatchCommand

    Me.FrontInputsValue = This.K2700.IsFrontInputs

    Me.ReadStopwatchCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Takes a single or multiple readings from the front or rear panel. </summary>
Public Sub MeasureImmediatelyCommand()

    Const p_procedureName = "MeasureImmediatelyCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' reinitialize if device not defined
    If Not This.K2700.Connected Then
        ' This.K2700.InitializeLegacy Me.host, Me.port, Me.ReceiveTimeout
    End If

    ' exit if inputs are not in the correct settings
    If Not Me.ValidateInputs() Then GoTo exit_Handler

    Dim p_channelNumber As Integer

    ' disable the relevant buttons.

    Me.MeasureExecutable = False
    Me.ImmediateTriggerOptionExecutable = False
    Me.ExternalTriggerOptionExecutable = False

    If Me.AutoIncrementChannelNoEnabled Then
    
        ' Clear all cells
        Me.ClearReadingsCommand

        ' scan all channels
        For p_channelNumber = 1 To Me.ChannelCount
    
            ' allow events to occur
            VBA.DoEvents
        
            ' set the current channel number
            Me.CurrentChannelNumber = p_channelNumber
        
            Me.RestartStopwatchCommand
            If Me.FrontInputsValue Then
                This.K2700.TakeFrontReading p_channelNumber
            Else
                This.K2700.TakeRearReading p_channelNumber
            End If
            Me.ReadStopwatchCommand
    
        Next p_channelNumber

    Else

        p_channelNumber = Me.SelectedChannelNumber
    
        Me.RestartStopwatchCommand
        If Me.FrontInputsValue Then
            This.K2700.TakeFrontReading p_channelNumber
        Else
            This.K2700.TakeRearReading p_channelNumber
        End If
    
        Me.ReadStopwatchCommand
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.MeasureExecutable = True
    Me.ImmediateTriggerOptionExecutable = True
    Me.ExternalTriggerOptionExecutable = True

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Clears the readings. </summary>
Public Sub ClearReadingsCommand()

    Dim p_channelNumber As Integer
    For p_channelNumber = 1 To Me.ChannelCount
        VBA.DoEvents
        OnChannelMeasured p_channelNumber, Me.ReadingOffset
    Next p_channelNumber
  
End Sub

''' <summary> Toggles monitoring the instrument for external triggered measurements. </summary>
Public Sub StartMonitoringExternalTriggersCommand()

    Const p_procedureName = "StartMonitoringExternalTriggersCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If Me.AutoIncrementChannelNoEnabled Then
    
        ' Clear all readings
        Me.ClearReadingsCommand

    End If

    Me.MonitoringState = "External Trigger On"

    If Me.AutoIncrementChannelNoEnabled Then
    
        ' set the first channel number for manual scan
        This.CurrentChannelNumber = 1
    
    Else

        ' set the single reading channel number for manual scan
        ' to the channel number entered by the operator.
        This.CurrentChannelNumber = Me.SelectedChannelNumber
    
    End If
   
    ' initiate awaiting for the external trigger
    This.ViSession.WriteLine ":INIT; *CLS"

    ' start the timer
    Me.StartTimer Me.TimerInterval
 
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Public Sub MonitorExtTrigToggleCommand()

    Const p_procedureName = "MonitorExtTrigToggleCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If Me.ExternalTrigMonitoringEnabled Then

        Me.ImmediateTriggerOptionExecutable = False
        Me.ExternalTriggerOptionExecutable = False

        Me.StartMonitoringExternalTriggersCommand

    Else

        ' once monitoring of external triggering is turned off,
        ' the timer is allowed to turn itself off
        Me.ImmediateTriggerOptionExecutable = True
        Me.ExternalTriggerOptionExecutable = True

        ' this is set to require initialization all over again
        Me.MeasureExecutable = False
        Me.MonitorExtTrigExecutable = False

    End If
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_k2700_ChannelMeasured"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' notify the observer of a channel measurements
    If VBA.vbNullString = a_eventArgs.Reading Then
        OnChannelMeasured a_eventArgs.ChannelNumber, -1#
    Else
        OnChannelMeasured a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
    End If

    If Me.AutoIncrementChannelNoEnabled Then

        ' increment the resistance for the next reading
        This.CurrentChannelNumber = a_eventArgs.ChannelNumber + 1
        If This.CurrentChannelNumber > Me.ChannelCount Then
            This.CurrentChannelNumber = 1
        End If

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.K2700.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName = "m_k2700_ConnectionChanged"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If a_eventArgs.Connected Then

        ' report the socket number
        Me.SocketAddress = This.ViSession.SocketAddress
        Me.SocketId = This.ViSession.SocketId
    
        If This.ViSession.UsingGpibLan Then
            Me.ReadAfterWriteEnabled = This.ViSession.GpibLan.ReadAfterWriteEnabledGetter
            If Me.ReadAfterWriteEnabled Then
                ' for the 2700 with needs this off
                This.ViSession.GpibLan.ReadAfterWriteEnabledSetter False
                Me.ReadAfterWriteEnabled = This.ViSession.GpibLan.ReadAfterWriteEnabledGetter
            End If
        End If
    
        This.ViSession.GpibLan.SelectiveDeviceClear
        This.Device.ResetKnownState
        This.Device.ClearExecutionState
    
        Me.TopCard = This.K2700.CardName(1)
        Me.BottomCard = This.K2700.CardName(2)
    
        If VBA.vbNullString <> This.SenseFunctionName Then
        
            ' set the instrument function scan list
            This.K2700.DefineFunctionScanLists This.SenseFunctionName
        
            ' show the results.
            Me.TopCardFunctionScanList = This.K2700.FunctionScanList(1)
            Me.BottomCardFunctionScanList = This.K2700.FunctionScanList(2)
        End If

        ' clear any errors
        This.Device.ClearExecutionState

    Else

        Me.SocketAddress = vbNullString
        Me.SocketId = 0
    
    End If

    Me.ClearReadingsExecutable = True

    Me.ResetKnownStateExecutable = Me.Connected
    Me.ClearExecutionStateExecutable = Me.Connected
    Me.QueryExecutable = Me.Connected
    Me.ReadExecutable = Me.Connected
    Me.WriteExecutable = Me.Connected
    Me.ReadStatusByteExecutable = Me.Connected
    Me.ReadStandardEventStatusExecutable = Me.Connected
    Me.QueryInputsExecutable = Me.Connected
    Me.ImmediateTriggerOptionExecutable = Me.Connected
    Me.ExternalTriggerOptionExecutable = Me.Connected

    ' this gets enabled when entering the immediate mode
    Me.MeasureExecutable = False
    
    ' this gets enabled when entering the external trigger mode
    Me.MonitorExtTrigExecutable = False

    Dim p_GpibLanConnected As Boolean
    p_GpibLanConnected = Me.Connected And This.ViSession.UsingGpibLan
    

    Me.SelectiveDeviceClearExecutable = p_GpibLanConnected
    Me.GoToLocalExecutable = p_GpibLanConnected
    Me.LocalLockoutExecutable = p_GpibLanConnected
    Me.ReadAfterWriteSetExecutable = p_GpibLanConnected
    Me.ReadAfterWriteGetExecutable = p_GpibLanConnected
    Me.SerialPollExecutable = p_GpibLanConnected
    Me.ServiceRequestedExecutable = p_GpibLanConnected
    Me.GpibAddressGetExecutable = p_GpibLanConnected
    Me.GpibAddressSetExecutable = p_GpibLanConnected
    Me.ReadTimeoutGetExecutable = p_GpibLanConnected
    Me.ReadTimeoutSetExecutable = p_GpibLanConnected

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Tcp_Scpi.K2700.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub




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
Public Property Let PauseRequested(ByVal a_value As Boolean)
    If This.PauseRequested <> a_value Then
        This.PauseRequested = a_value
        Me.OnPropertyChanged "PauseRequested"
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
        Me.OnPropertyChanged "TimerInterval"
    End If
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ElapsedTime() As Integer
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Elapsed Time in milliseconds. </value>
Public Property Let ElapsedTime(ByVal a_value As Integer)
    If This.ElapsedTime <> a_value Then
        This.ElapsedTime = a_value
        Me.OnPropertyChanged "ElapsedTime"
        Me.ElapsedTimeCaption = Format(a_value, This.ElapsedTimeFormat)
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
        Me.OnPropertyChanged "ElapsedTimeFormat"
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
        Me.OnPropertyChanged "ElapsedTimeCaption"
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
        Me.OnPropertyChanged "LastErrorMessage"
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
        Me.OnPropertyChanged "LastMessage"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Public Sub RestartStopwatchCommand()

    Me.ElapsedTimeCaption = vbNullString
    This.Stopwatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Public Sub ReadStopwatchCommand()

    Me.ElapsedTime = This.Stopwatch.ElapsedMilliseconds

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ChannelMeasured"/> event. </summary>
''' <para name="a_channelNumber">   [integer] The channel number. </param>
''' <para name="a_reading">         [double] The reading. </param>
Private Sub OnChannelMeasured(ByVal a_channelNumber As Integer, ByVal a_reading As String)
    RaiseEvent ChannelMeasured( _
            cc_isr_Tcp_Scpi.Factory.NewChannelReadingEventArgs.Initialize(a_channelNumber, _
                                                                          a_reading))
End Sub

''' <summary>   Notifies all observers of a property change. </summary>
''' <param name="a_propertyName">   [String] the property naMe. </param>
Public Sub OnPropertyChanged(ByVal a_propertyName As String)

    RaiseEvent PropertyChanged( _
            cc_isr_Core.Factory.CreatePropertyChangedEventArgs(a_propertyName))

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' TIMER CONTROL
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the termination of external trigger monitoring in one place. </summary>
''' <param name="a_infoMessage">   [String] info about the completion state. </param>
Friend Sub OnMonitoringDone(ByVal a_infoMessage As String)

    ' update the monitoring state
    Me.MonitoringState = a_infoMessage

    ' disable the action buttons so as to require
    ' to initialize monitoring over again
    Me.MonitorExtTrigExecutable = False
    Me.MeasureExecutable = False
    
    ' make sure the enable the option buttons to allow initialization
    ' of the two modes.
    Me.ImmediateTriggerOptionExecutable = True
    Me.ExternalTriggerOptionExecutable = True

    ' finally, toggle off the trigger monitoring. This
    ' should be superfluous
    Me.ExternalTrigMonitoringEnabled = False

End Sub

''' <summary>   Handles the external trigger monitoring timer event using the
''' device to handle the timing events. </summary>
''' <remarks>
''' Use this method to check if turn on or off manual triggering mode of the instrument. <para>
''' 10/16/00  David Hary  1.00.00 Add a beep </para>
''' 2023-06-23 David Hary 1.0.8574 Use the device event handler to update the reading values.
''' </remarks>
Private Sub HandleTimerEvent()

    Const p_procedureName = "HandleTimerEvent"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.PauseRequested Or This.HandlingTimerEvent Then Exit Sub
   
    ' prevent reentry
    This.HandlingTimerEvent = True

    Dim p_channelNumber As Integer

    ' check if scan or single
    If Me.AutoIncrementChannelNoEnabled Then
  
        ' in auto increment mode, use the current channel number
        p_channelNumber = This.CurrentChannelNumber
      
    Else

        ' in non auto increment mode, use the current channel number
        p_channelNumber = Me.SelectedChannelNumber

    End If
  
    ' allow other events to occur
    VBA.DoEvents

    Dim p_isDone As Boolean: p_isDone = Not Me.ExternalTrigMonitoringEnabled

    ' stop the timer.
    If p_isDone Then Me.PauseTimer

    This.K2700.MonitorExternalTrigger p_channelNumber, p_isDone

    ' if we ended manual mode, enable the rest of the buttons
    If p_isDone Then

        Me.OnMonitoringDone "Trigger monitoring done"
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release reentry prevention
    This.HandlingTimerEvent = False

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' stop the timer.
    Me.PauseTimer

    ' handle the termination of monitoring
    Me.OnMonitoringDone "Error handling the timer event"

    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Starts the event timer. </summary>
''' <param name="a_interval">   [Long] Specifies the timer interval in milliseconds. </param>
Public Sub StartTimer(Optional ByVal a_interval As Long = 100)

    This.HandlingTimerEvent = False
    This.PauseRequested = False
    If m_eventTimer Is Nothing Then
        Set m_eventTimer = cc_isr_Core.Factory.NewEventTimer.Initialize(a_interval)
        m_eventTimer.StartTimer
    Else
        m_eventTimer.ResumeTimer
    End If

End Sub

''' <summary>   Requests pausing the timer. </summary>
Public Sub PauseTimer()

    This.PauseRequested = True

End Sub

''' <summary>   Aborts the timer. </summary>
Public Sub AbortTimer()

    If Not m_eventTimer Is Nothing Then
      m_eventTimer.StopTimer
      Set m_eventTimer = Nothing
    End If

End Sub

''' <summary>   Handles a timer tick event. </summary>
Private Sub m_eventTimer_Tick()

    On Error Resume Next
    If This.PauseRequested Then
        If Not m_eventTimer Is Nothing Then m_eventTimer.PauseTimer
    Else
        HandleTimerEvent
    End If
    On Error GoTo 0

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the <see cref="cc_isr_Winsock.IConnectable"/>
''' <see cref="cc_isr_Ieee488.Ieee488Session"/> object . </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    Set Connectable = This.ViSession
End Property

''' <summary>   Returns the connection state of the client. </summary>
''' <value>   [Boolean] True if the IEEE488 Session is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean

    Connected = IConnectable_Connected

End Property

''' <summary>   Returns the connection state of the <see cref="cc_isr_Winsock.IConnectable"/>. </summary>
''' <value>   [Boolean] True if the <see cref="cc_isr_Winsock.IConnectable"/> is connected; otherwise, False. </value>
Private Property Get IConnectable_Connected() As Boolean

    If Me.Connectable Is Nothing Then
        IConnectable_Connected = False
    Else
        IConnectable_Connected = Me.Connectable.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean

    CanConnect = IConnectable_CanConnect

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Private Property Get IConnectable_CanConnect() As Boolean

    If Me.Connectable Is Nothing Then
        IConnectable_CanConnect = False
    Else
        IConnectable_CanConnect = Me.Connectable.CanConnect
    End If

End Property

''' <summary>   Closes the connection and releases the <see cref="Ieee488Session">Session</see>. </summary>
Public Sub Dispose()

    IConnectable_Dispose

End Sub

''' <summary>   Closes the connection and releases the <see cref="cc_isr_Winsock.IConnectable">Session</see>. </summary>
Private Sub IConnectable_Dispose()

    Const p_procedureName As String = "IConnectable_Dispose"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Set This.Stopwatch = Nothing

    If Not m_eventTimer Is Nothing Then

        m_eventTimer.StopTimer
  
        ' allow the timer to stop.
        cc_isr_Core_IO.CoreExtensions.Wait 0.0015 * Me.TimerInterval
    
        Set m_eventTimer = Nothing
    
    End If

    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If

    If Not m_k2700 Is Nothing Then m_k2700.Dispose
    Set m_k2700 = Nothing
    Set This.K2700 = Nothing

    Set This.Device = Nothing
    Set This.ViSession = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Opens a VI Session connection to the specified host at the specified port
''' as sets the connection timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name=a_"timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the VI Session is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500) As Boolean

    IConnectable_OpenConnection a_host, a_port, a_timeoutMs

End Function

''' <summary>   Opens an <see cref="cc_isr_Winsock.IConnectable"/> connection to the
''' specified host at the specified port as sets the socket receive timeout. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the <see cref="Connectable"/> is connected. </returns>
Private Function IConnectable_OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500&) As Boolean

    Dim p_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_eventArgs = cc_isr_Winsock.Factory.NewConnectionChangingEventArgs.Initialize(Me.Connected)
    IConnectable_OnConnectionChanging p_eventArgs

    If Not p_eventArgs.Cancel Then

        Connectable.OpenConnection a_host, a_port, a_timeoutMs
    
        If Me.Connected Then
    
            If This.ViSession.UsingGpibLan Then
                If This.ViSession.GpibLan.ReadAfterWriteEnabledGetter Then
                    ' for the 2700 with needs this off
                    This.ViSession.GpibLan.ReadAfterWriteEnabledSetter False
                End If
            End If
        
        End If
    
        OnConnectionChanged Connectable.Connected
    
    End If

    IConnectable_OpenConnection = Connectable.Connected

End Function

''' <summary>   Closes the <see cref="cc_isr_Ieee488.Ieee488Session">IEEE488 Session</see> connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    CloseConnection = IConnectable_CloseConnection

End Function

''' <summary>   Closes the <see cref="cc_isr_Winsock.IConnectable"/> connection. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Private Function IConnectable_CloseConnection() As Boolean

    Dim p_args As cc_isr_Winsock.ConnectionChangingEventArgs
    Set p_args = cc_isr_Winsock.Factory.NewConnectionChangingEventArgs.Initialize(Me.Connected)
    IConnectable_OnConnectionChanging p_args
   
    If Not p_args.Cancel Then

        IConnectable_CloseConnection = Connectable.CloseConnection

        OnConnectionChanged Connectable.Connected

    End If

    IConnectable_CloseConnection = Not Connectable.Connected

End Function

''' <summary>   Raises the <see cref="ConnectionChanged"/> event after the
''' connection was established and ended. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)

    IConnectable_OnConnectionChanged cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize(a_isConnected)
    OnConnectionChanged "Connected"

End Sub

''' <summary>   Raises the <see cref="ConnectionChanged"/> event after the
''' connection was established and ended. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    RaiseEvent ConnectionChanged(a_eventArgs)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    RaiseEvent ConnectionChanging(a_eventArgs)

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Class events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub Class_Initialize()
    This.ElapsedTimeFormat = "0.0"
End Sub
