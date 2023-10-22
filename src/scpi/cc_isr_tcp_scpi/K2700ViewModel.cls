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

''' <summary>   Values that represent mesaurement mode options. </summary>
Public Enum MeasurementModeOption
    ''' <summary>   An enum constant representing the none option. </summary>
    None = 0
    ''' <summary>   An enum constant representing the continuous trigger option. </summary>
    Continuous = 1
    ''' <summary>   An enum constant representing the immediate trigger option. </summary>
    Immediate = 2
    ''' <summary>   An enum constant representing the external trigger option. </summary>
    External = 3
    ''' <summary>   An enum constant representing the external trigger monitoring option. </summary>
    Monitoring = 4
End Enum

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a property value changes. </summary>
Implements cc_isr_Core.INotifyPropertyChanged

''' <summary>   Notifies all registered handlers of a property value change. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Core.PropertyChangedEventArgs"/> event args. </param>
Public Event PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

''' <summary>   The event that is raised upon completion of a channel measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/> event args. </param>
Public Event ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

''' <summary>   Notifies all registered handlers of a measurement state change. </summary>
''' <remarks>   The measurment state change is notified once all elements, such as measuring, measurement mode,
''' and single reading are set and the measurement process is effectively on or off. </remarks>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Tcp_Scpi.StateChangedEventArgs"/> event args. </param>
Public Event MeasurementStateChanged(ByVal a_eventArgs As cc_isr_Tcp_Scpi.StateChangedEventArgs)

''' <summary>   Notifies all registered handlers that a measurement state is about to change. </summary>
''' <remarks>   This method is used to disable controls and control value changes while the new measurement
''' state is established. The control states are restored once the state change is completed. </remarks>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Core.EventArgs"/> event args. </param>
Public Event MeasurementStateChanging(ByVal a_eventArgs As cc_isr_Core.EventArgs)

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a measurements are completed. </summary>
Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

''' <summary>   Represents an connectable pattern subject that can notify observers
'''             when a connection is changing or changed. </summary>
Private WithEvents m_socket As cc_isr_Winsock.IPv4StreamSocket
Attribute m_socket.VB_VarHelpID = -1

Private Type this_

    Stopper As cc_isr_Core_IO.Stopwatch

    ' view model properties
    ElapsedTime As Double
    ElapsedTimeFormat As String
    ElapsedTimeCaption As String
    LastErrorMessage As String
    LastMessage As String
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    Measuring As Boolean
    MaximumChannelNumber As Integer
    DesignMode As Boolean

    ' Winsock properties
    Host As String
    Port As Long
    SocketAddress As String
    SocketId As Integer
    SessionTimeout As Integer
    HasValidSocketAddress As Boolean

    ' session settings
    Session As cc_isr_Ieee488.TcpSession
    GpibLanControllerPort As Long
    ReadAfterWriteDelay As Long
    Termination As String

    ' ieee488 device properties
    Device As cc_isr_Ieee488.Device

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
    AssertTalkAfterWriteAllowed As Boolean
    AutoAssertTalk As Boolean
    GpibLanReadTimeout As Integer
    
    ' 2700 properties
    K2700 As K2700
    TopCard As String
    BottomCard As String
    TopCardFunctionScanList As String
    BottomCardFunctionScanList As String
    ReadingBeepEnabled As Boolean
    FrontInputsRequired As Boolean
    FrontInputsValue As Boolean
    FrontInputsHasValue As Boolean
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
    ReadingOffset As Double
    
    ' service request monitoring timer control
    ExternalTrigMonitoringEnabled As Boolean
    StopRequested As Boolean
    PauseRequested As Boolean
    TimerInterval As Integer
    ' true when handling the timer event to prevent timer
    ' event handler overrun. This effectively synchronizes the timer
    ' event with the instrument.
    HandlingTimerEvent As Boolean
    ' indicates that the timer started; used to monitor timer state.
    TimerStarted As Boolean
    TimerControlled As Boolean

    ' Service request monitoring properties
    MonitoringState As String
    AutoIncrementChannelNoEnabled As Boolean
    SingleReadEnabled As Boolean

    ' Winsock connection executables
    OpenConnectionExecutable As Boolean
    CloseConnectionExecutable As Boolean
    
    ' IEEE488 executables
    ClearExecutionStateExecutable As Boolean
    MessageToSendExecutable As Boolean
    QueryExecutable As Boolean
    ReadExecutable As Boolean
    QueryStatusByteExecutable As Boolean
    ReadStatusByteExecutable As Boolean
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
    ReadStatisByteExecutable As Boolean
    GpibAddressGetExecutable As Boolean
    GpibAddressSetExecutable As Boolean
    GpibLanReadTimeoutGetExecutable As Boolean
    GpibLanReadTimeoutSetExecutable As Boolean

End Type

Private This As this_

''' <summary>   Initialize this worksheet. </summary>
''' <returns>   [<see cref="K2700ViewModel"/>]. </returns>
Public Function Initialize() As K2700ViewModel

    ' clear the last error message
    Me.LastErrorMessage = VBA.vbNullString

    ' populate the combo box commands
    PopulateScpiCommands

    ' populate the function box commands
    PopulateSenseFunctions

    ' K2700 default settings
    Me.ReadAfterWriteDelay = 1
    Me.Termination = VBA.vbLf
    
    This.K2700.GpibLanControllerPort = This.GpibLanControllerPort
    This.K2700.ReadAfterWriteDelay = This.ReadAfterWriteDelay
    This.K2700.Termination = This.Termination
    
    ' initializing the virtual instrument initializes the device and the socket.
    This.K2700.Initialize
    Set m_k2700 = This.K2700
    Set m_socket = This.Session.Socket
    
    Me.MeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous
    
    ' update the executable settings
    Me.InitializeExecutablesState

    ' enable connection if all values are set
    Me.ConnectionExecutableSetter
    
    Set Initialize = Me

End Function

''' <summary>   Closes the connection and releases the <see cref="Ieee488Session">Session</see>. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' clear the last error message
    Me.LastErrorMessage = VBA.vbNullString

    Set This.Stopper = Nothing

    If Not m_eventTimer Is Nothing Then

        m_eventTimer.StopTimer
  
        ' allow the timer to stop.
        cc_isr_Core_IO.CoreExtensions.Wait 0.0015 * This.TimerInterval
    
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

    Set m_k2700 = Nothing
    If Not This.K2700 Is Nothing Then This.K2700.Dispose
    Set This.K2700 = Nothing

    Set This.Device = Nothing
    Set This.Session = Nothing
    Set m_socket = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' User interface settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the has valid socket address condition indicating that the socket address
''' was parsed and has the expected port numbers. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasValidSocketAddress() As Boolean
    HasValidSocketAddress = This.HasValidSocketAddress
End Property

''' <summary>   Parses the current socket address and sets the host and port values. </summary>
''' <param name="a_value">   [String] the socket address. </param>
Friend Sub ParseSocketAddress(ByVal a_Value As String)
    This.HasValidSocketAddress = False
    This.SocketAddress = a_Value
    This.Host = VBA.vbNullString
    This.Port = cc_isr_Winsock.ws32_INVALID_SOCKET
    If cc_isr_Winsock.Winsock.TryParseAddress(This.SocketAddress, This.Host, This.Port, Me.LastErrorMessage) Then _
        This.HasValidSocketAddress = VBA.vbNullString <> This.Host _
            And (5025 = This.Port Or 1234 = This.Port)
    Me.OnPropertyChanged "Port"
    Me.OnPropertyChanged "Host"
    Me.OnPropertyChanged "SocketAddress"
End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Long
    Port = This.Port
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_Value As String)
    If This.SocketAddress <> a_Value Then
        Me.ParseSocketAddress (a_Value)
        Me.ConnectionExecutableSetter
    End If
End Property

''' <summary>   Gets the socket id. </summary>
''' <value>   [Integer]. </value>
Public Property Get SocketId() As Integer
    SocketId = This.SocketId
End Property

''' <summary>   Sets the socket id. </summary>
''' <param name="a_value">   [Integer] the socket id. </value>
Friend Property Let SocketId(ByVal a_Value As Integer)
    If This.SocketId <> a_Value Then
        This.SocketId = a_Value
        Me.OnPropertyChanged "SocketId"
    End If
End Property

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the receive timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_Value As Integer)
    If This.SessionTimeout <> a_Value Then
        This.SessionTimeout = a_Value
        Me.OnPropertyChanged "SessionTimeout"
        ' connection executable require a positive timeout
        Me.ConnectionExecutableSetter
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
Friend Property Let OpenConnectionExecutable(ByVal a_Value As Boolean)
    If This.OpenConnectionExecutable <> a_Value Then
        This.OpenConnectionExecutable = a_Value
        Me.OnPropertyChanged "OpenConnectionExecutable"
        Me.OnPropertyChanged "ToggleConnectionExecutable"
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
Friend Property Let CloseConnectionExecutable(ByVal a_Value As Boolean)
    If This.CloseConnectionExecutable <> a_Value Then
        This.CloseConnectionExecutable = a_Value
        Me.OnPropertyChanged "CloseConnectionExecutable"
        Me.OnPropertyChanged "ToggleConnectionExecutable"
    End If
End Property

''' <summary>   Sets the condition indicating that a Connection command
'''             can be executed. </summary>
Friend Sub ConnectionExecutableSetter()
    Dim p_canExecute As Boolean
    p_canExecute = This.HasValidSocketAddress And 0 < This.SessionTimeout
    Me.OpenConnectionExecutable = p_canExecute And Not Me.Connected
    Me.CloseConnectionExecutable = p_canExecute And Me.Connected
    
    ' notify of connection state.
    Me.OnPropertyChanged "Connected"
    Me.OnPropertyChanged "CanConnect"
    
End Sub

''' <summary>   Gets the condition indicating that the toggle Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Connection command can be executed. </value>
Public Property Get ToggleConnectionExecutable() As Boolean
    ToggleConnectionExecutable = This.OpenConnectionExecutable Or This.CloseConnectionExecutable
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Winsock Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Tries to issue power on reset. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <param name="a_delay">     [Double] the post seset delay. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryPing(ByRef a_details As String, ByVal a_delay As Double) As Boolean

    Const p_procedureName = "TryPowerOnReset"

    Dim p_details As String
    Dim p_success As Boolean
    p_success = cc_isr_Ieee488.GpibLanController.TryPing(This.SocketAddress, This.SessionTimeout, p_details)
    
    If p_success Then
        p_success = 0 = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
        If Not p_success Then
            p_details = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        End If
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & p_details

    TryPing = p_success
    
End Function

''' <summary>   Tries to issue power on reset. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <param name="a_delay">     [Double] the post seset delay. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryPowerOnReset(ByRef a_details As String, ByVal a_delay As Double) As Boolean

    Const p_procedureName = "TryPowerOnReset"

    Dim p_details As String
    Dim p_success As Boolean
    p_success = This.Session.TryPowerOnReset(This.SocketAddress, p_details, a_delay)
    
    If p_success Then
        p_success = 0 = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
        If Not p_success Then
            p_details = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        End If
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & p_details

    TryPowerOnReset = p_success
    
End Function

''' <summary>   Handles the power on reset command event. </summary>
Public Sub PowerOnResetCommand()

    Const p_procedureName = "PowerOnResetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Me.LastMessage = "Resetting the controller to its power on state; please wait..."
    
    Dim p_delay As Double: p_delay = 3
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean
    p_success = Me.TryPowerOnReset(p_details, p_delay)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed power on reset; " & p_details
    
    If p_success Then _
        Me.LastMessage = "The controller was reset to its power on state."
    
    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Opens the connection and power on reset the GPIB Lan Controller. </summary>
''' <param name="a_delay">     [Double] the post seset delay. </param>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryOpenConnectionPowerOnReset(ByRef a_details As String, ByVal a_delay As Double) As Boolean

    Const p_procedureName = "OpenConnectionPowerOnReset"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' initialize all elements.
    Me.Initialize
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_success Then _
        p_success = Me.TryPowerOnReset(p_details, a_delay)
    
    If p_success Then _
        p_success = Me.Connectable.TryOpenConnection(This.SocketAddress, This.SessionTimeout, p_details)
    
    If Not p_success Then _
        p_details = "Failed opening the connection after power on reset; " & p_details

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & p_details
    
    TryOpenConnectionPowerOnReset = p_success
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    p_success = False
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' save the error message
    p_details = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Handles the open connection power on reset command event. </summary>
Public Sub OpenConnectionPowerOnResetCommand()

    Const p_procedureName = "OpenConnectionPowerOnResetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Dim p_delay As Double: p_delay = 3
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean
    p_success = Me.TryOpenConnectionPowerOnReset(p_details, p_delay)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed opening the connection; " & p_details
    
    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Opens the connection. </summary>
Public Sub OpenConnection()

    Const p_procedureName = "OpenConnection"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' initialize all elements.
    Me.Initialize
    
    ' disable the connectibles
    Me.OpenConnectionExecutable = False
    Me.CloseConnectionExecutable = False
    
    Me.LastMessage = "Pinging controller and device at " & This.SocketAddress & "; please wait..."
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True
    p_success = cc_isr_Ieee488.GpibLanController.TryPing(This.SocketAddress, This.SessionTimeout, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed pinging the instrument; " & p_details
    
    Me.LastMessage = "Opening connection to " & This.SocketAddress & "; please wait..."
    
    p_success = Me.Connectable.TryOpenConnection(This.SocketAddress, This.SessionTimeout, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed opening the connection; " & p_details

    If p_success Then _
        Me.LastMessage = This.SocketAddress & " is open."
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' ensure that either connection or disconnections are allowed
    Me.ConnectionExecutableSetter
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Handles the open connection command event. </summary>
Public Sub OpenConnectionCommand()

    Const p_procedureName = "OpenConnectionCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Me.OpenConnection
    
    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Closes the connection. </summary>
Public Sub CloseConnection()

    Const p_procedureName = "CloseConnection"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' disable the connectibles
    Me.OpenConnectionExecutable = False
    Me.CloseConnectionExecutable = False

    ' clear the last error message
    Me.LastErrorMessage = VBA.vbNullString

    ' turn off the event timer.
    If Not m_eventTimer Is Nothing Then

        m_eventTimer.StopTimer
  
        ' allow the timer to stop.
        cc_isr_Core_IO.CoreExtensions.Wait 0.0015 * This.TimerInterval
    
        Set m_eventTimer = Nothing
    
    End If

    Dim p_details As String
    If Me.Connected Then
        
        Me.LastMessage = "Closing " & This.SocketAddress & "; please wait..."
        
        If Me.Connectable.TryCloseConnection(p_details) Then
            Me.LastMessage = "Done Closing " & This.SocketAddress & "."
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed closing the connection; " & p_details
        End If
    End If
        
    ' turn of instrument event handling
    Set m_k2700 = Nothing

    ' turn of socket event handling.
    Set m_socket = Nothing
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:
    
    ' ensure that either connection or disconnections are allowed
    Me.ConnectionExecutableSetter

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Handles the close connection command event. </summary>
Public Sub CloseConnectionCommand()

    Const p_procedureName = "CloseConnectionCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.CloseConnection
    
    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Tcp Session Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.TcpSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.TcpSession"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.TcpSession
    Set Session = This.Session
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_Value As Long)
    If This.GpibLanControllerPort <> a_Value Then
        This.GpibLanControllerPort = a_Value
        Me.OnPropertyChanged "GpibLanControllerPort"
    End If
End Property

''' <summary>   Gets the Read After Write Delay in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = This.ReadAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Read After Write Delay Delay in milliseconds. </param>
Public Property Let ReadAfterWriteDelay(ByVal a_Value As Integer)
    If This.ReadAfterWriteDelay <> a_Value Then
        This.ReadAfterWriteDelay = a_Value
        Me.OnPropertyChanged "ReadAfterWriteDelay"
    End If
End Property

''' <summary>   Gets the termination to append to output messages. </summary>
''' <value>   [String]. </value>
Public Property Get Termination() As String
    Termination = This.Termination
End Property

''' <summary>   Sets the termination to append to output messages. </summary>
''' <param name="a_value">   [String] the termination. </param>
Public Property Let Termination(ByVal a_Value As String)
    If This.Termination <> a_Value Then
        This.Termination = a_Value
        Me.OnPropertyChanged "Termination"
    End If
End Property

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

''' <summary>   Gets the Message to Sent. </summary>
''' <value>   [String]. </value>
Public Property Get MessageToSend() As String
    MessageToSend = This.MessageToSend
End Property

''' <summary>   Sets the Message to Sent. </summary>
''' <param name="a_value">   [String] the Message to Sent. </value>
Public Property Let MessageToSend(ByVal a_Value As String)
    If This.MessageToSend <> a_Value Then
        This.MessageToSend = a_Value
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
Public Property Let SentMessage(ByVal a_Value As String)
    If This.SentMessage <> a_Value Then
        This.SentMessage = a_Value
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
Public Property Let ReceivedMessage(ByVal a_Value As String)
    If This.ReceivedMessage <> a_Value Then
        This.ReceivedMessage = a_Value
        Me.OnPropertyChanged "ReceivedMessage"
        Me.ReceivedMessageLength = VBA.Len(a_Value)
    End If
End Property

''' <summary>   Gets the Received Message Length. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceivedMessageLength() As Integer
    ReceivedMessageLength = This.ReceivedMessageLength
End Property

''' <summary>   Sets the Received Message Length. </summary>
''' <param name="a_value">   [Integer] the Message Length value. </value>
Public Property Let ReceivedMessageLength(ByVal a_Value As Integer)
    If This.ReceivedMessageLength <> a_Value Then
        This.ReceivedMessageLength = a_Value
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
Public Property Let StatusByte(ByVal a_Value As Integer)
    If This.StatusByte <> a_Value Then
        This.StatusByte = a_Value
        Me.OnPropertyChanged "StatusByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
                a_Value And cc_isr_Ieee488.ServiceRequestFlags.RequestingService
    End If
End Property

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    StandardByte = This.StandardByte
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Public Property Let StandardByte(ByVal a_Value As Integer)
    If This.StandardByte <> a_Value Then
        This.StandardByte = a_Value
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
Public Property Let ReadStatusAfterWriteEnabled(ByVal a_Value As Boolean)
    If This.ReadStatusAfterWriteEnabled <> a_Value Then
        This.ReadStatusAfterWriteEnabled = a_Value
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
Public Property Let ScpiCommands(ByVal a_Value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.ScpiCommands, a_Value) Then
        Set This.ScpiCommands = a_Value
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
    p_scpiCommands.Add "SYST:BEEP ON"
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
Public Property Let ResetKnownStateExecutable(ByVal a_Value As Boolean)
    If This.ResetKnownStateExecutable <> a_Value Then
        This.ResetKnownStateExecutable = a_Value
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
Public Property Let ClearExecutionStateExecutable(ByVal a_Value As Boolean)
    If This.ClearExecutionStateExecutable <> a_Value Then
        This.ClearExecutionStateExecutable = a_Value
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
Public Property Let QueryExecutable(ByVal a_Value As Boolean)
    If This.QueryExecutable <> a_Value Then
        This.QueryExecutable = a_Value
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
Public Property Let ReadExecutable(ByVal a_Value As Boolean)
    If This.ReadExecutable <> a_Value Then
        This.ReadExecutable = a_Value
        Me.OnPropertyChanged "ReadExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Query Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Status Byte command can be executed. </value>
Public Property Get QueryStatusByteExecutable() As Boolean
    QueryStatusByteExecutable = This.QueryStatusByteExecutable
End Property

''' <summary>   Sets the condition indicating that the Query Status Byte command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query Status Byte command can be executed. </param>
Public Property Let QueryStatusByteExecutable(ByVal a_Value As Boolean)
    If This.QueryStatusByteExecutable <> a_Value Then
        This.QueryStatusByteExecutable = a_Value
        Me.OnPropertyChanged "QueryStatusByteExecutable"
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
Public Property Let ReadStatusByteExecutable(ByVal a_Value As Boolean)
    If This.ReadStatusByteExecutable <> a_Value Then
        This.ReadStatusByteExecutable = a_Value
        Me.OnPropertyChanged "ReadStatusByteExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Status After Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status After Write command can be executed. </value>
Public Property Get ReadStatusAfterWriteExecutable() As Boolean
    ReadStatusAfterWriteExecutable = This.ReadStatusAfterWriteExecutable
End Property

''' <summary>   Sets the condition indicating that the Read Status After Write command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read Status After Write command can be executed. </param>
Public Property Let ReadStatusAfterWriteExecutable(ByVal a_Value As Boolean)
    If This.ReadStatusAfterWriteExecutable <> a_Value Then
        This.ReadStatusAfterWriteExecutable = a_Value
        Me.OnPropertyChanged "ReadStatusAfterWriteExecutable"
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
Public Property Let ReadStandardEventStatusExecutable(ByVal a_Value As Boolean)
    If This.ReadStandardEventStatusExecutable <> a_Value Then
        This.ReadStandardEventStatusExecutable = a_Value
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
Public Property Let WriteExecutable(ByVal a_Value As Boolean)
    If This.WriteExecutable <> a_Value Then
        This.WriteExecutable = a_Value
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
Public Property Let MessageToSendExecutable(ByVal a_Value As Boolean)
    If This.MessageToSendExecutable <> a_Value Then
        This.MessageToSendExecutable = a_Value
        Me.OnPropertyChanged "MessageToSendExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends the instrument to is Clear Execution State and returns
''' success or failure. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryClearExecutionState(ByVal a_details As String) As Boolean

    Const p_procedureName = "TryClearExecutionState"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    Me.LastErrorMessage = VBA.vbNullString

    ' clear execution state and check the operation completion reply.
    p_success = This.Device.TryClearExecutionState(p_details)
    If p_success Then
        a_details = VBA.vbNullString
    Else
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed clearing the IEEE488 device execution state; " & p_details
    End If

    TryClearExecutionState = p_success
    
End Function

''' <summary>   Handles the Clear Execution State command event. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function ClearExecutionStateCommand(ByVal a_details As String) As Boolean

    Const p_procedureName = "ClearExecutionStateCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString

    Me.RestartStopper

    p_success = Me.TryClearExecutionState(p_details)
    
    If Not p_success Then
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "; " & p_details
    End If

    Me.ReadStopper
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ClearExecutionStateCommand = p_success
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' for unit testing.
    a_details = Me.LastErrorMessage
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Writes the <see cref="MessageToSend"/> and update elapsed times. </summary>
Public Sub WriteCommand()

    Const p_procedureName = "WriteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.WriteLine This.MessageToSend

    Me.ReadStopper

    Me.SentMessage = This.MessageToSend

    If This.ReadStatusAfterWriteEnabled Then

        Dim p_canQuery As Boolean
        p_canQuery = Not cc_isr_Ieee488.Scpi.IsQuery(This.MessageToSend)
        Me.ReadStatusByte p_canQuery

        If p_canQuery Then
            Me.QueryStatusByteCommand
            Me.ReadStandardEventStatusCommand
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

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
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

    Me.RestartStopper

    Me.ReceivedMessage = This.Session.QueryLine(Me.MessageToSend)

    Me.ReadStopper

    Me.SentMessage = This.MessageToSend
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
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

    Me.RestartStopper

    Me.ReceivedMessage = This.Session.Read()

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte command event. </summary>
Public Sub ReadStandardEventStatusCommand()

    Const p_procedureName = "ReadStandardEventStatusCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.StandardByte = This.Device.QueryStandardEventsStatus

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Sub

''' <summary>   Sends the instrument to is reset known State and returns
''' success or failure. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryResetKnownState(ByVal a_details As String) As Boolean

    Const p_procedureName = "TryResetKnownState"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    Me.LastErrorMessage = VBA.vbNullString

    ' clear execution state and check the operation completion reply.
    p_success = This.Device.TryResetKnownState(p_details)
    If p_success Then
        a_details = VBA.vbNullString
    Else
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed reseting the IEEE488 device to its known state; " & p_details
    End If

    TryResetKnownState = p_success
    
End Function

''' <summary>   Handles the Reset Known State command event. </summary>
''' <param name="a_details">   [String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function ResetKnownStateCommand(ByVal a_details As String) As Boolean

    Const p_procedureName = "ResetKnownStateCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString

    Me.RestartStopper

    p_success = Me.TryResetKnownState(p_details)
    
    If Not p_success Then
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "; " & p_details
    End If

    Me.ReadStopper

    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ResetKnownStateCommand = p_success
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' for unit testing.
    a_details = Me.LastErrorMessage
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Tries to read (<c>*STB?</c>) the device status byte. </summary>
''' <remarks>   This method updates the <see cref="K2700ViewModel.StatusByte"/>. </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean] ture if success. </returns>
Public Function TryQueryStatusByte(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryQueryStatusByte"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_statusByte As Integer
    p_statusByte = This.Device.TryQueryServiceRequestStatus(p_details)
    If 0 > p_statusByte Then
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                p_details
    Else
        Me.StatusByte = p_statusByte
    End If
    
    TryQueryStatusByte = 0 <= p_statusByte

End Function


''' <summary>   Handles the Query Status Byte command event and sets the status byte. </summary>
''' <remarks>   This method updates the <see cref="K2700ViewModel.StatusByte"/>.
''' Uses <c>*STB?</c> </remarks>
Public Sub QueryStatusByteCommand()

    Const p_procedureName = "QueryStatusByteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Dim p_details As String: p_details = VBA.vbNullString
    If Not Me.TryQueryStatusByte(p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "; " & p_details

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
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
Public Property Let SerialPollByte(ByVal a_Value As Integer)
    If This.SerialPollByte <> a_Value Then
        This.SerialPollByte = a_Value
        Me.OnPropertyChanged "SerialPollByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
                a_Value And cc_isr_Ieee488.ServiceRequestFlags.RequestingService
    End If
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_Value As Integer)
    If This.ServiceRequested <> a_Value Then
        This.ServiceRequested = a_Value
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
Public Property Let GpibAddress(ByVal a_Value As Integer)
    If This.GpibAddress <> a_Value Then
        This.GpibAddress = a_Value
        Me.OnPropertyChanged "GpibAddress"
    End If
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let AutoAssertTalk(ByVal a_Value As Boolean)
    If This.AutoAssertTalk <> a_Value Then
        This.AutoAssertTalk = a_Value
        Me.OnPropertyChanged "AutoAssertTalk"
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
Public Property Let GpibLanReadTimeout(ByVal a_Value As Integer)
    If This.GpibLanReadTimeout <> a_Value Then
        This.GpibLanReadTimeout = a_Value
        Me.OnPropertyChanged "GpibLanReadTimeout"
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
Public Property Let SelectiveDeviceClearExecutable(ByVal a_Value As Boolean)
    If This.SelectiveDeviceClearExecutable <> a_Value Then
        This.SelectiveDeviceClearExecutable = a_Value
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
Public Property Let GoToLocalExecutable(ByVal a_Value As Boolean)
    If This.GoToLocalExecutable <> a_Value Then
        This.GoToLocalExecutable = a_Value
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
Public Property Let LocalLockoutExecutable(ByVal a_Value As Boolean)
    If This.LocalLockoutExecutable <> a_Value Then
        This.LocalLockoutExecutable = a_Value
        Me.OnPropertyChanged "LocalLockoutExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the auto assert TALK get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the auto assert TALK get command can be executed. </value>
Public Property Get AutoAssertTalkGetExecutable() As Boolean
    AutoAssertTalkGetExecutable = This.AutoAssertTalkGetExecutable
End Property

''' <summary>   Gets the condition indicating that the auto assert TALK get command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the auto assert TALK get command can be executed. </param>
Public Property Let AutoAssertTalkGetExecutable(ByVal a_Value As Boolean)
    If This.AutoAssertTalkGetExecutable <> a_Value Then
        This.AutoAssertTalkGetExecutable = a_Value
        Me.OnPropertyChanged "AutoAssertTalkGetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the auto assert TALK set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the auto assert TALK set command can be executed. </value>
Public Property Get AutoAssertTalkSetExecutable() As Boolean
    AutoAssertTalkSetExecutable = This.AutoAssertTalkSetExecutable
End Property

''' <summary>   Sets the condition indicating that the auto assert TALK set command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the auto assert TALK set command can be executed. </param>
Public Property Let AutoAssertTalkSetExecutable(ByVal a_Value As Boolean)
    If This.AutoAssertTalkSetExecutable <> a_Value Then
        This.AutoAssertTalkSetExecutable = a_Value
        Me.OnPropertyChanged "AutoAssertTalkSetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Status Byte (Serial poll) command
''' can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get ReadStatisByteExecutable() As Boolean
    ReadStatisByteExecutable = This.ReadStatisByteExecutable
End Property

''' <summary>   Sets the condition indicating that the read status byte (Serial Poll) command
''' can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Serial Poll command can be executed. </param>
Public Property Let ReadStatisByteExecutable(ByVal a_Value As Boolean)
    If This.ReadStatusByteExecutable <> a_Value Then
        This.ReadStatusByteExecutable = a_Value
        Me.OnPropertyChanged "ReadStatusByteExecutable"
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
Public Property Let ServiceRequestedExecutable(ByVal a_Value As Boolean)
    If This.ServiceRequestedExecutable <> a_Value Then
        This.ServiceRequestedExecutable = a_Value
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
Public Property Let GpibAddressGetExecutable(ByVal a_Value As Boolean)
    If This.GpibAddressGetExecutable <> a_Value Then
        This.GpibAddressGetExecutable = a_Value
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
Public Property Let GpibAddressSetExecutable(ByVal a_Value As Boolean)
    If This.GpibAddressSetExecutable <> a_Value Then
        This.GpibAddressSetExecutable = a_Value
        Me.OnPropertyChanged "GpibAddressSetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Get command can be executed. </value>
Public Property Get GpibLanReadTimeoutGetExecutable() As Boolean
    GpibLanReadTimeoutGetExecutable = This.GpibLanReadTimeoutGetExecutable
End Property

''' <summary>   Gets the Read Timeout Get command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Get command can be executed. </param>
Public Property Let GpibLanReadTimeoutGetExecutable(ByVal a_Value As Boolean)
    If This.GpibLanReadTimeoutGetExecutable <> a_Value Then
        This.GpibLanReadTimeoutGetExecutable = a_Value
        Me.OnPropertyChanged "GpibLanReadTimeoutGetExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Read Timeout Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Set command can be executed. </value>
Public Property Get GpibLanReadTimeoutSetExecutable() As Boolean
    GpibLanReadTimeoutSetExecutable = This.GpibLanReadTimeoutSetExecutable
End Property

''' <summary>   Sets the Read Timeout Set command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Set command can be executed. </param>
Public Property Let GpibLanReadTimeoutSetExecutable(ByVal a_Value As Boolean)
    If This.GpibLanReadTimeoutSetExecutable <> a_Value Then
        This.GpibLanReadTimeoutSetExecutable = a_Value
        Me.OnPropertyChanged "GpibLanReadTimeoutSetExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB-Lan Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Selective Device Clear command event. </summary>
Public Sub SelectiveDeviceClearCommand()

    Const p_procedureName = "SelectiveDeviceClearCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_details As String: p_details = VBA.vbNullString
    If Not This.Session.TrySelectiveDeviceClear(p_details, True) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "; " & p_details

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Tries to read (serial poll) or query (<c>*STB?</c>) the device status byte. </summary>
''' <remarks>   This method updates the <see cref="K2700ViewModel.SerialPollByte"/>. </remarks>
''' <param name="canQuery">    [Optional, Boolean, False] True if the <c>*STB?</c>
'''                            can be used to query the status byte in case
'''                            <see cref="TcpSession.GpibLanControllerAttached"/> is false. </param>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean] ture if success. </returns>
Public Function TryReadStatusByte(ByRef a_details As String, Optional ByVal a_canQuery As Boolean = False) As Boolean

    Const p_procedureName As String = "TryReadStatusByte"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_statusByte As Integer
    p_statusByte = This.Device.TryReadStatusByte(p_details)
    If 0 > p_statusByte Then
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                p_details
    Else
        This.SerialPollByte = p_statusByte
    End If
    
    TryReadStatusByte = 0 <= p_statusByte

End Function

''' <summary>   Internal. Serial polls the instrument. </summary>
''' <remarks>   Uses serial poll if available; otherwise <c>*STB?</></remarks>
''' <param name="canQuery">   [Optional, Boolean, False] True if the <c>*STB?</c>
'''                           can be used to query the status byte in case
'''                           <see cref="TcpSession.GpibLanControllerAttached"/> is false. </param>
Friend Sub ReadStatusByte(Optional ByVal a_canQuery As Boolean = False)

    Const p_procedureName = "ReadStatusByte"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_details As String: p_details = VBA.vbNullString
    If Not Me.TryReadStatusByte(p_details, a_canQuery) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "; " & p_details
    
    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte (serial poll) command event. </summary>
Public Sub ReadStatusByteCommand()

    Const p_procedureName = "ReadStatusByteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.ReadStatusByte

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Query the INET Socket. </summary>
Public Sub QuerySocketCommand()

    Const p_procedureName = "QuerySocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_details As String
    If 0 > This.Session.Socket.TryQuery(Me.MessageToSend & This.Session.Termination, _
        Me.ReceivedMessage, p_details, This.Session.ReadAfterWriteDelay) Then
        Me.LastErrorMessage = p_details
    End If
    
    Me.ReadStopper

    Me.SentMessage = This.MessageToSend

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Read from the INET socket. </summary>
Public Sub ReadSocketCommand()

    Const p_procedureName = "ReadSocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_details As String
    If 0 > This.Session.Socket.TryReceive(Me.ReceivedMessage, p_details) Then
        Me.LastErrorMessage = p_details
    End If

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Write a message to the INET socket. </summary>
Public Sub WriteSocketCommand()

    Const p_procedureName = "WriteSocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_success As Boolean
    Dim p_details As String
    p_success = 0 <= This.Session.Socket.TrySendMessage(Me.MessageToSend & This.Session.Termination, p_details, _
        This.Session.ReadAfterWriteDelay)

    Me.ReadStopper

    If p_success Then
    
        Me.SentMessage = This.MessageToSend
    
        If This.ReadStatusAfterWriteEnabled Then
    
            Dim p_canQuery As Boolean
            p_canQuery = Not cc_isr_Ieee488.Scpi.IsQuery(This.MessageToSend)
            Me.ReadStatusByte p_canQuery
    
            If p_canQuery Then
                Me.QueryStatusByteCommand
                Me.ReadStandardEventStatusCommand
            End If
    
        End If
    
    Else
    
        Me.LastErrorMessage = p_details
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Go To Local command event. </summary>
Public Sub GoToLocalCommand()

    Const p_procedureName = "GoToLocalCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.GoToLocal

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Local Lockout command event. </summary>
Public Sub LocalLockoutCommand()

    Const p_procedureName = "LocalLockoutCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.LocalLockout

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Auto Assert Talk Set command event. </summary>
Public Sub AutoAssertTalkSetCommand()

    Const p_procedureName = "AutoAssertTalkSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    If Not This.AutoAssertTalk Then
        This.Session.AutoAssertTalkSetter Me.AutoAssertTalk
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError _
            cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.Name & "." & TypeName(Me) & "." & p_procedureName, _
            "Setting Auto Assert Talk to '" & VBA.CStr(This.AutoAssertTalk) & _
            "' is not allowed."
    End If

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the auto assert TALK get command event. </summary>
Public Sub AutoAssertTalkGetCommand()

    Const p_procedureName = "AutoAssertTalkGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.AutoAssertTalk = This.Session.AutoAssertTalkGetter

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the SRQ command event. </summary>
Public Sub ServiceRequestedQueryCommand()

    ' this does nothing
    ' Me.ServiceRequested = This.Session.ServiceRequested

    ' do this instead.
    Me.ReadStatusByte

End Sub

''' <summary>   Handles the GPIB Address Get command event. </summary>
Public Sub GpibAddressGetCommand()

    Const p_procedureName = "GpibAddressGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_reading As String
    p_reading = This.Session.GpibAddressGetter()
    Me.GpibAddress = This.Session.PrimaryGpibAddress

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the GPIB Address set command event. </summary>
Public Sub GpibAddressSetCommand()

    Const p_procedureName = "GpibAddressSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.GpibAddressSetter Me.GpibAddress

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Timeout get command event. </summary>
Public Sub GpibLanReadTimeoutGetCommand()

    Const p_procedureName = "GpibLanReadTimeoutGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.GpibLanReadTimeout = This.Session.ReadTimeoutGetter

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Timeout set command event. </summary>
Public Sub GpibLanReadTimeoutSetCommand()

    Const p_procedureName = "GpibLanReadTimeoutSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.ReadTimeoutSetter Me.GpibLanReadTimeout

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
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
Public Property Let TopCard(ByVal a_Value As String)
    If This.TopCard <> a_Value Then
        This.TopCard = a_Value
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
Public Property Let TopCardFunctionScanList(ByVal a_Value As String)
    If This.TopCardFunctionScanList <> a_Value Then
        This.TopCardFunctionScanList = a_Value
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
Public Property Let BottomCard(ByVal a_Value As String)
    If This.BottomCard <> a_Value Then
        This.BottomCard = a_Value
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
Public Property Let BottomCardFunctionScanList(ByVal a_Value As String)
    If This.BottomCardFunctionScanList <> a_Value Then
        This.BottomCardFunctionScanList = a_Value
        Me.OnPropertyChanged "BottomCardFunctionScanList"
    End If
End Property

''' <summary>   Gets the measurement mode measuring state. </summary>
''' <remarks>   The condition is true when a measurement mode is configured and active.
''' In continuous and none measurement mode, measuring is false.
''' In immediate mode, measureing is true after the mode is configured and data is collected
''' and false after data collection is completed.
''' In External Mode (before monitoring states), measuring is true after the external trigger
''' is configured.
''' In monitoring mode, measuring is true while monitoring is active and false once monitoring
''' is ended.
''' </remarks>
''' <value>   [Boolean]. </value>
Public Property Get Measuring() As Boolean
    Measuring = This.Measuring
End Property

''' <summary>   Sets the measuring state. True when measurements are collected. </summary>
''' <param name="a_value">   [Boolean] the measuring state; true when measuring is active. </value>
Public Property Let Measuring(ByVal a_Value As Boolean)
    If This.Measuring <> a_Value Then
        This.Measuring = a_Value
        Me.OnPropertyChanged "Measuring"
    End If
End Property

''' <summary>   Gets the design mode for unit testing and restoration of state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get DesignMode() As Boolean
    DesignMode = This.DesignMode
End Property

''' <summary>   Sets the design mode for unit testing and restoration of state. </summary>
''' <param name="a_value">   [Boolean] true if in design mode. </param>
Public Property Let DesignMode(ByVal a_Value As Boolean)
    If This.DesignMode <> a_Value Then
        This.DesignMode = a_Value
        Me.OnPropertyChanged "DesignMode"
    End If
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get MeasurementMode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    MeasurementMode = This.MeasurementMode
End Property

''' <summary>   Sets the measurement mode. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Friend Property Let MeasurementMode(ByVal a_Value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_Value Then
        This.MeasurementMode = a_Value
        Me.OnPropertyChanged "MeasurementMode"
    End If
End Property

''' <summary>   Sets the measurement mode for unit testing. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
Public Sub MeasurementModeUnitTestSetter(ByVal a_Value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_Value Then
        This.MeasurementMode = a_Value
        Me.OnPropertyChanged "MeasurementMode"
    End If
End Sub

''' <summary>   Gets the Front Inputs Required. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsRequired() As Boolean
    FrontInputsRequired = This.FrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are Required. </value>
Public Property Let FrontInputsRequired(ByVal a_Value As Boolean)
    If This.FrontInputsRequired <> a_Value Then
        This.FrontInputsRequired = a_Value
        Me.OnPropertyChanged "FrontInputsRequired"
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
Public Property Let ReadingBeepEnabled(ByVal a_Value As Boolean)
    If This.ReadingBeepEnabled <> a_Value Then
        This.ReadingBeepEnabled = a_Value
        Me.OnPropertyChanged "ReadingBeepEnabled"
    End If
End Property

''' <summary>   Gets the Front Inputs has value condition. </summary>
''' <remarks>   True if front inputs where read from the instrument. </remarks>
''' <Validated>   [Boolean]. </Validated>
Public Property Get FrontInputsHasValue() As Boolean
    FrontInputsHasValue = This.FrontInputsHasValue
End Property

''' <summary>   Sets the Front Inputs has value condition. </summary>
''' <param name="a_Value">   [Boolean] true if Front Inputs value was read from the instrument. </Validated>
Public Property Let FrontInputsHasValue(ByVal a_Value As Boolean)
    If This.FrontInputsHasValue <> a_Value Then
        This.FrontInputsHasValue = a_Value
        Me.OnPropertyChanged "FrontInputsHasValue"
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
Public Property Let FrontInputsValue(ByVal a_Value As Boolean)
    If This.FrontInputsValue <> a_Value Then
        This.FrontInputsValue = a_Value
        Me.OnPropertyChanged "FrontInputsValue"
        Me.FrontInputsCaption = IIf(a_Value, "Front", "Rear")
    End If
End Property

''' <summary>   Gets the Front Inputs Caption. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsCaption() As String
    FrontInputsCaption = This.FrontInputsCaption
End Property

''' <summary>   Sets the Front Inputs Caption. </summary>
''' <param name="a_value">   [String] the Front Inputs Caption. </value>
Public Property Let FrontInputsCaption(ByVal a_Value As String)
    If This.FrontInputsCaption <> a_Value Then
        This.FrontInputsCaption = a_Value
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
Public Property Let ChannelCount(ByVal a_Value As Integer)
    If This.ChannelCount <> a_Value Then
        This.ChannelCount = a_Value
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
Public Property Let SelectedChannelNumber(ByVal a_Value As Integer)
    If This.SelectedChannelNumber <> a_Value Then
        This.SelectedChannelNumber = a_Value
        Me.OnPropertyChanged "SelectedChannelNumber"
    End If
End Property

''' <summary>   Gets the Target Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get TargetChannelNumber() As Integer
    TargetChannelNumber = This.TargetChannelNumber
End Property

''' <summary>   Sets the Target Channel Number. </summary>
''' <param name="a_value">   [Integer] the Target Channel Number value. </value>
Public Property Let TargetChannelNumber(ByVal a_Value As Integer)
    If This.TargetChannelNumber <> a_Value Then
        This.TargetChannelNumber = a_Value
        Me.OnPropertyChanged "TargetChannelNumber"
    End If
End Property

''' <summary>   Gets the Measured Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredChannelNumber() As Integer
    MeasuredChannelNumber = This.MeasuredChannelNumber
End Property

''' <summary>   Sets the Measured Channel Number. </summary>
''' <param name="a_value">   [Integer] the Measured Channel Number value. </value>
Public Property Let MeasuredChannelNumber(ByVal a_Value As Integer)
    If This.MeasuredChannelNumber <> a_Value Then
        This.MeasuredChannelNumber = a_Value
        Me.OnPropertyChanged "MeasuredChannelNumber"
    End If
End Property

''' <summary>   Gets the Sense Function Name for continuous mode. </summary>
''' <value>   [String]. </value>
Public Property Get ContinuousSenseFunctionName() As String
    ContinuousSenseFunctionName = This.ContinuousSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for continuous mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ContinuousSenseFunctionName(ByVal a_Value As String)
    If This.ContinuousSenseFunctionName <> a_Value Then
        This.ContinuousSenseFunctionName = a_Value
        Me.OnPropertyChanged "ContinuousSenseFunctionName"
    End If
End Property

''' <summary>   Gets the Sense Function Name for Immediate trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get ImmediateSenseFunctionName() As String
    ImmediateSenseFunctionName = This.ImmediateSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for Immediate trigger  mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ImmediateSenseFunctionName(ByVal a_Value As String)
    If This.ImmediateSenseFunctionName <> a_Value Then
        This.ImmediateSenseFunctionName = a_Value
        Me.OnPropertyChanged "ImmediateSenseFunctionName"
    End If
End Property

''' <summary>   Gets the Sense Function Name for External trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get ExternalSenseFunctionName() As String
    ExternalSenseFunctionName = This.ExternalSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for External trigger mode. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let ExternalSenseFunctionName(ByVal a_Value As String)
    If This.ExternalSenseFunctionName <> a_Value Then
        This.ExternalSenseFunctionName = a_Value
        Me.OnPropertyChanged "ExternalSenseFunctionName"
    End If
End Property

''' <summary>   Gets the sense function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Internal. Sets the sense function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Friend Property Let SenseFunctionName(ByVal a_Value As String)
    If This.SenseFunctionName <> a_Value Then
        This.SenseFunctionName = a_Value
        Me.OnPropertyChanged "SenseFunctionName"
        Me.OnSenseFunctionChanged
    End If
End Property

''' <summary>   Handles the sense function change; Set the sense faction caption and units. </summary>
Public Sub OnSenseFunctionChanged()

    Select Case This.SenseFunctionName

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
Public Property Let SenseFunctionCaption(ByVal a_Value As String)
    If This.SenseFunctionCaption <> a_Value Then
        This.SenseFunctionCaption = a_Value
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
Public Property Let SenseFunctionUnit(ByVal a_Value As String)
    If This.SenseFunctionUnit <> a_Value Then
        This.SenseFunctionUnit = a_Value
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
Public Property Let MeasuredReading(ByVal a_Value As String)
    If This.MeasuredReading <> a_Value Then
        This.MeasuredReading = a_Value
        Me.OnPropertyChanged "MeasuredReading"
        If 0 < VBA.Len(a_Value) Then
            Dim p_value As Double
            Dim p_details As String
            If cc_isr_Core.StringExtensions.TryParseDouble(a_Value, p_value, p_details) Then
                Me.MeasuredValue = cc_isr_Ieee488.Syntax.NotANumber
            Else
                Me.MeasuredValue = p_value
            End If
        Else
            Me.MeasuredValue = cc_isr_Ieee488.Syntax.NotANumber
        End If
    End If
End Property

''' <summary>   Gets the Measured Value Format. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredValueFormat() As String
    MeasuredValueFormat = This.MeasuredValueFormat
End Property

''' <summary>   Sets the Measured Value Format. </summary>
''' <param name="a_value">   [String] the Measured Value Format value. </value>
Public Property Let MeasuredValueFormat(ByVal a_Value As String)
    If This.MeasuredValueFormat <> a_Value Then
        This.MeasuredValueFormat = a_Value
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
Public Property Let MeasuredValue(ByVal a_Value As Double)
    If This.MeasuredValue <> a_Value Then
        This.MeasuredValue = a_Value
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
Public Property Let SenseFunctions(ByVal a_Value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.SenseFunctions, a_Value) Then
        Set This.SenseFunctions = a_Value
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
Public Property Let ReadingOffset(ByVal a_Value As Double)
    If This.ReadingOffset <> a_Value Then
        This.ReadingOffset = a_Value
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
Public Property Let MonitoringState(ByVal a_Value As String)
    If This.MonitoringState <> a_Value Then
        This.MonitoringState = a_Value
        Me.OnPropertyChanged "MonitoringState"
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
Public Property Let AutoIncrementChannelNoEnabled(ByVal a_Value As Boolean)
    If This.AutoIncrementChannelNoEnabled <> a_Value Then
        This.AutoIncrementChannelNoEnabled = a_Value
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
Public Property Let SingleReadEnabled(ByVal a_Value As Boolean)
    If This.SingleReadEnabled <> a_Value Then
        This.SingleReadEnabled = a_Value
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
Public Property Let ClearReadingsExecutable(ByVal a_Value As Boolean)
    If This.ClearReadingsExecutable <> a_Value Then
        This.ClearReadingsExecutable = a_Value
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
Public Property Let QueryInputsExecutable(ByVal a_Value As Boolean)
    If This.QueryInputsExecutable <> a_Value Then
        This.QueryInputsExecutable = a_Value
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
Public Property Let MeasureExecutable(ByVal a_Value As Boolean)
    If This.MeasureExecutable <> a_Value Then
        This.MeasureExecutable = a_Value
        Me.OnPropertyChanged "MeasureExecutable"
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
Public Property Let ImmediateTriggerOptionExecutable(ByVal a_Value As Boolean)
    If This.ImmediateTriggerOptionExecutable <> a_Value Then
        This.ImmediateTriggerOptionExecutable = a_Value
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
Public Property Let ExternalTriggerOptionExecutable(ByVal a_Value As Boolean)
    If This.ExternalTriggerOptionExecutable <> a_Value Then
        This.ExternalTriggerOptionExecutable = a_Value
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
Public Property Let FunctionNameExecutable(ByVal a_Value As Boolean)
    If This.FunctionNameExecutable <> a_Value Then
        This.FunctionNameExecutable = a_Value
        Me.OnPropertyChanged "FunctionNameExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Start Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Start Monitoring External Trigger command can be executed. </value>
Public Property Get StartMonitoringExecutable() As Boolean
    StartMonitoringExecutable = This.StartMonitoringExecutable
End Property

''' <summary>   Sets the condition indicating that the Start Monitoring External Trigger command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Start Monitoring External Trigger command can be executed. </param>
Public Property Let StartMonitoringExecutable(ByVal a_Value As Boolean)
    If This.StartMonitoringExecutable <> a_Value Then
        This.StartMonitoringExecutable = a_Value
        Me.OnPropertyChanged "StartMonitoringExecutable"
        Me.OnPropertyChanged "ToggleMonitoringExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </value>
Public Property Get StopMonitoringExecutable() As Boolean
    StopMonitoringExecutable = This.StopMonitoringExecutable
End Property

''' <summary>   Sets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </param>
Public Property Let StopMonitoringExecutable(ByVal a_Value As Boolean)
    If This.StopMonitoringExecutable <> a_Value Then
        This.StopMonitoringExecutable = a_Value
        Me.OnPropertyChanged "StopMonitoringExecutable"
        Me.OnPropertyChanged "ToggleMonitoringExecutable"
    End If
End Property

''' <summary>   Gets the condition indicating that the Toggle Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Toggle Monitoring External Trigger command can be executed. </value>
Public Property Get ToggleMonitoringExecutable() As Boolean
    ToggleMonitoringExecutable = This.StartMonitoringExecutable Or This.StopMonitoringExecutable
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns true if the inputs are valid for front or rear settings. </summary>
''' <param name"a_userMessage">   [Out, String] User message. </param>
''' <param name"a_promptUser">    [Boolean] User message. </param>
''' <returns>   [Boolean]. </returns>
Public Function ValidateInputs(ByRef a_userMessage As String, ByVal a_promptUser As Boolean) As Boolean

    ' synchronize.
    ' Me.Device.TryQueryOperationCompleted(p_details)
    
    ' query front switch has been returning false values
    ' it seems this is occurring after the instrument is configured for its
    ' triggering mode.
    
    ' until this cludge was added.
    ' cc_isr_Core_IO.Factory.NewStopwatch.Wait 10
    ' Me.K2700.ScpiSystem.QueryFrontSwitch
    
    Dim p_success As Boolean: p_success = False
    Dim p_userMessage As String: p_userMessage = VBA.vbNullString

    ' gets the front input condition: True is Front.
    Me.FrontInputsValue = This.K2700.ScpiSystem.QueryFrontSwitch
    
    ' indicate that front inputs has a value.
    Me.FrontInputsHasValue = True
    
    ' check if front or rear
    If This.FrontInputsRequired And Not This.FrontInputsValue Then

        ' set the proper message
        p_userMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."

        p_success = False
    
    ElseIf Not This.FrontInputsRequired And This.FrontInputsValue Then

        ' set the proper message
        p_userMessage = "Please depress the front panel INPUTS button to its REAR (R) position and try again."
    
        p_success = False

    Else
    
        p_userMessage = VBA.vbNullString

        p_success = True

    End If

    If Not p_success And a_promptUser Then
        MsgBox p_userMessage, vbOKOnly Or vbInformation, "INPUTS button change required"
    End If
    
    a_userMessage = p_userMessage
    ValidateInputs = p_success

End Function

''' <summary>   Tries to validate the expected function by name. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Friend Function TryValidateSenseFunction(ByVal a_functionName As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryValidateSenseFunction"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    
    ' validate the sense function.
    Dim p_senseFunctionName As String
    p_senseFunctionName = This.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
    p_success = 1 = VBA.InStr(1, a_functionName, p_senseFunctionName, vbTextCompare)
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; The instrument sense function '" & p_senseFunctionName & _
            "' is different from the expected '" & This.SenseFunctionName & "' sense function."

    TryValidateSenseFunction = p_success
    
End Function

''' <summary>   Internal. Tries to configure the intial trigger mode to continuous trigger mode. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Friend Function TryConfigureIntitialTriggerMode(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureIntitialTriggerMode"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' notify of the changing measurement state
    Me.OnMeasurementStateChanging
    
    ' reset and clear the instrument
    p_success = This.K2700.Device.TryResetClear(p_details)
    
    If Not p_success Then _
        p_details = "Failed resetting and clearing the IEEE488 device to its known state; " & p_details
    
    If p_success Then
    
        ' default to continuous sense function
        Me.SenseFunctionName = This.ContinuousSenseFunctionName

        ' configure measuring a single value on each read.
        p_success = This.K2700.TryConfigureContinuousFunctionAutoRange(This.SenseFunctionName, p_details)
        
        If Not p_success Then _
            p_details = "Failed configuring continous trigger state; " & p_details
        
        If p_success Then
        
            ' validate the sense function.
            p_success = Me.TryValidateSenseFunction(This.SenseFunctionName, p_details)
        
            If Not p_success Then _
                p_details = "Failed setting the sense function; " & p_details
        
        End If
        
    End If
 
    ' enable the mode executables and conditions.
    If p_success Then
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous, "continuous", False
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False
    End If

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring initial trigger mode; " & p_details
    
    TryConfigureIntitialTriggerMode = p_success
    
End Function

''' <summary>   Internal. Configures the intial trigger mode to continuous trigger mode. </summary>
''' <exception cref="cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError">   Thrown if this method
''' failed. </exception>
Friend Sub ConfigureIntitialTriggerMode()

    Const p_procedureName = "ConfigureIntitialTriggerMode"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_success As Boolean
    Dim p_details As String
    p_success = Me.TryConfigureIntitialTriggerMode(p_details)
    
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Internal. Tries to configure the continuous trigger mode. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Friend Function TryConfigureContinuousMode(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureContinuousMode"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' notify of the changing measurement state
    Me.OnMeasurementStateChanging

    If p_success Then
    
        ' reset if changing modes.
        If cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous <> This.MeasurementMode Then
            
            p_success = This.K2700.Device.TryResetClear(p_details)
        
            If Not p_success Then _
                p_details = "Failed resetting and clearing the IEEE488 device state; " & p_details
        
        Else
        
            ' clear execution state
            p_success = This.K2700.Device.TryClearExecutionState(p_details)
        
            If Not p_success Then _
                p_details = "Failed clearing the IEEE488 device execution state; " & p_details
        
        End If
    
        If Not p_success Then
        
            ' if instrument failed to reset, attempt to restore its initial state
            
            p_success = Me.TryRestoreInitialState(p_details)
            
            If p_success Then
            
                ' return failure if the initial state failed to restore
                p_success = Not Me.ShouldRestoreInitialState(p_details)
                
                If Not p_success Then _
                    p_details = "Restoring initial state failed; " & p_details
        
            Else
            
                p_details = "Failed restoring initial state; " & p_details
            
            End If
        
        End If
        
    End If
    
    If p_success Then
        
        Me.SenseFunctionName = This.ContinuousSenseFunctionName
    
        ' configure measuring a single value on each read.
        p_success = This.K2700.TryConfigureContinuousFunctionAutoRange(This.SenseFunctionName, p_details)
        
        If Not p_success Then _
            p_details = "Failed configuring the 2700 continuous trigger mode; " & p_details
        
    End If

    If p_success Then
    
        ' validate the sense function.
        p_success = Me.TryValidateSenseFunction(This.SenseFunctionName, p_details)
    
        If Not p_success Then _
            p_details = "Failed setting the 2700 instrument sense function; " & p_details
    
    End If
    
    ' enable the mode executables and conditions.
    If p_success Then
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous, "continuous", False
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring continuous mode; " & p_details

    TryConfigureContinuousMode = p_success
    
End Function

''' <summary>   Configure continuous trigger mode. </summary>
''' <param name="a_details">   [Out, String] Details why the command failed. </param>
''' <returns>   [Boolean] True if the command succeeded. </returns>
Public Function ConfigureContinuousModeCommand(ByVal a_details As String) As Boolean

    Const p_procedureName = "ConfigureContinuousModeCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = False
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' configure for continuous trigger

    p_success = Me.TryConfigureContinuousMode(p_details)
    
    ' throw an error if failed to configure
    
    If Not p_success Then
        a_details = "Failed configuring continuous trigger mode; " & p_details
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " " & a_details
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ConfigureContinuousModeCommand = p_success
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    p_success = False
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' this primarily is for unit testing.
    a_details = Me.LastErrorMessage

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Internal. Tries to configure the external trigger reading mode. </summary>
''' <remarks>   This method requires that instrument be connected as per
''' <see cref=TryRestoreInitialState"/> and inputs are set as required.
''' This method enables the trigger monitoring executable and disables the
''' immediate reading executable (Measure). </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Friend Function TryConfigureExternalTriggerReading(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureExternalTriggerReading"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' notify of the changing measurement state
    Me.OnMeasurementStateChanging

    If p_success Then
    
        ' unfortunately, reading the inputs after the reset belore returns an incorrect value
        ' so this was moved as a workaround
        
        ' it looks like we need to validate because the instrument inputs can be changed at any time/
        If True Then ' Not Me.FrontInputsHasValue Or (This.FrontInputsRequired <> Not This.FrontInputsValue) Then
            
            p_success = Me.ValidateInputs(p_details, True)
        
            Me.LastMessage = p_details
            
        End If
        
    End If

    If p_success Then
    
        ' reset if changing modes.
        If cc_isr_Tcp_Scpi.MeasurementModeOption.External <> This.MeasurementMode Then
            
            p_success = This.K2700.Device.TryResetClear(p_details)
        
            If Not p_success Then _
                p_details = "Failed resetting and clearing the IEEE488 device; " & p_details
        
        Else
        
            p_success = This.K2700.Device.TryClearExecutionState(p_details)
        
            If Not p_success Then _
                p_details = "Failed clearing the IEEE488 device execution state; " & p_details
        
        End If
    
        If Not p_success Then
        
            ' if instrument failed to reset, attempt to restore its initial state
            
            p_success = Me.TryRestoreInitialState(p_details)
            
            If p_success Then
            
                ' return failure if the initial state failed to restore
                p_success = Not Me.ShouldRestoreInitialState(p_details)
                
                If Not p_success Then _
                    p_details = "Restoring initial state failed; " & p_details
        
            Else
            
                p_details = "Failed restoring initial state; " & p_details
            
            End If
        
        End If
        
    End If

    If p_success Then
    
        ' Configure external triggering reading requiring the front panel
        
        Me.SenseFunctionName = This.ExternalSenseFunctionName
        
        p_success = This.K2700.TryConfigureExternalTriggerReading(This.SenseFunctionName, p_details)
        
        If Not p_success Then _
            p_details = "Failed configuring the 2700 external trigger reading mode; " & p_details
    
    End If
    
    ' this is done in the previous command.
    ' This.Session.WriteLine ":TRIG:SOUR IMM"
    ' This.Session.WriteLine ":TRIG:SOUR EXT"
    
    ' clear the serial poll register
    ' these are also part of the previous command.
    ' This.Session.WriteLine "*ESE 0"
    ' This.Session.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    ' This.Session.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    ' This.Session.WriteLine "STAT:MEAS:ENAB 32"
    
    ' This.Device.QueryServiceRequestStatus
    
    ' move this to starting the monitoring.
    ' initiate awaiting for the external trigger
    ' This.Session.WriteLine ":INIT;*CLS"

    If p_success Then
    
        ' validate the sense function.
        p_success = Me.TryValidateSenseFunction(This.SenseFunctionName, p_details)
    
        If Not p_success Then _
            p_details = "Failed setting the 2700 instrument sense function; " & p_details
    
    End If

    ' enable the mode executables and conditions.
    If p_success Then
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.External, "external trigger mode", True
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring external trigger readings; " & p_details

    TryConfigureExternalTriggerReading = p_success
    
End Function

''' <summary>   Configures external trigger reading mode. </summary>
''' <param name="a_details">   [Out, String] Details why the command failed. </param>
''' <returns>   [Boolean] True if the command succeeded. </returns>
Public Function ConfigureExternalTriggerReadingCommand(ByRef a_details As String) As Boolean

    Const p_procedureName = "ConfigureExternalTriggerReadingCommand"

    ' Trap errors to the error handler
    
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = False
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' configure for external trigger readings
    p_success = Me.TryConfigureExternalTriggerReading(p_details)
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring external trigger reading measurement state; " & p_details
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ConfigureExternalTriggerReadingCommand = p_success
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' this primarily is for unit testing.
    a_details = Me.LastErrorMessage

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler


End Function


''' <summary>   Tries to configure the immediate trigger reading mode. </summary>
''' <remarks>   This method requires that instrument be connected as per
''' <see cref="TryRestoreInitialState"/> and inputs are set as required.
''' This method enables the immediate reading executable (Measure) and
''' disables trigger monitoring executable. </remarks>
''' <param name="a_details">   [Out, String] Details why the method failed. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureImmediateTriggerReadings(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureImmediateTriggerReadings"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' notify of the changing measurement state
    Me.OnMeasurementStateChanging

    If p_success Then
    
        ' unfortunately, reading the inputs after the reset belore returns an incorrect value
        ' so this was moved as a workaround
        
        ' it looks like we need to validate because the instrument inputs can be changed at any time/
        If True Then ' Not Me.FrontInputsHasValue Or (This.FrontInputsRequired <> Not This.FrontInputsValue) Then
            
            p_success = Me.ValidateInputs(p_details, True)
        
            Me.LastMessage = p_details
            
        End If
        
    End If

    If p_success Then
    
        ' reset if changing modes.
        If cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate <> This.MeasurementMode Then
            
            p_success = This.K2700.Device.TryResetClear(p_details)
        
            If Not p_success Then _
                p_details = "Failed resetting and clearing the IEEE488 device state; " & p_details
        
        Else
        
            ' clear execution state
            p_success = This.K2700.Device.TryClearExecutionState(p_details)
        
            If Not p_success Then _
                p_details = "Failed clearing the IEEE488 device execution state; " & p_details
        
        End If
    
        If Not p_success Then
        
            ' if instrument failed to reset, attempt to restore its initial state
            
            p_success = Me.TryRestoreInitialState(p_details)
            
            If p_success Then
            
                ' return failure if the initial state failed to restore
                p_success = Not Me.ShouldRestoreInitialState(p_details)
                
                If Not p_success Then _
                    p_details = "Restoring initial state failed; " & p_details
        
            Else
            
                p_details = "Failed restoring initial state; " & p_details
            
            End If
        
        End If
        
    End If
    
    If p_success Then
    
        Me.SenseFunctionName = This.ImmediateSenseFunctionName
        
        ' configure measuring a single value on each read.
        p_success = This.K2700.TryConfigureSingleFunctionReading(This.SenseFunctionName, p_details)
        
        If Not p_success Then _
            p_details = "Failed configuring the 2700 immediate trigger reading mode; " & p_details
        
    End If
    
    If p_success Then
    
        ' validate the sense function.
        p_success = Me.TryValidateSenseFunction(This.SenseFunctionName, p_details)
    
        If Not p_success Then _
            p_details = "Failed setting the 2700 instrument sense function; " & p_details
    
    End If
    
    ' enable the mode executables and conditions.
    If p_success Then
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate, "immediate mode", True
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring immediate trigger readings; " & p_details
    
    TryConfigureImmediateTriggerReadings = p_success

End Function

''' <summary>   Configures immediate trigger reading mode. </summary>
''' <param name="a_details">   [Out, String] Details why the command failed. </param>
''' <returns>   [Boolean] True if the command succeeded. </returns>
Public Function ConfigureImmediateTriggerReadingsCommand(ByRef a_details As String) As Boolean

    Const p_procedureName = "ConfigureImmediateTriggerReadingsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = False
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' configure for immediate trigger readings
    p_success = Me.TryConfigureImmediateTriggerReadings(p_details)
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring immediate trigger readings measurement state; " & Me.LastMessage

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ConfigureImmediateTriggerReadingsCommand = p_success
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
    
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' this primarily is for unit testing.
    a_details = Me.LastErrorMessage
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary> Queries the Inputs status. </summary>
Public Sub QueryInputsCommand()

    Const p_procedureName = "QueryInputsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.FrontInputsValue = This.K2700.ScpiSystem.QueryFrontSwitch

    Me.ReadStopper

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Tries to measure immediately. </summary>
''' <remarks>   This method assumes that the inputs were validated before configuring the
''' measurement state. </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryMeasureImmediately(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryMeasureImmediately"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    Dim p_reading As String
    Dim p_channelNumber As Integer

    If This.AutoIncrementChannelNoEnabled Then
    
        ' Clear all cells
        Me.ClearReadings

        ' scan all channels
        For p_channelNumber = 1 To Me.ChannelCount
    
            ' allow events to occur
            VBA.DoEvents
        
            ' set the Target Channel Number
            Me.TargetChannelNumber = p_channelNumber
        
            Me.RestartStopper
            If This.FrontInputsValue Then
                p_success = This.K2700.TryTakeFrontReading(p_reading, p_details, p_channelNumber)
                If Not p_success Then _
                    p_details = "Failed taking a front reading; " & p_details
            Else
                p_success = This.K2700.TryTakeRearReading(p_channelNumber, p_reading, p_details)
                If Not p_success Then _
                    p_details = "Failed taking a rear reading; " & p_details
            End If
            Me.ReadStopper
    
            If Not p_success Then Exit For
            
        Next p_channelNumber

    Else

        p_channelNumber = This.SelectedChannelNumber
    
        Me.RestartStopper
        
        If This.FrontInputsValue Then
            p_success = This.K2700.TryTakeFrontReading(p_reading, p_details, p_channelNumber)
            If Not p_success Then _
                p_details = "Failed taking a front reading; " & p_details
        Else
            p_success = This.K2700.TryTakeRearReading(p_channelNumber, p_reading, p_details)
            If Not p_success Then _
                p_details = "Failed taking a rear reading; " & p_details
        End If
        
        Me.ReadStopper
    
    End If
    
    If Not p_success Then
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    End If

    TryMeasureImmediately = p_success

End Function

''' <summary>   Takes a single or multiple readings from the front or rear panel. </summary>
''' <remarks>   This method assumes that the inputs were verified upon configuring the
''' measurement mode. </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Public Function MeasureImmediatelyCommand(ByRef a_details As String) As Boolean

    Const p_procedureName = "MeasureImmediatelyCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' synchronize.
    ' ?? p_success = Me.Device.TryQueryOperationCompleted(p_details)
    
    If p_success Then
    
        ' disable the relevant commands.
        Me.MeasureExecutable = False
        Me.ImmediateTriggerOptionExecutable = False
        Me.ExternalTriggerOptionExecutable = False
    
        ' throw an error if measurement failed
        p_success = Me.TryMeasureImmediately(p_details)
        
        If Not p_success Then
        
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " failed taking an immediate measurment; " & p_details
        
        End If
    
    Else
    
        Me.LastMessage = p_details
        a_details = p_details
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' turn off the measuring state
    Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate, "Done", False

    MeasureImmediatelyCommand = p_success
    
    Me.MeasureExecutable = True
    Me.ImmediateTriggerOptionExecutable = True
    Me.ExternalTriggerOptionExecutable = True

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' for unit testing
    a_details = Me.LastErrorMessage
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Validates the Inputs and takes a single or multiple readings from the front or rear panel. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
''' <returns>   [Boolean]. </returns>
Public Function ValidateInputsAndMeasureCommand(ByRef a_details As String) As Boolean

    Const p_procedureName = "ValidateInputsAndMeasureCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' reinitialize if device not defined
    If Not This.K2700.Connected Then
        ' This.K2700.InitializeLegacy Me.host, Me.port, Me.SessionTimeout
    End If

    ' synchronize.
    Dim p_reply As String
    p_success = Me.Device.TryQueryOperationCompleted(p_reply, p_details)
    
    ' validate inputs returns incorrect values at this point.
    ' it must have something to do with the configuration
    p_success = Me.ValidateInputs(p_details, True)
    If Not p_success Then
        ' this cludge was added.
        cc_isr_Core_IO.Factory.NewStopwatch.Wait 50 ' 25 this was not enough
        p_success = Me.ValidateInputs(p_details, True)
    End If
    
    If p_success Then
    
        ' disable the relevant commands.
        Me.MeasureExecutable = False
        Me.ImmediateTriggerOptionExecutable = False
        Me.ExternalTriggerOptionExecutable = False
    
        ' throw an error if measurement failed
        p_success = Me.TryMeasureImmediately(p_details)
        
        If Not p_success Then
        
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " failed taking an immediate measurment; " & p_details
        
        End If
    
    Else
    
        Me.LastMessage = p_details
        a_details = p_details
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ValidateInputsAndMeasureCommand = p_success
    
    Me.MeasureExecutable = True
    Me.ImmediateTriggerOptionExecutable = True
    Me.ExternalTriggerOptionExecutable = True

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' for unit testing
    a_details = Me.LastErrorMessage
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Internal. Clears the readings. </summary>
Friend Sub ClearReadings()

    Const p_procedureName = "ClearReadings"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_channelNumber As Integer
    For p_channelNumber = 1 To Me.MaximumChannelNumber
        VBA.DoEvents
        OnChannelMeasured p_channelNumber, Me.ReadingOffset
    Next p_channelNumber
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Clears the readings. </summary>
Public Sub ClearReadingsCommand()

    Const p_procedureName = "ClearReadingsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.ClearReadings
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the Stop Requested condition. </summary>
''' <remarks>   When true, a Stop request is set so that the timer event could be used to stop
''' monitring for external triggers. Thereafter a pause request is turned on so as to
''' end the timer events. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get StopRequested() As Boolean
    StopRequested = This.StopRequested
End Property

''' <summary>   Sets the Stop Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Stop Requested value. </value>
Friend Property Let StopRequested(ByVal a_Value As Boolean)
    If This.StopRequested <> a_Value Then
        This.StopRequested = a_Value
        Me.OnPropertyChanged "StopRequested"
    End If
End Property

''' <summary>   Gets the External Trigger Monitoring Enabled condition. </summary>
''' <remarks>   Indicates that external trigger monitoring is on. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = This.ExternalTrigMonitoringEnabled
End Property

''' <summary>   Sets the External Trigger Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trigger Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_Value As Boolean)
    If This.ExternalTrigMonitoringEnabled <> a_Value Then
        This.ExternalTrigMonitoringEnabled = a_Value
        Me.OnPropertyChanged "ExternalTrigMonitoringEnabled"
    End If
End Property

''' <summary>   Gets the Timer Controlled condition. </summary>
''' <value>   [Boolean] True if external trigger monitoring is timer controlled; otherwise,
''' external trigger monitoring is polled, which is solely for testing purposes. </value>
Public Property Get TimerControlled() As Boolean
    TimerControlled = This.TimerControlled
End Property

''' <summary>   Gets the Timer Started condition. </summary>
''' <value>   [Boolean] True if the timer is active. </value>
Public Property Get TimerStarted() As Boolean
    TimerStarted = This.TimerStarted
End Property


''' <summary>   Internal. Starts monitoring the instrument for external triggered measurements. </summary>
''' <param name="a_timerControlled">   [Optional, Boolean, True] true time controlled; otherwise, the timer event
''' handler will be polled. </value>
Public Sub StartMonitoringExternalTriggers(Optional ByVal a_timerControlled As Boolean = True)

    Const p_procedureName = "StartMonitoringExternalTriggers"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.AutoIncrementChannelNoEnabled Then
    
        ' Clear all readings
        Me.ClearReadings

    End If

    Me.MonitoringState = "External Trigger On"

    If This.AutoIncrementChannelNoEnabled Then
    
        ' set the first channel number for manual scan
        This.TargetChannelNumber = 1
    
    Else

        ' set the single reading channel number for manual scan
        ' to the channel number entered by the operator.
        This.TargetChannelNumber = This.SelectedChannelNumber
    
    End If
   
    ' initiate awaiting for the external trigger
    ' this is handled by the K2700.MonitorExternalTrigger method.
    ' This.Session.WriteLine ":INIT;*CLS"

    This.TimerControlled = a_timerControlled
    
    If This.TimerControlled Then
    
       ' start the timer
       Me.StartTimer Me.TimerInterval
    
       ' take some time for it to take effect.
       cc_isr_Core_IO.Factory.NewStopwatch.Wait Me.TimerInterval
    
    Else
        
        ' if not using timer, go to monitoring mode
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, "Trigger monitoring commenced", True
        
    End If
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' on exception, stop the timer
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Starts monitoring the instrument for external triggered measurements. </summary>
Public Sub StartMonitoringExternalTriggersCommand()

    Const p_procedureName = "StartMonitoringExternalTriggersCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.StartMonitoringExternalTriggers
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' on exception, stop the timer.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Internal. Stops monitoring the instrument for external triggered measurements. </summary>
Friend Sub StopMonitoringExternalTriggers()

    Const p_procedureName = "StopMonitoringExternalTriggers"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' request a stop of external trigger monitoring.
    Me.StopRequested = True
    
    If This.TimerControlled Then
    
        Dim p_stopper As cc_isr_Core_IO.Stopwatch
        Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch
        
        ' wait for the timer to stop over two timer intervals
        Dim p_timeout As Integer: p_timeout = 2 * Me.TimerInterval
        p_stopper.Restart
        
        ' wait until timer stopped
        Do While (Not m_eventTimer Is Nothing) And This.TimerStarted And p_stopper.ElapsedMilliseconds < p_timeout
            DoEvents
        Loop
        
        If This.TimerStarted Then
        
            ' if timer has not stopped, force stop
            This.TimerStarted = False
            If Not m_eventTimer Is Nothing Then m_eventTimer.PauseTimer
            
        End If

    Else
    
        ' pass throw one more time the timer event handler, which will
        ' turn on the done flag, pause the timer and move the
        ' measurement mode to none.
        
        Me.HandleTimerEvent
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' if necessary turn off measuring state
    If cc_isr_Tcp_Scpi.MeasurementModeOption.None <> This.MeasurementMode Then _
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, "Trigger monitoring done", False

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' on exception, stop the timer.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Stops monitoring the instrument for external triggered measurements. </summary>
Public Sub StopMonitoringExternalTriggersCommand()

    Const p_procedureName = "StopMonitoringExternalTriggersCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.StopMonitoringExternalTriggers
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' on exception stop the timer.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Pause Requested condition. </summary>
''' <remarks>   When true, the timer is pauses on the next timer event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get PauseRequested() As Boolean
    PauseRequested = This.PauseRequested
End Property

''' <summary>   Sets the Pause Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Pause Requested value. </value>
Friend Property Let PauseRequested(ByVal a_Value As Boolean)
    If This.PauseRequested <> a_Value Then
        This.PauseRequested = a_Value
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
Public Property Let TimerInterval(ByVal a_Value As Integer)
    If This.TimerInterval <> a_Value Then
        This.TimerInterval = a_Value
        Me.OnPropertyChanged "TimerInterval"
    End If
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Double]. </value>
Public Property Get ElapsedTime() As Double
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Double] the Elapsed Time in milliseconds. </value>
Public Property Let ElapsedTime(ByVal a_Value As Double)
    If This.ElapsedTime <> a_Value Then
        This.ElapsedTime = a_Value
        Me.OnPropertyChanged "ElapsedTime"
        Me.ElapsedTimeCaption = VBA.Format$(a_Value, This.ElapsedTimeFormat)
    End If
End Property

''' <summary>   Gets the Elapsed Time Format. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeFormat() As String
    ElapsedTimeFormat = This.ElapsedTimeFormat
End Property

''' <summary>   Sets the Elapsed Time Format. </summary>
''' <param name="a_value">   [String] the elapsed time format. </value>
Public Property Let ElapsedTimeFormat(ByVal a_Value As String)
    If This.ElapsedTimeFormat <> a_Value Then
        This.ElapsedTimeFormat = a_Value
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
Public Property Let ElapsedTimeCaption(ByVal a_Value As String)
    If This.ElapsedTimeCaption <> a_Value Then
        This.ElapsedTimeCaption = a_Value
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
Public Property Let LastErrorMessage(ByVal a_Value As String)
    If This.LastErrorMessage <> a_Value Then
        This.LastErrorMessage = a_Value
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
Public Property Let LastMessage(ByVal a_Value As String)
    If This.LastMessage <> a_Value Then
        This.LastMessage = a_Value
        Me.OnPropertyChanged "LastMessage"
    End If
End Property

''' <summary>   Gets the maximum Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumChannelNumber() As Integer
    MaximumChannelNumber = This.MaximumChannelNumber
End Property

''' <summary>   Sets the maximum Channel Number. </summary>
''' <param name="a_value">   [Integer] the maximum Channel Number value. </value>
Public Property Let MaximumChannelNumber(ByVal a_Value As Integer)
    If This.MaximumChannelNumber <> a_Value Then
        This.MaximumChannelNumber = a_Value
        Me.OnPropertyChanged "MaximumChannelNumber "
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Internal. Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopper()

    Me.ElapsedTimeCaption = VBA.vbNullString
    This.Stopper.Restart

End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Public Sub RestartStopperCommand()

    Me.RestartStopper
    
End Sub

''' <summary>   Internal. Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopper()

    Me.ElapsedTime = This.Stopper.ElapsedMilliseconds

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Public Sub ReadStopperCommand()

    Me.ReadStopper

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
    ' allow the event to take place.
    DoEvents
End Sub

''' <summary>   Notifies all observers of a property change. </summary>
''' <param name="a_propertyName">   [String] the property naMe. </param>
Public Sub OnPropertyChanged(ByVal a_propertyName As String)

    RaiseEvent PropertyChanged( _
            cc_isr_Core.Factory.CreatePropertyChangedEventArgs(a_propertyName))
    ' allow the event to take place.
    DoEvents

End Sub

''' <summary>   Notifies of a pending change in measurement state by
''' raising the <see cref="MeasurementStateChanging"/> event. </summary>
Public Sub OnMeasurementStateChanging()
    
    RaiseEvent MeasurementStateChanging(cc_isr_Core.Factory.NewEventArgs.Initialize())
    ' allow the event to take place.
    DoEvents
    
End Sub


''' <summary>   Handles the change of measurement mode and measuring state and
''' Raises the <see cref="MeasurementStateChanged"/> event. </summary>
''' <param name="a_mode">        [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the
'''                              measurement mode option. </param>
''' <param name="a_info">        [String] info about the measurement state. </param>
''' <param name="a_measuring">   [Boolean] <see cref="Measuring"/>. </param>
Public Sub OnMeasurementStateChanged(ByVal a_mode As cc_isr_Tcp_Scpi.MeasurementModeOption, _
    ByVal a_info As String, ByVal a_measuring As Boolean)
    
    ' turn off the flag handling a timer event.
    This.HandlingTimerEvent = False
    
    ' set the measuring state:
    Me.Measuring = a_measuring
    
    ' set the measurement mode.
    Me.MeasurementMode = a_mode

    ' update the monitoring state
    Me.MonitoringState = a_info

    ' turn on the stop requested in case the timer is still active except for monitoring mode.
    Me.StopRequested = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring <> a_mode

    ' turn off the pause requested when monitoring started
    Me.PauseRequested = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring <> a_mode

    ' disable the stop monitoring action
    Me.StopMonitoringExecutable = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring = a_mode
    
    ' enable the start monitoring action in external mode
    Me.StartMonitoringExecutable = cc_isr_Tcp_Scpi.MeasurementModeOption.External = a_mode
    
    ' enable immediate readings in immediate mode
    Me.MeasureExecutable = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate = a_mode
    
    ' enable toggling the configurations
    Me.ImmediateTriggerOptionExecutable = a_mode <> cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
    Me.ExternalTriggerOptionExecutable = a_mode <> cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
    
    ' finally, indicates that external trigger monitoring is enabled thus updating the related commands.
    Me.ExternalTrigMonitoringEnabled = a_mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
    
    RaiseEvent MeasurementStateChanged( _
            cc_isr_Tcp_Scpi.Factory.NewStateChangedEventArgs.Initialize())
    ' allow the event to take place.
    DoEvents

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' TIMER CONTROL
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the external trigger monitoring timer event using the
''' device to handle the timing events. </summary>
''' <remarks>
''' Use this method to check if turn on or off manual triggering mode of the instrument. <para>
''' 10/16/00  David Hary  1.00.00 Add a beep </para>
''' 2023-06-23 David Hary 1.0.8574 Use the device event handler to update the reading values.
''' </remarks>
Public Sub HandleTimerEvent()

    Const p_procedureName = "HandleTimerEvent"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.PauseRequested Or This.HandlingTimerEvent Then Exit Sub
   
    ' prevent reentry
    This.HandlingTimerEvent = True

    Dim p_channelNumber As Integer

    ' check if scan or single
    If This.AutoIncrementChannelNoEnabled Then
  
        ' in auto increment mode, use the Target Channel Number
        p_channelNumber = This.TargetChannelNumber
      
    Else

        ' in non auto increment mode, use the Target Channel Number
        p_channelNumber = This.SelectedChannelNumber

    End If
  
    ' allow other events to occur
    VBA.DoEvents

    Dim p_isDone As Boolean: p_isDone = This.StopRequested
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True

    p_success = This.K2700.TryMonitorExternalTrigger(p_channelNumber, p_isDone, p_details)
    
    ' report an error on failure.
    If Not p_success Then
        cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " failed handling timer event; " & p_details
        Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_isDone Then
    
        ' if done, request the pause of the timer.
        Me.PauseTimer
        
        ' if not using the timer, go to none mode.
        If Not This.TimerControlled Then _
            Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, "Trigger monitoring done", False
        
#If False Then
        
        ' this is done in the K2700 handler when done.

        ' reset known state.
        p_success = This.Device.TryResetClear(p_details)
        
        ' report an error on failure.
        If Not p_success Then
            cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " failed resetting known state after handling timer event is done; " & p_details
            Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        End If
        
        If p_success Then
        
            ' clear execution state.
            p_success = This.Device.TryClearExecutionState(p_details)
            
            ' report an error on failure.
            If Not p_success Then
                cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " failed clearing execution state after handling timer event is done; " & p_details
                Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
            End If
        
        End If
#End If
        
    ElseIf Not p_success Then
    
        ' if monitoring failed, request a stop and report.
        Me.StopRequested = True
        
    End If

    ' release reentry prevention
    This.HandlingTimerEvent = False

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' inidicate a failure to elicitate a stop request
    p_success = False
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Starts the event timer. </summary>
''' <param name="a_interval">   [Optiona, Long, 500] Specifies the timer interval in milliseconds. </param>
Public Sub StartTimer(Optional ByVal a_interval As Long = 500)

    This.HandlingTimerEvent = False
    This.PauseRequested = False
    This.TimerStarted = True
    Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, "Trigger monitoring commenced", True
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
    DoEvents
    If This.PauseRequested Then
        This.TimerStarted = False
        If Not m_eventTimer Is Nothing Then m_eventTimer.PauseTimer
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, "Trigger monitoring done", False
    Else
        Me.HandleTimerEvent
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
    If This.Session Is Nothing Then
        Set Connectable = Nothing
    Else
        Set Connectable = This.Session.Socket
    End If
End Property

''' <summary>   Returns the connection state of the <see cref="cc_isr_Winsock.IConnectable"/>. </summary>
''' <value>   [Boolean] True if the <see cref="cc_isr_Winsock.IConnectable"/> is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean

    If Me.Connectable Is Nothing Then
        Connected = False
    Else
        Connected = Me.Connectable.Connected
    End If

End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean

    If Me.Connectable Is Nothing Then
        CanConnect = False
    Else
        CanConnect = Me.Connectable.CanConnect
    End If

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Restore Initial State
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initialize the state of the view model executables.  </summary>
Friend Sub InitializeExecutablesState()

    Me.OpenConnectionExecutable = Not Me.Connected
    Me.CloseConnectionExecutable = Me.Connected
    
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
    Me.StartMonitoringExecutable = False
    Me.StopMonitoringExecutable = False

    If This.Session.GpibLanControllerAttached Then
        Me.SelectiveDeviceClearExecutable = Me.Connected
        Me.GoToLocalExecutable = Me.Connected
        Me.LocalLockoutExecutable = Me.Connected
        Me.AutoAssertTalkSetExecutable = Me.Connected
        Me.AutoAssertTalkGetExecutable = Me.Connected
        Me.ReadStatusByteExecutable = Me.Connected
        Me.ServiceRequestedExecutable = Me.Connected
        Me.GpibAddressGetExecutable = Me.Connected
        Me.GpibAddressSetExecutable = Me.Connected
        Me.GpibLanReadTimeoutGetExecutable = Me.Connected
        Me.GpibLanReadTimeoutSetExecutable = Me.Connected
    End If
    
    ' set the connection executables.
    Me.ConnectionExecutableSetter

End Sub

''' <summary>   Read the Tcp Session state. </summary>
''' <remarks>   This method updates the View Model with the Tcp Session State
''' As follows:
''' <see cref="TcpSession.GpibAddress"/>,
''' This method assumes that the following information is already set when connecting:
''' <see cref="TcpSession.ReadTimeout"/>, <see cref="TcpSession.AppendTermination"/>,
''' <see cref="TcpSession.EndOfIdentity"/>, <see cref="TcpSession.AutoAssertTalk"/>,
''' <see cref="TcpSession.LocalLockout"/>.
''' </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
Public Function TryReadSessionState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryReadSessionState"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    
    If p_success Then
    
        Dim p_reading As String: p_reading = This.Session.GpibAddressGetter
        
        p_success = 0 < This.Session.PrimaryGpibAddress
        If p_success Then
            This.GpibAddress = This.Session.PrimaryGpibAddress
        Else
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                " failed reading GPIB Address: '" & p_reading & "'."
        End If
    
    End If
    
    TryReadSessionState = p_success

End Function

''' <summary>   Process the initial connection instrument actions. </summary>
''' <param name="a_errors">         [Out, String] contains any error information that existed before this
'''                                 method. </param>
''' <param name="a_methoderrors">   [Out, String] contains any error due to this method. </param>
''' <returns>   [Boolean] true if no errors are reported. </returns>
Friend Function TryConfigureInstrumentInitialState(ByRef a_errors As String, _
    ByRef a_methodErrors As String) As Boolean

    ' report the socket number
    Me.SocketAddress = This.Session.SocketAddress
    Me.SocketId = This.Session.Socket.SocketId
    
    ' report device errors if any before clearing states.
    
    Dim p_hasErrors As Boolean
    a_errors = This.Device.EnumerateErrors()
    p_hasErrors = VBA.vbNullString <> a_errors
    If p_hasErrors Then _
        a_errors = "Error(s) occurred on session connection: " & VBA.vbCrLf & a_errors

    Dim p_success As Boolean
    Dim p_details As String
    
    ' *RST is issued first in case the instrument is in external trigger mode
    ' where *OPC? fails (on the 2700 instrument).
    ' this is done on configuration using the TryResetClear method.
    ' p_success = This.Device.Try Reset Known State(p_details)
    ' p_success = This.Session.Try Selective Device Clear(p_details, True)
    ' p_success = This.Device.Try Clear Execution State(p_details)
    
    Me.TopCard = This.K2700.RouteSystem.CardName(1)
    Me.BottomCard = This.K2700.RouteSystem.CardName(2)
    Me.ChannelCount = This.K2700.RouteSystem.ChannelCount
    
    ' initialize the scan list for immediate trigger.
    If VBA.vbNullString <> This.ImmediateSenseFunctionName Then
    
        ' set the instrument function scan list
        This.K2700.RouteSystem.DefineFunctionScanLists This.ImmediateSenseFunctionName
    
        ' show the results.
        Me.TopCardFunctionScanList = This.K2700.RouteSystem.FunctionScanList(1)
        Me.BottomCardFunctionScanList = This.K2700.RouteSystem.FunctionScanList(2)
        
    End If

    ' configure the instrument for continuous mode.
    
    Me.ConfigureIntitialTriggerMode
    
    ' report any method errors.
    
    Dim p_hasMethodErrors As Boolean
    a_methodErrors = This.Device.EnumerateErrors()
    p_hasMethodErrors = VBA.vbNullString <> a_methodErrors
    If p_hasMethodErrors Then _
        a_methodErrors = "Error(s) occurred on instrument connection: " & VBA.vbCrLf & a_methodErrors
    
    TryConfigureInstrumentInitialState = Not (p_hasMethodErrors Or p_hasErrors)
    
End Function

''' <summary>   Configures the 2700 Multimeter scanner initial state. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State. </remarks>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            failed to restore state. </param>
Public Function TryConfigureInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"

    Dim p_success As Boolean: p_success = True
    
    If p_success Then
    
        Dim a_errors As String
        Dim a_methodErrors As String
        p_success = TryConfigureInstrumentInitialState(a_errors, a_methodErrors)
        If Not p_success Then
            If VBA.vbNullString = a_errors Then
                a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                    "; " & a_methodErrors
            ElseIf VBA.vbNullString = a_methodErrors Then
                a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                    "; " & a_errors
            Else
                a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                    "; " & a_errors & VBA.vbCrLf & a_methodErrors
            End If
            Me.LastErrorMessage = a_details
        End If
    
    End If
    
    TryConfigureInitialState = p_success

End Function

''' <summary>   Returns true and a message if the specified sense function matched neither
''' one of sense functions of the defferent measurement modes. </summary>
''' <param name="a_senseFunctionName">   [String] The actual sense function name. </param>
''' <param name="a_details">         [String] Details why the sense function needs to be restored. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreSenseFunction(ByVal a_senseFunctionName As String, _
    ByRef a_details As String) As Boolean

    If Not ((1 = InStr(1, This.ContinuousSenseFunctionName, a_senseFunctionName, vbTextCompare)) Or _
            (1 = InStr(1, This.ImmediateSenseFunctionName, a_senseFunctionName, vbTextCompare)) Or _
            (1 = InStr(1, This.ExternalSenseFunctionName, a_senseFunctionName, vbTextCompare))) Then
        a_details = "The K2700 device sense function '" & a_senseFunctionName & _
            "' is not expected; it must be either " & _
            This.ContinuousSenseFunctionName & " or " & _
            This.ImmediateSenseFunctionName & " or " & _
            This.ExternalSenseFunctionName & "."
        ShouldRestoreSenseFunction = True
    Else
        ShouldRestoreSenseFunction = False
    End If

End Function

''' <summary>   Returns true and a message if the view model state does not match the expected restore state. </summary>
''' <remarks>   This is the expected state for making resistance measurements. In future releases
''' we could define a set of states. </remarks>
''' <param name="a_details">   [String] Details why the view model needs to be resotred
'''                            to its Initial State. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreInitialState(ByRef a_details As String) As Boolean

    If This.K2700 Is Nothing Then
        a_details = "The K2700 device instance is nothing."
        ShouldRestoreInitialState = True
    ElseIf This.K2700.ShouldRestoreInitialState(a_details) Then
        ShouldRestoreInitialState = True
    Else
        ' validate the sense function
        Dim p_senseFunctionName As String
        p_senseFunctionName = This.K2700.SenseSystem.SenseSystem.SenseFunctionGetter()
        ShouldRestoreInitialState = Me.ShouldRestoreSenseFunction(p_senseFunctionName, a_details)
    End If

End Function

''' <summary>   Restores the instrument to its Initial State in case the instrument power
''' was toggled or a recovery from error is needed. </summary>
''' <remarks>   This method attempts to restore the instrument to a Initial State in case it is
''' powered off while the program is active as follows:
''' - The user defined errors are cleared;
''' - The View Model is initialized if the instrument is nothing;
''' - The View Model is connected if it is not connected;
''' - The GPIB-Lan device is restored to its Initial State;
''' - The instrument is set to its Initial State if sense function needs to be restored;
'''     - otherwise, the instrument is set to its clear Initial State.
''' </remarks>
''' <param name="a_details">   [String] Details why the view model needs to be resotred
'''                            to its Initial State. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryRestoreInitialState(ByRef a_details As String)

    Const p_procedureName As String = "TryRestoreInitialState"

    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString

    If p_success And This.K2700 Is Nothing Then Me.Initialize
    
    p_success = Not (Me.Device Is Nothing Or Me.Session Is Nothing Or Me.Session.Socket Is Nothing)
    
    If Not p_success Then
        p_details = "Unable to restore the view model to its initial state because the view model failed to initialize the K2700 instance. Please restart the program."
    End If
    
    If p_success And Not (Me.Connected Or This.Session.Socket.CanConnect) Then
        
        ' if not connected and not 'can connect' it means that the
        ' socket was disposed and needs to be rebuilt.
        Me.Initialize
            
        p_success = This.Session.Socket.CanConnect
        
        If Not p_success Then
            p_details = "Failed initializing the socket; the socket id is invalid ('" & _
                VBA.CStr(cc_isr_Winsock.ws32_INVALID_SOCKET) & "')."
        End If
    
    End If
    
    If p_success Then
        p_success = This.Device.TryRestoreInitialState(p_details)
        If Not p_success Then _
            p_details = "The IEEE488 device failed restoring its initial state; " & p_details
    End If
    
    If p_success Then
        ' set the executable condition first, which might
        ' change depending on how the instrument connected methods affects the executable commands.
        Me.InitializeExecutablesState
    End If
    
    If p_success Then
        ' reset the IEEE488 device state reading the operation completion reply of '1'.
        ' this is done on try to configure initial state.
        ' p_success = This.Device.TryResetClear(p_details)
        ' If Not p_success Then _
        '    p_details = "The IEEE488 device failed reseting its known state; " & p_details
    End If

    If p_success Then
        ' configure the initial state
        p_success = Me.TryConfigureInitialState(p_details)
    End If

    If p_success Then
        ' clear the instrument execution state to remove any lingering errors
        p_success = Me.TryClearExecutionState(p_details)
    End If
    
    If Not p_success Then
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed restoring initial state; " & p_details
    End If
    
    TryRestoreInitialState = p_success
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Socket Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Winsock.IPv4StreamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event arguments. </param>
Private Sub m_socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName = "m_socket_ConnectionChanged"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' set the executable condition first, which might
    ' change depending on how the instrument connected methods affects the executable commands.
    Me.InitializeExecutablesState
    
    If a_eventArgs.Connected Then
    
        Dim p_details As String
         
        ' configure the initial state
        If Not Me.TryConfigureInitialState(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed configuring the initial state; " & p_details
    
        ' read the session state
        If Not Me.TryReadSessionState(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed reading the session state; " & p_details
    
    Else
    
        ' we do not wish to remove a user settings.
        ' Me.SocketAddress = VBA.vbNullString
        Me.SocketId = 0
    
    End If
    
    ' notify of connection state.
    Me.ConnectionExecutableSetter
    
    Me.OnMeasurementStateChanged Me.MeasurementMode, "connection changed", Me.Measuring

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4StreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event arguments. </param>
Private Sub m_socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    
    Const p_procedureName = "m_k2700_ConnectionChanging"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
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
    
    ' issue a beep if reading beep is enabled
    If This.ReadingBeepEnabled Then VBA.Beep

    ' notify the observer of a channel measurements
    If VBA.vbNullString = a_eventArgs.Reading Then
        OnChannelMeasured a_eventArgs.ChannelNumber, -1#
    Else
        OnChannelMeasured a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
    End If

    If This.AutoIncrementChannelNoEnabled Then

        ' increment the resistance for the next reading
        
        Me.TargetChannelNumber = IIf(a_eventArgs.ChannelNumber < Me.ChannelCount, _
            a_eventArgs.ChannelNumber + 1, 1)

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Class events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub Class_Initialize()
    
    ' view model properties
    This.ElapsedTimeFormat = "0.0"
    
    ' Winsock properties
    This.SessionTimeout = 3000
   
    ' 2700 properties
    This.ContinuousSenseFunctionName = "FRES"
    This.ImmediateSenseFunctionName = "RES"
    This.ExternalSenseFunctionName = "FRES"
    This.MeasuredValueFormat = "0.0"
    This.GpibLanControllerPort = 1234
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    This.SelectedChannelNumber = 1
    
    ' service request monitoring timer control
    This.TimerInterval = 500
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700()
    Set This.Device = This.K2700.Device
    Set This.Session = This.Device.Session
    Set This.Stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    
    This.TopCard = VBA.vbNullString
    
End Sub
