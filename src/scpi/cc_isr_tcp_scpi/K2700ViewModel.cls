VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700ViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A public creatable View Model for the
'''             2700 instrument control. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Represents an observer pattern subject that can notify observers
'''             when a property value changes. </summary>
Implements cc_isr_Core.INotifyPropertyChanged

''' <summary>   Notifies all registered handlers of a property value change. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Core.PropertyChangedEventArgs"/> event args. </param>
Public Event PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

''' <summary>   The event that is raised upon completion of a DUT measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Tcp_Scpi.DutReadingEventArgs"/> event args. </param>
Public Event DutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

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
    ScpiCommands As Collection
    
    ' GPIB-Lan properties.
    ServiceRequested As Boolean
    SerialPollByte As Integer
    SerialPollExecutable As Boolean
    GpibAddress As Integer
    GpibAddressReading As String
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
    DutCount As Integer
    MaximumDutCount As Integer
    SelectedDutNumber As Integer
    TargetDutNumber As Integer
    MeasuredDutNumber As Integer
    FrontInputsSenseFunctionName As String
    RearInputsSenseFunctionName As String
    ImmediateFrontInputsRequired As Boolean
    ExternalFrontInputsRequired As Boolean
    SenseFunctionName As String
    SenseFunctionCaption As String
    SenseFunctionUnit As String
    SenseFunctionUnicode As Integer
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

    ' Service request monitoring properties
    MonitoringState As String
    AutoIncrementDutNumberEnabled As Boolean
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

End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Construction and cleanup
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub Class_Initialize()
    
    ' view model properties
    This.ElapsedTimeFormat = "0.0"
    
    ' Winsock properties
    This.SessionTimeout = 3000
   
    ' 2700 properties
    This.RearInputsSenseFunctionName = "RES"
    This.FrontInputsSenseFunctionName = "FRES"
    
    This.ImmediateFrontInputsRequired = False
    This.ExternalFrontInputsRequired = True
    
    This.MeasuredValueFormat = "0.0"
    This.GpibLanControllerPort = 1234
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    This.SelectedDutNumber = 1
    
    ' service request monitoring timer control
    This.TimerInterval = 500
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700()
    Set This.Device = This.K2700.Device
    Set This.Session = This.Device.Session
    Set This.Stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    
    This.TopCard = VBA.vbNullString
    
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

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

' + + + + + + + + + + + + + + + + + + + + + + + + +
' User Interface Control
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <remarks>   True if instrument beep sound on each reading is enabled. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ReadingBeepEnabled() As Boolean
    ReadingBeepEnabled = This.ReadingBeepEnabled
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <param name="a_value">   [Boolean] instrument beep sound on each reading is enabled. </value>
Friend Property Let ReadingBeepEnabled(ByVal a_value As Boolean)
    If This.ReadingBeepEnabled <> a_value Then
        This.ReadingBeepEnabled = a_value
        Me.OnPropertyChanged "ReadingBeepEnabled"
    End If
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Double]. </value>
Public Property Get ElapsedTime() As Double
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Double] the Elapsed Time in milliseconds. </value>
Friend Property Let ElapsedTime(ByVal a_value As Double)
    If This.ElapsedTime <> a_value Then
        This.ElapsedTime = a_value
        Me.OnPropertyChanged "ElapsedTime"
        Me.ElapsedTimeCaption = VBA.Format$(a_value, This.ElapsedTimeFormat)
    End If
End Property

''' <summary>   Gets the Elapsed Time Format. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeFormat() As String
    ElapsedTimeFormat = This.ElapsedTimeFormat
End Property

''' <summary>   Sets the Elapsed Time Format. </summary>
''' <param name="a_value">   [String] the elapsed time format. </value>
Friend Property Let ElapsedTimeFormat(ByVal a_value As String)
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
Friend Property Let ElapsedTimeCaption(ByVal a_value As String)
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
Friend Property Let LastErrorMessage(ByVal a_value As String)
    If This.LastErrorMessage <> a_value Then
        This.LastErrorMessage = a_value
        Me.OnPropertyChanged "LastErrorMessage"
    End If
End Property

''' <summary>   Reports an error. </summary>
''' <param name="a_value">   [String] the last error message value. </value>
Public Sub OnError(ByVal a_value As String)
    Me.LastErrorMessage = a_value
End Sub

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = This.LastMessage
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the last message value. </value>
Friend Property Let LastMessage(ByVal a_value As String)
    If This.LastMessage <> a_value Then
        This.LastMessage = a_value
        Me.OnPropertyChanged "LastMessage"
    End If
End Property

''' <summary>   Clears the last error message and the Last Message. </summary>
''' <param name="a_lastErrorMessage">   [Optional, String, Empty] the last error message value. </value>
''' <param name="a_lastMessage">        [Optional, String, Empty] the last message value. </value>
Public Sub ClearMessages(Optional ByVal a_lastErrorMessage As String = VBA.vbNullString, _
    Optional ByVal a_lastMessage As String = VBA.vbNullString)
    Me.LastErrorMessage = a_lastErrorMessage
    Me.LastMessage = a_lastMessage
End Sub

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

' + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock settings, commands and executables
' + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock: Socket address, port and host
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the has valid socket address condition indicating that the socket address
''' was parsed and has the expected port numbers. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasValidSocketAddress() As Boolean
    HasValidSocketAddress = This.HasValidSocketAddress
End Property

''' <summary>   Sets the has valid socket address condition indicating that the socket address
''' was parsed and has the expected port numbers. </summary>
''' <param name="a_value">   [Boolean] the value. </param>
Friend Property Let HasValidSocketAddress(ByVal a_value As Boolean)
    If This.HasValidSocketAddress <> a_value Then
        This.HasValidSocketAddress = a_value
        Me.OnPropertyChanged "HasValidSocketAddress"
    End If
End Property

''' <summary>   Parses the current socket address and sets the host and port values. </summary>
''' <param name="a_value">   [String] the socket address. </param>
Friend Sub ParseSocketAddress(ByVal a_value As String)
    Dim p_hasValidSocketAddress As Boolean
    p_hasValidSocketAddress = False
    This.SocketAddress = a_value
    This.Host = VBA.vbNullString
    This.Port = cc_isr_Winsock.ws32_INVALID_SOCKET
    If cc_isr_Winsock.Winsock.TryParseAddress(This.SocketAddress, This.Host, This.Port, Me.LastErrorMessage) Then _
        p_hasValidSocketAddress = VBA.vbNullString <> This.Host _
            And (5025 = This.Port Or 1234 = This.Port)
    Me.OnPropertyChanged "Port"
    Me.OnPropertyChanged "Host"
    Me.OnPropertyChanged "SocketAddress"
    Me.HasValidSocketAddress = p_hasValidSocketAddress
End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Sets the Host dotted IP address. </summary>
''' <param name="a_value">   [String] the host dotted IP address. </value>
Friend Property Let Host(ByVal a_value As String)
    If This.Host <> a_value Then
        This.Host = a_value
        Me.OnPropertyChanged "Host"
    End If
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Long
    Port = This.Port
End Property

''' <summary>   Sets the port number. </summary>
''' <param name="a_value">   [Long] the port number. </value>
Friend Property Let Port(ByVal a_value As Long)
    If This.Port <> a_value Then
        This.Port = a_value
        Me.OnPropertyChanged "Port"
    End If
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket dotted address and port number. </summary>
''' <param name="a_value">   [String] the socket dotted address and
'''                          port number, e.g., <c>192.168.0.252:1234</c>. </value>
Friend Property Let SocketAddress(ByVal a_value As String)
    If This.SocketAddress <> a_value Then
        Me.ParseSocketAddress a_value
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
        Me.OnPropertyChanged "SocketId"
    End If
End Property

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Long
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket receive timeout in milliseconds. </summary>
''' <param name="a_value">   [Long] the timeout value in milliseconds. </value>
Friend Property Let SessionTimeout(ByVal a_value As Long)
    If This.SessionTimeout <> a_value Then
        This.SessionTimeout = a_value
        Me.OnPropertyChanged "SessionTimeout"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' Winsock Connection Manangement
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Opens the connection. </summary>
''' <param name="a_socketAddress">    [String] the socket dotted address and
'''                                   port number, e.g., <c>192.168.0.252:1234</c>. </value>
''' <param name="a_sessionTimeout">   [Integer] the timeout value in milliseconds. </value>
Public Sub OpenConnection(ByVal a_socketAddress As String, ByVal a_sessionTimeout As Long)

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
    p_success = cc_isr_Ieee488.GpibLanController.TryPing(a_socketAddress, a_sessionTimeout, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed pinging the instrument; " & p_details
    
    Me.LastMessage = "Opening connection to " & a_socketAddress & "; please wait..."
    
    p_success = Me.Connectable.TryOpenConnection(a_socketAddress, a_sessionTimeout, p_details)
    If Not p_success Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Failed opening the connection; " & p_details

    If p_success Then
        Me.SocketAddress = a_socketAddress
        Me.SessionTimeout = a_sessionTimeout
        Me.LastMessage = "Socket opened at " & a_socketAddress & "."
    End If
    
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
''' <param name="a_socketAddress">    [String] the socket dotted address and
'''                                   port number, e.g., <c>192.168.0.252:1234</c>. </value>
''' <param name="a_sessionTimeout">   [Integer] the timeout value in milliseconds. </value>
Public Sub OpenConnectionCommand(ByVal a_socketAddress As String, ByVal a_sessionTimeout As Long)

    Const p_procedureName = "OpenConnectionCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Me.OpenConnection a_socketAddress, a_sessionTimeout
    
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

''' <summary>   Gets the condition indicating that the Open Connection command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Open Connection command can be executed. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = This.OpenConnectionExecutable
End Property

''' <summary>   Sets the condition indicating that the Open Connection command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Open Connection command can be executed. </param>
Friend Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    If This.OpenConnectionExecutable <> a_value Then
        This.OpenConnectionExecutable = a_value
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
Friend Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    If This.CloseConnectionExecutable <> a_value Then
        This.CloseConnectionExecutable = a_value
        Me.OnPropertyChanged "CloseConnectionExecutable"
        Me.OnPropertyChanged "ToggleConnectionExecutable"
    End If
End Property

''' <summary>   Sets the condition indicating that a Connection command
'''             can be executed. </summary>
Friend Sub ConnectionExecutableSetter()
    Me.OpenConnectionExecutable = Not Me.Connected
    Me.CloseConnectionExecutable = Me.Connected
    
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

''' <summary>   Tries to issue power on reset. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' Winsock I/O
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Sent Message. </summary>
''' <value>   [String]. </value>
Public Property Get SentMessage() As String
    SentMessage = This.SentMessage
End Property

''' <summary>   Sets the Sent Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Friend Property Let SentMessage(ByVal a_value As String)
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
Friend Property Let ReceivedMessage(ByVal a_value As String)
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
Friend Property Let ReceivedMessageLength(ByVal a_value As Integer)
    If This.ReceivedMessageLength <> a_value Then
        This.ReceivedMessageLength = a_value
        Me.OnPropertyChanged "ReceivedMessageLength"
    End If
End Property

''' <summary>   Query the INET Socket. </summary>
''' <param name="a_value">   [String] the Message to Send. </value>
Public Sub QuerySocketCommand(ByVal a_value As String)

    Const p_procedureName = "QuerySocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If VBA.Len(a_value) = 0 Then Exit Sub

    Dim p_receivedMessage As String
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer

    Me.RestartStopper

    p_length = This.Session.Socket.TryQuery(a_value & This.Session.Termination, _
        p_receivedMessage, p_details, This.Session.ReadAfterWriteDelay)
    p_success = p_length > 0
    If p_success Then
        Me.ReceivedMessage = p_receivedMessage
        Me.ReceivedMessageLength = p_length
        Me.SentMessage = a_value
    End If
    Me.LastErrorMessage = p_details
    
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

''' <summary>   Read from the INET socket. </summary>
Public Sub ReadSocketCommand()

    Const p_procedureName = "ReadSocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_receivedMessage As String
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer

    Me.RestartStopper

    If 0 > This.Session.Socket.TryReceive(Me.ReceivedMessage, p_details) Then
        Me.LastErrorMessage = p_details
    End If

    p_length = This.Session.Socket.TryReceive(p_receivedMessage, p_details)
    p_success = p_length > 0
    If p_success Then
        Me.ReceivedMessage = p_receivedMessage
        Me.ReceivedMessageLength = p_length
    End If
    Me.LastErrorMessage = p_details

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
''' <param name="a_value">   [String] the Message to Send. </value>
Public Sub WriteSocketCommand(ByVal a_value As String)

    Const p_procedureName = "WriteSocketCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If VBA.Len(a_value) = 0 Then Exit Sub

    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer

    Me.RestartStopper

    p_length = This.Session.Socket.TrySendMessage(a_value & This.Session.Termination, p_details, _
        This.Session.ReadAfterWriteDelay)
        
    p_success = p_length > 0

    Me.ReadStopper

    If p_success Then
    
        Me.SentMessage = a_value
    
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

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Tcp Session Settings, Commands and Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.TcpSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.TcpSession"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.TcpSession
    Set Session = This.Session
End Property

''' <summary>   Gets the Read After Write Delay in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = This.ReadAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Read After Write Delay Delay in milliseconds. </param>
Public Property Let ReadAfterWriteDelay(ByVal a_value As Integer)
    If This.ReadAfterWriteDelay <> a_value Then
        This.ReadAfterWriteDelay = a_value
        Me.OnPropertyChanged "ReadAfterWriteDelay"
        If Not This.K2700 Is Nothing Then _
            This.K2700.ReadAfterWriteDelay = a_value
    End If
End Property

''' <summary>   Gets the termination to append to output messages. </summary>
''' <value>   [String]. </value>
Public Property Get Termination() As String
    Termination = This.Termination
End Property

''' <summary>   Sets the termination to append to output messages. </summary>
''' <param name="a_value">   [String] the termination. </param>
Public Property Let Termination(ByVal a_value As String)
    If This.Termination <> a_value Then
        This.Termination = a_value
        Me.OnPropertyChanged "Termination"
        If Not This.K2700 Is Nothing Then _
            This.K2700.Termination = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 properties, commands and executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.Device"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.Device"/>. </value>
Public Property Get Device() As cc_isr_Ieee488.Device
    Set Device = This.Device
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: Status Byte
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Status Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StatusByte() As Integer
    StatusByte = This.StatusByte
End Property

''' <summary>   Sets the Status Byte. </summary>
''' <param name="a_value">   [Integer] the Status Byte value. </value>
Friend Property Let StatusByte(ByVal a_value As Integer)
    If This.StatusByte <> a_value Then
        This.StatusByte = a_value
        Me.OnPropertyChanged "StatusByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
                a_value And cc_isr_Ieee488.ServiceRequestFlags.RequestingService
    End If
End Property

''' <summary>   Tries to read (<c>*STB?</c>) the device status byte. </summary>
''' <remarks>   This method updates the <see cref="K2700ViewModel.StatusByte"/>. </remarks>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean] true if success. </returns>
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

''' <summary>   Gets the condition indicating that the Query Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Status Byte command can be executed. </value>
Public Property Get QueryStatusByteExecutable() As Boolean
    QueryStatusByteExecutable = This.QueryStatusByteExecutable
End Property

''' <summary>   Sets the condition indicating that the Query Status Byte command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query Status Byte command can be executed. </param>
Friend Property Let QueryStatusByteExecutable(ByVal a_value As Boolean)
    If This.QueryStatusByteExecutable <> a_value Then
        This.QueryStatusByteExecutable = a_value
        Me.OnPropertyChanged "QueryStatusByteExecutable"
    End If
End Property


' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: Standard Byte
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    StandardByte = This.StandardByte
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Friend Property Let StandardByte(ByVal a_value As Integer)
    If This.StandardByte <> a_value Then
        This.StandardByte = a_value
        Me.OnPropertyChanged "StandardByte"
    End If
End Property

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

''' <summary>   Gets the condition indicating that the Read Standard Event Status command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Standard Event Status command can be executed. </value>
Public Property Get ReadStandardEventStatusExecutable() As Boolean
    ReadStandardEventStatusExecutable = This.ReadStandardEventStatusExecutable
End Property

''' <summary>   Sets the condition indicating that the Read Standard Event Status command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read Standard Event Status command can be executed. </param>
Friend Property Let ReadStandardEventStatusExecutable(ByVal a_value As Boolean)
    If This.ReadStandardEventStatusExecutable <> a_value Then
        This.ReadStandardEventStatusExecutable = a_value
        Me.OnPropertyChanged "ReadStandardEventStatusExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: SCPI Commands
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Scpi Commands. </summary>
''' <value>   [Collection]. </value>
Public Property Get ScpiCommands() As Collection
    Set ScpiCommands = This.ScpiCommands
End Property

''' <summary>   Sets the Scpi Commands. </summary>
''' <param name="a_value">   [Collection] the Scpi Commands value. </value>
Friend Property Let ScpiCommands(ByVal a_value As Collection)
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: I/O
' . . . . . . . . . . . . . . . . . . . . . . . . .

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: I/O: Query
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Query the instrument using the message to send and update then elapsed tiMe. </summary>
''' <param name="a_value">        [String] the Message to Send. </value>
Public Sub QueryCommand(ByVal a_value As String)

    Const p_procedureName = "QueryCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If VBA.Len(a_value) = 0 Then Exit Sub

    Dim p_receivedMessage As String
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer

    Me.RestartStopper

    p_length = This.Session.TryQueryLine(a_value, p_receivedMessage, p_details)

    Me.ReadStopper

    p_success = p_length > 0

    If p_success Then
    
        Me.SentMessage = a_value
        Me.ReceivedMessage = p_receivedMessage
        Me.ReceivedMessageLength = p_length
    
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

''' <summary>   Gets the condition indicating that the Query command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query command can be executed. </value>
Public Property Get QueryExecutable() As Boolean
    QueryExecutable = This.QueryExecutable
End Property

''' <summary>   Sets the condition indicating that the Query command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query command can be executed. </param>
Friend Property Let QueryExecutable(ByVal a_value As Boolean)
    If This.QueryExecutable <> a_value Then
        This.QueryExecutable = a_value
        Me.OnPropertyChanged "QueryExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: I/O: Read
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Read from the instrument updating the elap[es times. </summary>
Public Sub ReadCommand()

    Const p_procedureName = "ReadCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_receivedMessage As String
    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer
    
    Me.RestartStopper

    p_length = This.Session.TryRead(p_receivedMessage, p_details)
    
    Me.ReadStopper

    p_success = p_length > 0
    If p_success Then
    
        Me.ReceivedMessage = p_receivedMessage
        Me.ReceivedMessageLength = p_length
    
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

''' <summary>   Gets the condition indicating that the Read command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read command can be executed. </value>
Public Property Get ReadExecutable() As Boolean
    ReadExecutable = This.ReadExecutable
End Property

''' <summary>   Sets the condition indicating that the Read command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read command can be executed. </param>
Friend Property Let ReadExecutable(ByVal a_value As Boolean)
    If This.ReadExecutable <> a_value Then
        This.ReadExecutable = a_value
        Me.OnPropertyChanged "ReadExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: I/O: Write
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Write a message to the INET socket and query the status bytes. </summary>
''' <param name="a_value">        [String] the Message to Send. </value>
''' <param name="a_readStatus">   [String] reads server request and standard event
'''                                statuses depending on this being a query message. </value>
Public Sub WriteCommand(ByVal a_value As String, ByVal a_readStatus As Boolean)

    Const p_procedureName = "WriteCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If VBA.Len(a_value) = 0 Then Exit Sub

    Dim p_details As String
    Dim p_success As Boolean
    Dim p_length As Integer

    Me.RestartStopper

    p_length = This.Session.TryWriteLine(a_value, p_details)
    p_success = p_length > 0

    Me.ReadStopper

    If p_success Then
    
        Me.SentMessage = a_value
    
        If a_readStatus Then
    
            Dim p_canQuery As Boolean
            p_canQuery = Not cc_isr_Ieee488.Scpi.IsQuery(a_value)
            
            ' serial poll if this is the GPIB-Lan controller is attached.
            Me.ReadStatusByte False
    
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

''' <summary>   Gets the condition indicating that the Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Write command can be executed. </value>
Public Property Get WriteExecutable() As Boolean
    WriteExecutable = This.WriteExecutable
End Property

''' <summary>   Sets the condition indicating that the Write command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Write command can be executed. </param>
Friend Property Let WriteExecutable(ByVal a_value As Boolean)
    If This.WriteExecutable <> a_value Then
        This.WriteExecutable = a_value
        Me.OnPropertyChanged "WriteExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: Device Managemenr: CLS
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Sends the instrument to is Clear Execution State and returns
''' success or failure. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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

''' <summary>   Gets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Clear Execution State Command can be executed. </value>
Public Property Get ClearExecutionStateExecutable() As Boolean
    ClearExecutionStateExecutable = This.ClearExecutionStateExecutable
End Property

''' <summary>   Sets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Clear Execution State Command can be executed. </param>
Friend Property Let ClearExecutionStateExecutable(ByVal a_value As Boolean)
    If This.ClearExecutionStateExecutable <> a_value Then
        This.ClearExecutionStateExecutable = a_value
        Me.OnPropertyChanged "ClearExecutionStateExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' IEEE 488: Device Managemenr: RST
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Sends the instrument to is reset known State and returns
''' success or failure. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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

''' <summary>   Gets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Reset Known State command can be executed. </value>
Public Property Get ResetKnownStateExecutable() As Boolean
    ResetKnownStateExecutable = This.ResetKnownStateExecutable
End Property

''' <summary>   Sets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Reset Known State command can be executed. </param>
Friend Property Let ResetKnownStateExecutable(ByVal a_value As Boolean)
    If This.ResetKnownStateExecutable <> a_value Then
        This.ResetKnownStateExecutable = a_value
        Me.OnPropertyChanged "ResetKnownStateExecutable"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Settings, Commands, and Executables
' + + + + + + + + + + + + + + + + + + + + + + + + +

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Auto Asser Talk
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Sets the Auto Assert Talk condition (Read After Write Enabled). </summary>
''' <param name="a_value">   [Boolean] True to Auto Assert Talk upon write. </value>
Friend Property Let AutoAssertTalk(ByVal a_value As Boolean)
    If This.AutoAssertTalk <> a_value Then
        This.AutoAssertTalk = a_value
        Me.OnPropertyChanged "AutoAssertTalk"
    End If
End Property

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

''' <summary>   Gets the condition indicating that the auto assert TALK get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the auto assert TALK get command can be executed. </value>
Public Property Get AutoAssertTalkGetExecutable() As Boolean
    AutoAssertTalkGetExecutable = This.AutoAssertTalkGetExecutable
End Property

''' <summary>   Gets the condition indicating that the auto assert TALK get command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the auto assert TALK get command can be executed. </param>
Friend Property Let AutoAssertTalkGetExecutable(ByVal a_value As Boolean)
    If This.AutoAssertTalkGetExecutable <> a_value Then
        This.AutoAssertTalkGetExecutable = a_value
        Me.OnPropertyChanged "AutoAssertTalkGetExecutable"
    End If
End Property

''' <summary>   Handles the Auto Assert Talk Set command event. </summary>
''' <param name="a_value">   [Boolean] True to Auto Assert Talk upon write. </value>
Public Sub AutoAssertTalkSetCommand(ByVal a_value As Boolean)

    Const p_procedureName = "AutoAssertTalkSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    If Not This.AutoAssertTalk Then
        This.Session.AutoAssertTalkSetter a_value
        Me.AutoAssertTalk = a_value
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

''' <summary>   Gets the condition indicating that the auto assert TALK set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the auto assert TALK set command can be executed. </value>
Public Property Get AutoAssertTalkSetExecutable() As Boolean
    AutoAssertTalkSetExecutable = This.AutoAssertTalkSetExecutable
End Property

''' <summary>   Sets the condition indicating that the auto assert TALK set command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the auto assert TALK set command can be executed. </param>
Friend Property Let AutoAssertTalkSetExecutable(ByVal a_value As Boolean)
    If This.AutoAssertTalkSetExecutable <> a_value Then
        This.AutoAssertTalkSetExecutable = a_value
        Me.OnPropertyChanged "AutoAssertTalkSetExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB Address
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Gpib Address reading. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibAddressReading() As String
    GpibAddressReading = This.GpibAddressReading
End Property

''' <summary>   Sets the Gpib Address reading. </summary>
''' <param name="a_value">   [String] the Gpib Address reading. </value>
Friend Property Let GpibAddressReading(ByVal a_value As String)
    If This.GpibAddressReading <> a_value Then
        This.GpibAddressReading = a_value
        Me.OnPropertyChanged "GpibAddressReading"
    End If
End Property

''' <summary>   Gets the Primary Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibAddress() As Integer
    GpibAddress = This.GpibAddress
End Property

''' <summary>   Sets the Primary Gpib Address. </summary>
''' <param name="a_value">   [Integer] the Primary Gpib Address value. </value>
Friend Property Let GpibAddress(ByVal a_value As Integer)
    If This.GpibAddress <> a_value Then
        This.GpibAddress = a_value
        Me.OnPropertyChanged "GpibAddress"
    End If
End Property

''' <summary>   Handles the GPIB Address Get command event. </summary>
Public Sub GpibAddressGetCommand()

    Const p_procedureName = "GpibAddressGetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Me.GpibAddressReading = This.Session.GpibAddressGetter()
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

''' <summary>   Gets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Get Command can be executed. </value>
Public Property Get GpibAddressGetExecutable() As Boolean
    GpibAddressGetExecutable = This.GpibAddressGetExecutable
End Property

''' <summary>   Gets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Gpib Address Get Command can be executed. </param>
Friend Property Let GpibAddressGetExecutable(ByVal a_value As Boolean)
    If This.GpibAddressGetExecutable <> a_value Then
        This.GpibAddressGetExecutable = a_value
        Me.OnPropertyChanged "GpibAddressGetExecutable"
    End If
End Property

''' <summary>   Handles the GPIB Address set command event. </summary>
''' <param name="a_value">   [Integer] the Primary Gpib Address value. </value>
Public Sub GpibAddressSetCommand(ByVal a_value As Integer)

    Const p_procedureName = "GpibAddressSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.GpibAddressSetter a_value
    Me.GpibAddress = a_value

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

''' <summary>   Gets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Set Command can be executed. </value>
Public Property Get GpibAddressSetExecutable() As Boolean
    GpibAddressSetExecutable = This.GpibAddressSetExecutable
End Property

''' <summary>   Sets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Gpib Address Set Command can be executed. </param>
Friend Property Let GpibAddressSetExecutable(ByVal a_value As Boolean)
    If This.GpibAddressSetExecutable <> a_value Then
        This.GpibAddressSetExecutable = a_value
        Me.OnPropertyChanged "GpibAddressSetExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Controller Port
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    If This.GpibLanControllerPort <> a_value Then
        This.GpibLanControllerPort = a_value
        Me.OnPropertyChanged "GpibLanControllerPort"
        If Not This.K2700 Is Nothing Then _
            This.K2700.GpibLanControllerPort = a_value
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Read Timeout
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanReadTimeout() As Integer
    GpibLanReadTimeout = This.GpibLanReadTimeout
End Property

''' <summary>   Sets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Friend Property Let GpibLanReadTimeout(ByVal a_value As Integer)
    If This.GpibLanReadTimeout <> a_value Then
        This.GpibLanReadTimeout = a_value
        Me.OnPropertyChanged "GpibLanReadTimeout"
    End If
End Property

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

''' <summary>   Gets the condition indicating that the Read Timeout Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Get command can be executed. </value>
Public Property Get GpibLanReadTimeoutGetExecutable() As Boolean
    GpibLanReadTimeoutGetExecutable = This.GpibLanReadTimeoutGetExecutable
End Property

''' <summary>   Gets the Read Timeout Get command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Get command can be executed. </param>
Friend Property Let GpibLanReadTimeoutGetExecutable(ByVal a_value As Boolean)
    If This.GpibLanReadTimeoutGetExecutable <> a_value Then
        This.GpibLanReadTimeoutGetExecutable = a_value
        Me.OnPropertyChanged "GpibLanReadTimeoutGetExecutable"
    End If
End Property

''' <summary>   Handles the Read Timeout set command event. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Sub GpibLanReadTimeoutSetCommand(ByVal a_value As Integer)

    Const p_procedureName = "GpibLanReadTimeoutSetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    This.Session.ReadTimeoutSetter a_value
    Me.GpibLanReadTimeout = a_value

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

''' <summary>   Gets the condition indicating that the Read Timeout Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Set command can be executed. </value>
Public Property Get GpibLanReadTimeoutSetExecutable() As Boolean
    GpibLanReadTimeoutSetExecutable = This.GpibLanReadTimeoutSetExecutable
End Property

''' <summary>   Sets the Read Timeout Set command condition. </summary>
''' <param name="a_value">   [Boolean] True if the Read Timeout Set command can be executed. </param>
Friend Property Let GpibLanReadTimeoutSetExecutable(ByVal a_value As Boolean)
    If This.GpibLanReadTimeoutSetExecutable <> a_value Then
        This.GpibLanReadTimeoutSetExecutable = a_value
        Me.OnPropertyChanged "GpibLanReadTimeoutSetExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Serial Poll
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Serial Poll Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get SerialPollByte() As Integer
    SerialPollByte = This.SerialPollByte
End Property

''' <summary>   Sets the Serial Poll Byte. </summary>
''' <param name="a_value">   [Integer] the Serial Poll Byte value. </value>
Friend Property Let SerialPollByte(ByVal a_value As Integer)
    If This.SerialPollByte <> a_value Then
        This.SerialPollByte = a_value
        Me.OnPropertyChanged "SerialPollByte"
        Me.ServiceRequested = cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
                a_value And cc_isr_Ieee488.ServiceRequestFlags.RequestingService
    End If
End Property

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

''' <summary>   Gets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status Byte command can be executed. </value>
Public Property Get ReadStatusByteExecutable() As Boolean
    ReadStatusByteExecutable = This.ReadStatusByteExecutable
End Property

''' <summary>   Sets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Read Status Byte command can be executed. </param>
Friend Property Let ReadStatusByteExecutable(ByVal a_value As Boolean)
    If This.ReadStatusByteExecutable <> a_value Then
        This.ReadStatusByteExecutable = a_value
        Me.OnPropertyChanged "ReadStatusByteExecutable"
    End If
End Property

''' <summary>   Handles the Serial Poll command event. </summary>
Public Sub SerialPollCommand()

    Const p_procedureName = "SerialPollCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper

    Dim p_details As String: p_details = VBA.vbNullString
    If Not Me.TryReadStatusByte(p_details, True) Then _
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

''' <summary>   Gets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get SerialPollExecutable() As Boolean
    SerialPollExecutable = This.SerialPollExecutable
End Property

''' <summary>   Sets the condition indicating that the Serial Poll command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Serial Poll command can be executed. </param>
Friend Property Let SerialPollExecutable(ByVal a_value As Boolean)
    If This.SerialPollExecutable <> a_value Then
        This.SerialPollExecutable = a_value
        Me.OnPropertyChanged "SerialPollExecutable"
    End If
End Property

''' <summary>   Tries to read (serial poll) or query (<c>*STB?</c>) the device status byte. </summary>
''' <remarks>   This method updates the <see cref="K2700ViewModel.SerialPollByte"/>. </remarks>
''' <param name="canQuery">    [Optional, Boolean, False] True if the <c>*STB?</c>
'''                            can be used to query the status byte in case
'''                            <see cref="TcpSession.GpibLanControllerAttached"/> is false. </param>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean] true if success. </returns>
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
        Me.SerialPollByte = p_statusByte
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Service Requested
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Friend Property Let ServiceRequested(ByVal a_value As Integer)
    If This.ServiceRequested <> a_value Then
        This.ServiceRequested = a_value
        Me.OnPropertyChanged "ServiceRequested"
    End If
End Property

''' <summary>   Handles the SRQ command event. </summary>
Public Sub ServiceRequestedQueryCommand()

    ' this does nothing
    ' Me.ServiceRequested = This.Session.ServiceRequested

    ' do this instead.
    Me.ReadStatusByte

End Sub

''' <summary>   Gets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Service Requested command can be executed. </value>
Public Property Get ServiceRequestedExecutable() As Boolean
    ServiceRequestedExecutable = This.ServiceRequestedExecutable
End Property

''' <summary>   Sets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Service Requested command can be executed. </param>
Friend Property Let ServiceRequestedExecutable(ByVal a_value As Boolean)
    If This.ServiceRequestedExecutable <> a_value Then
        This.ServiceRequestedExecutable = a_value
        Me.OnPropertyChanged "ServiceRequestedExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB interface management
' . . . . . . . . . . . . . . . . . . . . . . . . .

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB interface management: GO TO LOCAL
' . . . . . . . . . . . . . . . . . . . . . . . . .

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

''' <summary>   Gets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Go To Local command can be executed. </value>
Public Property Get GoToLocalExecutable() As Boolean
    GoToLocalExecutable = This.GoToLocalExecutable
End Property

''' <summary>   Sets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Go To Local command can be executed. </param>
Friend Property Let GoToLocalExecutable(ByVal a_value As Boolean)
    If This.GoToLocalExecutable <> a_value Then
        This.GoToLocalExecutable = a_value
        Me.OnPropertyChanged "GoToLocalExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB interface management: LOCAL LOCKOUT
' . . . . . . . . . . . . . . . . . . . . . . . . .

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

''' <summary>   Gets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Local Lockout command can be executed. </value>
Public Property Get LocalLockoutExecutable() As Boolean
    LocalLockoutExecutable = This.LocalLockoutExecutable
End Property

''' <summary>   Sets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Local Lockout command can be executed. </param>
Friend Property Let LocalLockoutExecutable(ByVal a_value As Boolean)
    If This.LocalLockoutExecutable <> a_value Then
        This.LocalLockoutExecutable = a_value
        Me.OnPropertyChanged "LocalLockoutExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB interface management: Power On Reset
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Tries to issue power on reset. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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
''' <param name="a_socketAddress">    [String] the socket dotted address and
'''                                   port number, e.g., <c>192.168.0.252:1234</c>. </value>
''' <param name="a_sessionTimeout">   [Integer] the timeout value in milliseconds. </value>
Public Sub PowerOnResetCommand(ByVal a_socketAddress As String, ByVal a_sessionTimeout As Long)

    Const p_procedureName = "PowerOnResetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    This.SocketAddress = a_socketAddress
    This.SessionTimeout = a_sessionTimeout
    
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: GPIB interface management: SDC
' . . . . . . . . . . . . . . . . . . . . . . . . .

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

''' <summary>   Gets the condition indicating that the Selective Device Clear command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Selective Device Clear command can be executed. </value>
Public Property Get SelectiveDeviceClearExecutable() As Boolean
    SelectiveDeviceClearExecutable = This.SelectiveDeviceClearExecutable
End Property

''' <summary>   Sets the condition indicating that the Selective Device Clear command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Selective Device Clear command can be executed. </param>
Friend Property Let SelectiveDeviceClearExecutable(ByVal a_value As Boolean)
    If This.SelectiveDeviceClearExecutable <> a_value Then
        This.SelectiveDeviceClearExecutable = a_value
        Me.OnPropertyChanged "SelectiveDeviceClearExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' GPIB-Lan: Connection Manangement
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Opens the connection and power on reset the GPIB Lan Controller. </summary>
''' <param name="a_socketAddress">    [String] the socket dotted address and
'''                                   port number, e.g., <c>192.168.0.252:1234</c>. </value>
''' <param name="a_sessionTimeout">   [Integer] the timeout value in milliseconds. </value>
''' <param name="a_delay">            [Double] the post seset delay. </param>
''' <param name="a_details">          [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryOpenConnectionPowerOnReset(ByVal a_socketAddress As String, ByVal a_sessionTimeout As Long, _
    ByVal a_delay As Double, ByRef a_details As String) As Boolean

    Const p_procedureName = "TryOpenConnectionPowerOnReset"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    ' initialize all elements.
    Me.Initialize
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_success Then _
        p_success = Me.TryPowerOnReset(p_details, a_delay)
    
    If p_success Then _
        p_success = Me.Connectable.TryOpenConnection(a_socketAddress, a_sessionTimeout, p_details)
    
    If Not p_success Then _
        p_details = "Failed opening the connection after power on reset; " & p_details

    If p_success Then
        This.SocketAddress = a_socketAddress
        This.SessionTimeout = a_sessionTimeout
    End If

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
''' <param name="a_socketAddress">    [String] the socket dotted address and
'''                                   port number, e.g., <c>192.168.0.252:1234</c>. </value>
''' <param name="a_sessionTimeout">   [Integer] the timeout value in milliseconds. </value>
Public Sub OpenConnectionPowerOnResetCommand(ByVal a_socketAddress As String, ByVal a_sessionTimeout As Long)

    Const p_procedureName = "OpenConnectionPowerOnResetCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopper
    
    Dim p_delay As Double: p_delay = 3
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean
    p_success = Me.TryOpenConnectionPowerOnReset(a_socketAddress, a_sessionTimeout, p_delay, p_details)
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

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Settings, Commands and Executables
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Tcp_Scpi.K2700"/>. </summary>
''' <value>   An <see cref="cc_isr_Tcp_Scpi.K2700"/>. </value>
Public Property Get K2700() As cc_isr_Tcp_Scpi.K2700
    Set K2700 = This.K2700
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Options
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Top Card name. </summary>
''' <value>   [String]. </value>
Public Property Get TopCard() As String
    TopCard = This.TopCard
End Property

''' <summary>   Sets the Top Card name. </summary>
''' <param name="a_value">   [String] the Top Card name. </value>
Friend Property Let TopCard(ByVal a_value As String)
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
Friend Property Let TopCardFunctionScanList(ByVal a_value As String)
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
Friend Property Let BottomCard(ByVal a_value As String)
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
Friend Property Let BottomCardFunctionScanList(ByVal a_value As String)
    If This.BottomCardFunctionScanList <> a_value Then
        This.BottomCardFunctionScanList = a_value
        Me.OnPropertyChanged "BottomCardFunctionScanList"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Device Under Test
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the number of DUTs that can be tested using either the front inputs or
''' the rear channels. </summary>
''' <param name="a_frontInputs">       [Boolean] True to take readings from the front inputs; otherwise,
'''                                    readings are taken from the rear inputs. </param>
''' <param name="a_maximumDutCount">   [Integer] the maximum number of DUTs that can be tested. </param>
Public Function GetDutCount(ByVal a_frontInputs As Boolean, ByVal a_maximumDutCount As Integer) As Integer
    If a_frontInputs Then
        GetDutCount = a_maximumDutCount
    Else
        GetDutCount = IIf(This.K2700.RouteSystem.ChannelCount <= a_maximumDutCount, _
            This.K2700.RouteSystem.ChannelCount, a_maximumDutCount)
    End If
End Function

''' <summary>   Gets the number of DUTs to be tested. The selected
''' DUT number must be between 1 and this value. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutCount() As Integer
    DutCount = This.DutCount
End Property

''' <summary>   Gets the number of DUTs to be tested. The selected
''' DUT number must be between 1 and this value. </summary>
''' <param name="a_value">   [Integer] the number of DUTs to be tested. </param>
Friend Property Let DutCount(ByVal a_value As Integer)
    If This.DutCount <> a_value Then
        This.DutCount = a_value
        Me.OnPropertyChanged "DutCount"
    End If
End Property

''' <summary>   Gets the maximum number of DUTs. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutCount() As Integer
    MaximumDutCount = This.MaximumDutCount
End Property

''' <summary>   Gets the maximum number of DUTs. </summary>
''' <param name="a_value">   [Integer] the maximum number of DUTs. </param>
Friend Property Let MaximumDutCount(ByVal a_value As Integer)
    If This.MaximumDutCount <> a_value Then
        This.MaximumDutCount = a_value
        Me.OnPropertyChanged "MaximumDutCount"
    End If
End Property

''' <summary>   Gets the selected DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedDutNumber() As Integer
    SelectedDutNumber = This.SelectedDutNumber
End Property

''' <summary>   Sets the selected DUT number. </summary>
''' <param name="a_value">   [Integer] the selected DUT number value. </value>
Friend Property Let SelectedDutNumber(ByVal a_value As Integer)
    If This.SelectedDutNumber <> a_value Then
        This.SelectedDutNumber = a_value
        Me.OnPropertyChanged "SelectedDutNumber"
    End If
End Property

''' <summary>   Gets the target DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get TargetDutNumber() As Integer
    TargetDutNumber = This.TargetDutNumber
End Property

''' <summary>   Sets the target DUT number. </summary>
''' <param name="a_value">   [Integer] the target DUT number value. </value>
Friend Property Let TargetDutNumber(ByVal a_value As Integer)
    If This.TargetDutNumber <> a_value Then
        This.TargetDutNumber = a_value
        Me.OnPropertyChanged "TargetDutNumber"
    End If
End Property

''' <summary>   Gets a DUT number from the DUT number caption. </summary>
''' <param name="a_value">                [String] the selected DUT number caption, e.g., R2. </value>
''' <param name="a_prefix">               [String] the DUT number caption prefix, e.g., R. </value>
''' <param name="a_title">                [String] the DUT number title, e.g., resistance. </value>
''' <param name="a_maximumDutNumber">     [Integer] the maxiimum DUT number. </value>
''' <param name="a_details">              [Out, String] Details the failure information if any. </param>
''' <returns>   [Integer] Returns 0 if invalid DUT number caption was selected. </returns>
Public Function TryParseDutNumberCaption(ByVal a_value As String, ByVal a_prefix As String, _
    ByVal a_title As String, ByVal a_maximumDutNumber As Integer, ByRef a_details As String) As Integer

    Const p_procedureName = "TryParseSelectedDutNumber"
    
    Dim p_dutNumber As Integer: p_dutNumber = 0
    
    If 0 >= VBA.InStr(1, VBA.Left$(a_value, VBA.Len(a_prefix)), _
        a_prefix, VBA.VbCompareMethod.vbTextCompare) Then
     
        a_details = "The selected value '" & a_value & _
               "' is not a valid " & a_title & _
               " number. Please select a valid " & a_title & _
               " number, e.g., " & a_prefix & "2, and try again."
  
    Else
    
        p_dutNumber = VBA.CInt(Mid$(a_value, VBA.Len(a_prefix) + 1))
        
        If a_maximumDutNumber < 1 Then
        
            TryParseDutNumberCaption = 0
            a_details = "Failed parsing DUT number caption; The maxumum DUT number was not set."
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " " & a_details
        
        ElseIf p_dutNumber <= 0 Or p_dutNumber > a_maximumDutNumber Then
        
            a_details = "The selected value '" & a_value & _
                   "' yields an out of range " & a_title & _
                   " number. Valid " & a_title & " numbers are between 1 and " & _
                   VBA.CStr(a_maximumDutNumber) & _
                   ". Please select a valid " & a_title & _
                   " number, e.g., " & a_prefix & "2, and try again."
            
            p_dutNumber = 0
            
        End If
    
    End If
  
    TryParseDutNumberCaption = p_dutNumber
  
End Function

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measured Device Under Test
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the measured DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredDutNumber() As Integer
    MeasuredDutNumber = This.MeasuredDutNumber
End Property

''' <summary>   Sets the measured DUT number. </summary>
''' <param name="a_value">   [Integer] the measured DUT number value. </value>
Friend Property Let MeasuredDutNumber(ByVal a_value As Integer)
    If This.MeasuredDutNumber <> a_value Then
        This.MeasuredDutNumber = a_value
        Me.OnPropertyChanged "MeasuredDutNumber"
    End If
End Property

''' <summary>   Gets the auto increment DUT number Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementDutNumberEnabled() As Boolean
    AutoIncrementDutNumberEnabled = This.AutoIncrementDutNumberEnabled
End Property

''' <summary>   Sets the auto increment DUT number Enabled. </summary>
''' <param name="a_value">   [Boolean] the auto increment DUT number Enabled value. </value>
Friend Property Let AutoIncrementDutNumberEnabled(ByVal a_value As Boolean)
    If This.AutoIncrementDutNumberEnabled <> a_value Then
        This.AutoIncrementDutNumberEnabled = a_value
        Me.OnPropertyChanged "AutoIncrementDutNumberEnabled"
    End If
End Property

''' <summary>   Gets the Single Read Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get SingleReadEnabled() As Boolean
    SingleReadEnabled = This.SingleReadEnabled
End Property

''' <summary>   Sets the Single Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Single Read Enabled value. </value>
Friend Property Let SingleReadEnabled(ByVal a_value As Boolean)
    If This.SingleReadEnabled <> a_value Then
        This.SingleReadEnabled = a_value
        Me.OnPropertyChanged "SingleReadEnabled"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Input terminals
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsRequired() As Boolean
    FrontInputsRequired = This.FrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are Required. </value>
Friend Property Let FrontInputsRequired(ByVal a_value As Boolean)
    If This.FrontInputsRequired <> a_value Then
        This.FrontInputsRequired = a_value
        Me.OnPropertyChanged "FrontInputsRequired"
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
Friend Property Let FrontInputsHasValue(ByVal a_value As Boolean)
    If This.FrontInputsHasValue <> a_value Then
        This.FrontInputsHasValue = a_value
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
Friend Property Let FrontInputsValue(ByVal a_value As Boolean)
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
Friend Property Let FrontInputsCaption(ByVal a_value As String)
    If This.FrontInputsCaption <> a_value Then
        This.FrontInputsCaption = a_value
        Me.OnPropertyChanged "FrontInputsCaption"
    End If
End Property

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

''' <summary>   Gets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Inputs command can be executed. </value>
Public Property Get QueryInputsExecutable() As Boolean
    QueryInputsExecutable = This.QueryInputsExecutable
End Property

''' <summary>   Sets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Query Inputs command can be executed. </param>
Friend Property Let QueryInputsExecutable(ByVal a_value As Boolean)
    If This.QueryInputsExecutable <> a_value Then
        This.QueryInputsExecutable = a_value
        Me.OnPropertyChanged "QueryInputsExecutable"
    End If
End Property

''' <summary>   Returns true if the inputs are valid for front or rear settings. </summary>
''' <param name="a_frontInputs">    [Boolean] True if readings are made from the front inputs. </param>
''' <param name"a_userMessage">      [Out, String] User message. </param>
''' <param name"a_promptUser">       [Boolean] true to prompt user. </param>
''' <returns>   [Boolean]. </returns>
Public Function ValidateInputs(ByVal a_frontInputs As Boolean, _
    ByRef a_userMessage As String, ByVal a_promptUser As Boolean) As Boolean

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
    
    Me.FrontInputsRequired = a_frontInputs
    
    ' check if front or rear
    If a_frontInputs And Not This.FrontInputsValue Then

        ' set the proper message
        p_userMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."

        p_success = False
    
    ElseIf Not a_frontInputs And This.FrontInputsValue Then

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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode
' . . . . . . . . . . . . . . . . . . . . . . . . .

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
Friend Property Let Measuring(ByVal a_value As Boolean)
    If This.Measuring <> a_value Then
        This.Measuring = a_value
        Me.OnPropertyChanged "Measuring"
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
Friend Property Let MeasurementMode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_value Then
        This.MeasurementMode = a_value
        Me.OnPropertyChanged "MeasurementMode"
    End If
End Property

''' <summary>   Sets the measurement mode for unit testing. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
Public Sub MeasurementModeUnitTestSetter(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_value Then
        This.MeasurementMode = a_value
        Me.OnPropertyChanged "MeasurementMode"
    End If
End Sub

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Measurement Sate
' . . . . . . . . . . . . . . . . . . . . . . . . .

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
    ByVal a_info As String, ByVal a_measuring As Boolean, ByVal a_singleRead As Boolean)
    
    ' turn off the flag handling a timer event.
    This.HandlingTimerEvent = False
    
    ' set the single read mode.
    Me.SingleReadEnabled = a_singleRead
    
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

    ' enable selecting a new function when not monitoring
    Me.FunctionNameExecutable = Me.Connected And (a_mode <> cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring)
    
    ' disable the stop monitoring action
    Me.StopMonitoringExecutable = Me.Connected And (cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring = a_mode)
    
    ' enable the start monitoring action in external mode
    Me.StartMonitoringExecutable = Me.Connected And (cc_isr_Tcp_Scpi.MeasurementModeOption.External = a_mode)
    
    ' enable immediate readings in immediate mode
    Me.MeasureExecutable = Me.Connected And (cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate = a_mode)
    
    ' enable toggling the configurations
    Me.ImmediateTriggerOptionExecutable = Me.Connected And (a_mode <> cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring)
    Me.ExternalTriggerOptionExecutable = Me.Connected And (a_mode <> cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring)
    
    ' finally, indicates that external trigger monitoring is enabled thus updating the related commands.
    Me.ExternalTrigMonitoringEnabled = a_mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
    
    RaiseEvent MeasurementStateChanged( _
            cc_isr_Tcp_Scpi.Factory.NewStateChangedEventArgs.Initialize())
    ' allow the event to take place.
    DoEvents

End Sub

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Sense function
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the sense function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Internal. Sets the sense function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Friend Property Let SenseFunctionName(ByVal a_value As String)
    If This.SenseFunctionName <> a_value Then
        This.SenseFunctionName = a_value
        Me.OnPropertyChanged "SenseFunctionName"
        Me.OnSenseFunctionChanged
    End If
End Property

''' <summary>   Handles the sense function change; Set the sense faction caption and units. </summary>
Public Sub OnSenseFunctionChanged()

    Select Case This.SenseFunctionName

        Case "CONT"
    
            Me.SenseFunctionCaption = "Continuity"
            Me.SenseFunctionUnicode = VBA.Asc("1")
            ' Me.SenseFunctionUnit = "T/F"
    
        Case "CURR:AC"
    
            Me.SenseFunctionCaption = "Current AC"
            Me.SenseFunctionUnicode = VBA.Asc("A")
            Me.SenseFunctionUnit = "A"
    
        Case "CURR:DC"
    
            Me.SenseFunctionCaption = "Current DC"
            Me.SenseFunctionUnicode = VBA.Asc("A")
            Me.SenseFunctionUnit = "A"
    
        Case "FREQ"
    
            Me.SenseFunctionCaption = "Frequency"
            Me.SenseFunctionUnicode = 13200
            Me.SenseFunctionUnit = "Hz"
    
        Case "FRES"
    
            Me.SenseFunctionCaption = "4-Wire Res."
            Me.SenseFunctionUnicode = 937
            Me.SenseFunctionUnit = VBA.Chr$(234)
    
        Case "PER"
    
            Me.SenseFunctionCaption = "Period"
            Me.SenseFunctionUnicode = VBA.Asc("s")
            Me.SenseFunctionUnit = "s"
    
        Case "RES"
    
            Me.SenseFunctionCaption = "2-Wire Res."
            Me.SenseFunctionUnicode = 937
            Me.SenseFunctionUnit = VBA.Chr$(234)
    
        Case "TEMP"
    
            Me.SenseFunctionCaption = "Temperature"
            Me.SenseFunctionUnicode = VBA.Asc("K")
            Me.SenseFunctionUnit = "K"
    
        Case "VOLT:AC"
    
            Me.SenseFunctionCaption = "Voltage AC"
            Me.SenseFunctionUnicode = VBA.Asc("v")
            Me.SenseFunctionUnit = "v"
    
        Case "VOLT:DC"
    
            Me.SenseFunctionCaption = "Voltage DC"
            Me.SenseFunctionUnicode = VBA.Asc("v")
            Me.SenseFunctionUnit = "v"
    
        Case Else
    
            Me.SenseFunctionCaption = "unknown"
            Me.SenseFunctionUnicode = VBA.Asc("?")
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
Friend Property Let SenseFunctionCaption(ByVal a_value As String)
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

''' <summary>   Gets the Sense Function Unit. </summary>
''' <value>   [Integer]. </value>
Public Property Get SenseFunctionUnicode() As Integer
    SenseFunctionUnicode = This.SenseFunctionUnicode
End Property

''' <summary>   Sets the Sense Function Unit unicoide. </summary>
''' <param name="a_value">   [Integer] the Sense Function Unit unicode value. </value>
Friend Property Let SenseFunctionUnicode(ByVal a_value As Integer)
    If This.SenseFunctionUnicode <> a_value Then
        This.SenseFunctionUnicode = a_value
        Me.OnPropertyChanged "SenseFunctionUnicode"
    End If
End Property

''' <summary>   Tries to validate the expected function by name. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Sense functions
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Sense Functions. </summary>
''' <value>   [Collection]. </value>
Public Property Get SenseFunctions() As Collection
    Set SenseFunctions = This.SenseFunctions
End Property

''' <summary>   Sets the Sense Functions. </summary>
''' <param name="a_value">   [Collection] the Sense Functions value. </value>
Friend Property Let SenseFunctions(ByVal a_value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.SenseFunctions, a_value) Then
        Set This.SenseFunctions = a_value
        Me.OnPropertyChanged "SenseFunctions"
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

''' <summary>   Gets the condition indicating that the function name command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the function name command can be executed. </value>
Public Property Get FunctionNameExecutable() As Boolean
    FunctionNameExecutable = This.FunctionNameExecutable
End Property

''' <summary>   Sets the function name condition. </summary>
''' <param name="a_value">   [Boolean] True if the function name command can be executed. </param>
Friend Property Let FunctionNameExecutable(ByVal a_value As Boolean)
    If This.FunctionNameExecutable <> a_value Then
        This.FunctionNameExecutable = a_value
        Me.OnPropertyChanged "FunctionNameExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Reading
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Measured Reading. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = This.MeasuredReading
End Property

''' <summary>   Sets the Measured Reading. </summary>
''' <param name="a_value">   [String] the Measured Reading value. </value>
Friend Property Let MeasuredReading(ByVal a_value As String)
    If This.MeasuredReading <> a_value Then
        This.MeasuredReading = a_value
        Me.OnPropertyChanged "MeasuredReading"
        If 0 < VBA.Len(a_value) Then
            Dim p_value As Double
            Dim p_details As String
            If cc_isr_Core.StringExtensions.TryParseDouble(a_value, p_value, p_details) Then
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
Friend Property Let MeasuredValue(ByVal a_value As Double)
    If This.MeasuredValue <> a_value Then
        This.MeasuredValue = a_value
        Me.OnPropertyChanged "MeasuredValue"
    End If
End Property

''' <summary>   Gets the Reading Offset. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadingOffset() As Double
    ReadingOffset = This.ReadingOffset
End Property

''' <summary>   Sets the Reading Offset. </summary>
''' <param name="a_value">   [Double] the Reading Offset value. </value>
Friend Property Let ReadingOffset(ByVal a_value As Double)
    If This.ReadingOffset <> a_value Then
        This.ReadingOffset = a_value
        Me.OnPropertyChanged "ReadingOffset"
    End If
End Property

''' <summary>   Clears the readings. </summary>
''' <param name="a_value">            [Double] the value to use for clearing the readings.
'''                                   This value might represent a probe resistance or
'''                                   DC offset voltage. </value>
''' <param name="a_firstDutNumber">   [Integer] the first DUT number. </value>
''' <param name="a_lastDutNumber">    [Integer] the last DUT number. </value>
Public Sub ClearReadings(ByVal a_value As Double, _
    ByVal a_firstDutNumber As Integer, ByVal a_lastDutNumber As Integer)

    Const p_procedureName = "ClearReadings"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReadingOffset = a_value
    Dim p_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs
    Set p_eventArgs = cc_isr_Tcp_Scpi.Factory.NewDutReadingEventArgs.Initialize(0, VBA.vbNullString)
    Dim p_dutNumber As Integer
    For p_dutNumber = a_firstDutNumber To a_lastDutNumber
        VBA.DoEvents
        OnDutMeasured p_eventArgs.Initialize(p_dutNumber, VBA.CStr(a_value))
    Next p_dutNumber
  
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
''' <param name="a_value">            [Double] the value to use for clearing the readings.
'''                                   This value might represent a probe resistance or
'''                                   DC offset voltage. </value>
''' <param name="a_firstDutNumber">   [Integer] the first DUT number. </value>
''' <param name="a_lastDutNumber">    [Integer] the last DUT number. </value>
Public Sub ClearReadingsCommand(ByVal a_value As Double, _
    ByVal a_firstDutNumber As Integer, ByVal a_lastDutNumber As Integer)

    Const p_procedureName = "ClearReadingsCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.ClearReadings a_value, a_firstDutNumber, a_lastDutNumber
  
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

''' <summary>   Gets the condition indicating that the Clear Readings command
'''              can be executed. </summary>
''' <value>   [Boolean] True if the Clear Readings command can be executed. </value>
Public Property Get ClearReadingsExecutable() As Boolean
    ClearReadingsExecutable = This.ClearReadingsExecutable
End Property

''' <summary>   Sets the condition indicating that the Clear Readings command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Clear Readings command can be executed. </param>
Friend Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If This.ClearReadingsExecutable <> a_value Then
        This.ClearReadingsExecutable = a_value
        Me.OnPropertyChanged "ClearReadingsExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode Configuration
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Internal. Tries to configure a measurement mode. </summary>
''' <remarks>   This method requires that instrument be connected as per
''' <see cref=TryRestoreInitialState"/>. </remarks>
''' <param name="a_mode">      [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] defines the
'''                            measurment settings. </value>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Friend Function TryConfigureMeasure(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureMeasure"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' notify of the changing measurement state
    Me.OnMeasurementStateChanging

    ' it looks like we need to validate because the instrument inputs can be changed at any time/
    Dim p_inputsValidationRequired As Boolean
    p_inputsValidationRequired = cc_isr_Tcp_Scpi.MeasurementModeOption.External = a_mode.Mode Or _
        cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate = a_mode.Mode

    If p_success And p_inputsValidationRequired Then
    
        ' unfortunately, reading the inputs after the reset belore returns an incorrect value
        ' so this was moved as a workaround
        
        If p_inputsValidationRequired Then
            
            Me.FrontInputsRequired = a_mode.FrontInputs
            
            p_success = Me.ValidateInputs(a_mode.FrontInputs, p_details, True)
        
            Me.LastMessage = p_details
            
        End If
        
    End If

    If p_success Then
    
        ' reset if changing modes.
        If a_mode.Mode <> This.MeasurementMode Then
            
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

    Dim p_modeCaption As String: p_modeCaption = VBA.vbNullString
    If p_success Then
    
        ' set the measurement conditions
        Me.DutCount = a_mode.DutCount
        Me.MaximumDutCount = a_mode.MaximumDutCount
        Me.ReadingOffset = a_mode.ReadingOffset
        Me.SenseFunctionName = a_mode.SenseFunction
        Me.SingleReadEnabled = a_mode.SingleRead
        Me.AutoIncrementDutNumberEnabled = a_mode.AutoIncrement
        Me.ReadingBeepEnabled = a_mode.BeepEnabled
        Me.TimerInterval = a_mode.TimerInterval
        
        If a_mode.SingleRead Then
            Me.SelectedDutNumber = a_mode.DutNumber
        Else
            Me.TargetDutNumber = a_mode.DutNumber
        End If
        
        Select Case a_mode.Mode
            
            Case cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous
            
                p_modeCaption = "continuous trigger mode"
                ' configure measuring a single value on each read.
                p_success = This.K2700.TryConfigureContinuousFunctionAutoRange(This.SenseFunctionName, p_details)
            
            Case cc_isr_Tcp_Scpi.MeasurementModeOption.External
            
                Me.TargetDutNumber = a_mode.DutNumber
                p_modeCaption = "external trigger mode"
                ' Configure external triggering reading
                p_success = This.K2700.TryConfigureExternalTriggerReading(This.SenseFunctionName, p_details)
                    
            Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
            
                p_modeCaption = "immediate trigger mode"
            
                ' configure measuring a single value on each read.
                p_success = This.K2700.TryConfigureSingleFunctionReading(This.SenseFunctionName, p_details)
            
        End Select
            
        If Not p_success Then _
            p_details = "Failed configuring the " & p_modeCaption & "; " & p_details
    
    End If
    
    If p_success Then
    
        ' validate the sense function.
        p_success = Me.TryValidateSenseFunction(This.SenseFunctionName, p_details)
    
        If Not p_success Then _
            p_details = "Failed setting the 2700 instrument sense function; " & p_details
    
    End If

    ' enable the mode executables and conditions.
    If p_success Then
        Me.OnMeasurementStateChanged a_mode.Mode, p_modeCaption, True, Me.SingleReadEnabled
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False, Me.SingleReadEnabled
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring measurement; " & p_details

    TryConfigureMeasure = p_success
    
End Function

''' <summary>   Configure measure mode. </summary>
''' <param name="a_mode">      [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] defines the
'''                            measurment settings. </value>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function ConfigureMeasureCommand(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "ConfigureMeasureCommand"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_success As Boolean: p_success = False
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' configure for continuous trigger

    p_success = Me.TryConfigureMeasure(a_mode, p_details)
    
    ' throw an error if failed to configure
    
    If Not p_success Then
        a_details = "Failed configuring measure; " & p_details
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " " & a_details
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ConfigureMeasureCommand = p_success
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Initial
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Internal. Tries to configure the intial trigger mode to continuous trigger mode. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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
    
        ' default to front inputs sense function
        Me.SenseFunctionName = This.FrontInputsSenseFunctionName

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
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous, _
            "continuous", False, Me.SingleReadEnabled
    Else
        Me.OnMeasurementStateChanged Me.MeasurementMode, "unchanged", False, Me.SingleReadEnabled
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Function Names
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Sense Function Name for rear inputs. </summary>
''' <value>   [String]. </value>
Public Property Get RearInputsSenseFunctionName() As String
    RearInputsSenseFunctionName = This.RearInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for rear inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Friend Property Let RearInputsSenseFunctionName(ByVal a_value As String)
    If This.RearInputsSenseFunctionName <> a_value Then
        This.RearInputsSenseFunctionName = a_value
        Me.OnPropertyChanged "RearInputsSenseFunctionName"
    End If
End Property

''' <summary>   Gets the Sense Function Name for front inputs. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsSenseFunctionName() As String
    FrontInputsSenseFunctionName = This.FrontInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for front inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Friend Property Let FrontInputsSenseFunctionName(ByVal a_value As String)
    If This.FrontInputsSenseFunctionName <> a_value Then
        This.FrontInputsSenseFunctionName = a_value
        Me.OnPropertyChanged "FrontInputsSenseFunctionName"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Immediate
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required for Immediate trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ImmediateFrontInputsRequired() As Boolean
    ImmediateFrontInputsRequired = This.ImmediateFrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required for Immediate trigger  mode. </summary>
''' <param name="a_value">   [Boolean] the Front Inputs Required. </value>
Friend Property Let ImmediateFrontInputsRequired(ByVal a_value As Boolean)
    If This.ImmediateFrontInputsRequired <> a_value Then
        This.ImmediateFrontInputsRequired = a_value
        Me.OnPropertyChanged "ImmediateFrontInputsRequired"
    End If
End Property

''' <summary>   Tries to measure immediately. </summary>
''' <remarks>   This method assumes that the inputs were validated before configuring the
''' measurement state. </remarks>
''' <param name="a_details">         [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryMeasureImmediately(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryMeasureImmediately"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    Dim p_reading As String
    Dim p_dutNumber As Integer

    If This.AutoIncrementDutNumberEnabled Then
    
        ' Clear all cells
        Me.ClearReadings Me.ReadingOffset, 1, Me.MaximumDutCount

        ' scan all channels
        For p_dutNumber = 1 To Me.DutCount
    
            ' allow events to occur
            VBA.DoEvents
        
            ' set the target DUT number
            Me.TargetDutNumber = p_dutNumber
        
            Me.RestartStopper
            If This.FrontInputsValue Then
                p_success = This.K2700.TryTakeFrontReading(p_reading, p_details, p_dutNumber)
                If Not p_success Then _
                    p_details = "Failed taking a front reading; " & p_details
            Else
                p_success = This.K2700.TryTakeRearReading(p_dutNumber, p_reading, p_details)
                If Not p_success Then _
                    p_details = "Failed taking a rear reading; " & p_details
            End If
            Me.ReadStopper
    
            If Not p_success Then Exit For
            
        Next p_dutNumber

    Else

        p_dutNumber = This.SelectedDutNumber
    
        Me.RestartStopper
        
        If This.FrontInputsValue Then
            p_success = This.K2700.TryTakeFrontReading(p_reading, p_details, p_dutNumber)
            If Not p_success Then _
                p_details = "Failed taking a front reading; " & p_details
        Else
            p_success = This.K2700.TryTakeRearReading(p_dutNumber, p_reading, p_details)
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
''' <param name="a_details">         [Out, String] Details the failure information if any. </param>
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
    Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate, _
        "Done", False, Me.SingleReadEnabled

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
''' <param name="a_frontInputs">     [Boolean] True if readings are made from the front inputs. </param>
''' <param name="a_details">         [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function ValidateInputsAndMeasureCommand(ByVal a_frontInputs As Boolean, a_details As String) As Boolean

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
    p_success = Me.ValidateInputs(a_frontInputs, p_details, True)
    If Not p_success Then
        ' this cludge was added.
        cc_isr_Core_IO.Factory.NewStopwatch.Wait 50 ' 25 this was not enough
        p_success = Me.ValidateInputs(a_frontInputs, p_details, True)
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

''' <summary>   Gets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Measure command can be executed. </value>
Public Property Get MeasureExecutable() As Boolean
    MeasureExecutable = This.MeasureExecutable
End Property

''' <summary>   Sets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Measure command can be executed. </param>
Friend Property Let MeasureExecutable(ByVal a_value As Boolean)
    If This.MeasureExecutable <> a_value Then
        This.MeasureExecutable = a_value
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
Friend Property Let ImmediateTriggerOptionExecutable(ByVal a_value As Boolean)
    If This.ImmediateTriggerOptionExecutable <> a_value Then
        This.ImmediateTriggerOptionExecutable = a_value
        Me.OnPropertyChanged "ImmediateTriggerOptionExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: External Trigger
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required for External trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExternalFrontInputsRequired() As Boolean
    ExternalFrontInputsRequired = This.ExternalFrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required for External trigger  mode. </summary>
''' <param name="a_value">   [Boolean] the Front Inputs Required. </value>
Friend Property Let ExternalFrontInputsRequired(ByVal a_value As Boolean)
    If This.ExternalFrontInputsRequired <> a_value Then
        This.ExternalFrontInputsRequired = a_value
        Me.OnPropertyChanged "ExternalFrontInputsRequired"
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
Friend Property Let ExternalTriggerOptionExecutable(ByVal a_value As Boolean)
    If This.ExternalTriggerOptionExecutable <> a_value Then
        This.ExternalTriggerOptionExecutable = a_value
        Me.OnPropertyChanged "ExternalTriggerOptionExecutable"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: External Trigger Monitring
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Monitoring State. </summary>
''' <value>   [String]. </value>
Public Property Get MonitoringState() As String
    MonitoringState = This.MonitoringState
End Property

''' <summary>   Internal. Starts monitoring the instrument for external triggered measurements. </summary>
Public Sub StartMonitoringExternalTriggers()

    Const p_procedureName = "StartMonitoringExternalTriggers"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.AutoIncrementDutNumberEnabled Then
    
        ' Clear all readings
        Me.ClearReadings Me.ReadingOffset, 1, Me.MaximumDutCount

    End If

    Me.MonitoringState = "External Trigger On"

    If This.SingleReadEnabled Then
        ' set the single reading DUT number for manual scan
        ' to the DUT number entered by the operator.
        Me.TargetDutNumber = This.SelectedDutNumber
    Else
        ' set the first DUT number for manual scan
        Me.TargetDutNumber = 1
    End If

#If False Then
    If This.AutoIncrementDutNumberEnabled Then
        ' set the first DUT number for manual scan
        Me.TargetDutNumber = 1
    Else
        ' set the single reading DUT number for manual scan
        ' to the DUT number entered by the operator.
        Me.TargetDutNumber = This.SelectedDutNumber
    End If
#End If
   
    ' initiate awaiting for the external trigger
    ' this is handled by the K2700.MonitorExternalTrigger method.
    ' This.Session.WriteLine ":INIT;*CLS"

    If Me.TimerControlled Then
    
       ' start the timer
       Me.StartTimer Me.TimerInterval
    
       ' take some time for it to take effect.
       cc_isr_Core_IO.Factory.NewStopwatch.Wait Me.TimerInterval
    
    Else
        
        ' if not using timer, go to monitoring mode
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, _
            "Trigger monitoring commenced", True, Me.SingleReadEnabled
        
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

''' <summary>   Sets the Monitoring State. </summary>
''' <param name="a_value">   [String] the Monitoring State value. </value>
Friend Property Let MonitoringState(ByVal a_value As String)
    If This.MonitoringState <> a_value Then
        This.MonitoringState = a_value
        Me.OnPropertyChanged "MonitoringState"
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
Friend Property Let StartMonitoringExecutable(ByVal a_value As Boolean)
    If This.StartMonitoringExecutable <> a_value Then
        This.StartMonitoringExecutable = a_value
        Me.OnPropertyChanged "StartMonitoringExecutable"
        Me.OnPropertyChanged "ToggleMonitoringExecutable"
    End If
End Property

''' <summary>   Internal. Stops monitoring the instrument for external triggered measurements. </summary>
Friend Sub StopMonitoringExternalTriggers()

    Const p_procedureName = "StopMonitoringExternalTriggers"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' request a stop of external trigger monitoring.
    Me.StopRequested = True
    
    If Me.TimerControlled Then
    
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
            Me.TimerStarted = False
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
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, _
            "Trigger monitoring done", False, Me.SingleReadEnabled

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

''' <summary>   Gets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </value>
Public Property Get StopMonitoringExecutable() As Boolean
    StopMonitoringExecutable = This.StopMonitoringExecutable
End Property

''' <summary>   Sets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <param name="a_value">   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </param>
Friend Property Let StopMonitoringExecutable(ByVal a_value As Boolean)
    If This.StopMonitoringExecutable <> a_value Then
        This.StopMonitoringExecutable = a_value
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

''' <summary>   Gets the External Trigger Monitoring Enabled condition. </summary>
''' <remarks>   Indicates that external trigger monitoring is on. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = This.ExternalTrigMonitoringEnabled
End Property

''' <summary>   Sets the External Trigger Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trigger Monitoring Enabled value. </value>
Friend Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If This.ExternalTrigMonitoringEnabled <> a_value Then
        This.ExternalTrigMonitoringEnabled = a_value
        Me.OnPropertyChanged "ExternalTrigMonitoringEnabled"
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: External Trigger Monitring Timer Control
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Pause Requested condition. </summary>
''' <remarks>   When true, the timer is pauses on the next timer event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get PauseRequested() As Boolean
    PauseRequested = This.PauseRequested
End Property

''' <summary>   Sets the Pause Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Pause Requested value. </value>
Friend Property Let PauseRequested(ByVal a_value As Boolean)
    If This.PauseRequested <> a_value Then
        This.PauseRequested = a_value
        Me.OnPropertyChanged "PauseRequested"
    End If
End Property

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
Friend Property Let StopRequested(ByVal a_value As Boolean)
    If This.StopRequested <> a_value Then
        This.StopRequested = a_value
        Me.OnPropertyChanged "StopRequested"
    End If
End Property

''' <summary>   Gets the Timer Controlled condition. </summary>
''' <value>   [Boolean] True if external trigger monitoring is timer controlled; otherwise,
''' external trigger monitoring is polled, which is solely for testing purposes. </value>
Public Property Get TimerControlled() As Boolean
    TimerControlled = This.TimerInterval > 0
End Property

''' <summary>   Gets the Timer Started condition. </summary>
''' <value>   [Boolean] True if the timer is active. </value>
Public Property Get TimerStarted() As Boolean
    TimerStarted = This.TimerStarted
End Property

''' <summary>   Sets the Timer Started condition. </summary>
''' <param name="a_value">   [Boolean] the value. </param>
Friend Property Let TimerStarted(ByVal a_value As Boolean)
    This.TimerStarted = a_value
End Property

''' <summary>   Gets the Timer Interval in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Timer Interval in milliseconds value. </value>
Friend Property Let TimerInterval(ByVal a_value As Integer)
    If This.TimerInterval <> a_value Then
        This.TimerInterval = a_value
        Me.OnPropertyChanged "TimerInterval"
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

' . . . . . . . . . . . . . . . . . . . . . . . . .
' Data Acquisition: Auto scan
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Starts an immediate triggering and take multiple readings. </summary>
''' <remarks>   In immeidate scan we have the following settings:
''' auto increment DUT number: True;
''' Front Inputs Required: False;
''' Single Readings: False;
''' Triggering Mode: Immediate.
''' </remarks>
''' <param name="a_mode">      [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean] True if the command succeeded. </returns>
Public Function StartMeasureCommand(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    ByRef a_details As String) As Boolean

    Const p_procedureName As String = "StartMeasureCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString

    ' next set immediate triggering mode.
    p_success = Me.ConfigureMeasureCommand(a_mode, p_details)
    
    If p_success Then
    
        ' next takes the measurements
    
        If Me.MeasureExecutable Then
            ' make an immediate measure of one or all channels.
            p_success = Me.MeasureImmediatelyCommand(p_details)
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Measure command not enabled."
        End If
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details
    StartMeasureCommand = p_success

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not Me Is Nothing Then _
        Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    p_details = Me.LastErrorMessage
    p_success = False
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Starts external trigger monitoring in multiple or single readings mode. </summary>
''' <remarks>   In manual scan we have the following settings:
''' auto increment DUT number: True;
''' Front Inputs Required: True;
''' Single Readings: False;
''' Triggering Mode: External.
''' </remarks>
''' <param name="a_mode">      [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>] the measure configuration. </param>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean] True if the command succeeded. </returns>
Public Function StartMonitoringCommand(ByVal a_mode As cc_isr_Tcp_Scpi.MeasureMode, _
    a_details As String) As Boolean

    Const p_procedureName As String = "StartMonitoringCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_success As Boolean: p_success = True
    Dim p_details As String: p_details = VBA.vbNullString

    ' next set external triggering triggering mode.
    p_success = Me.ConfigureMeasureCommand(a_mode, p_details)
   
    If p_success Then
    
        ' next start external trigger monitoring
        ' this should wait for monitoring to commence or raise an error.
        If Me.StartMonitoringExecutable Then
            Me.StartMonitoringExternalTriggersCommand
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Start monitoring command not enabled."
        End If
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details
    StartMonitoringCommand = p_success

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not Me Is Nothing Then _
        Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    p_details = Me.LastErrorMessage
    p_success = False
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    Me.StopMonitoringExternalTriggersCommand

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="DutMeasured"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref=""/>] The DUT reading event arguments. </param>
Private Sub OnDutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

    Const p_procedureName = "OnDutMeasured"

    Dim p_reading As String: p_reading = a_eventArgs.Reading
    
    ' set the measured value.
    Me.MeasuredReading = a_eventArgs.Reading
    
    Dim p_value As Double
    If VBA.Len(p_reading) = 0 Then
        p_value = cc_isr_Ieee488.Syntax.NotANumber
        Me.MeasuredValue = cc_isr_Ieee488.Syntax.NotANumber
    Else
        Dim p_details As String
        If cc_isr_Core.StringExtensions.TryParseDouble(p_reading, p_value, p_details) Then
            Me.MeasuredValue = p_value
        Else
            Me.MeasuredValue = cc_isr_Ieee488.Syntax.NotANumber
            Me.LastErrorMessage = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                " Failed parsing reading; " & p_details
        End If
    End If

    ' raise the measurement event.
    RaiseEvent DutMeasured(a_eventArgs)
                                                                      
    ' allow the event to take place.
    DoEvents
    
    ' update the event arguments in case changes were made.
    a_eventArgs.Initialize a_eventArgs.DutNumber, a_eventArgs.Reading

    ' set the DUT number
    Me.MeasuredDutNumber = a_eventArgs.DutNumber
                                                                      
End Sub

''' <summary>   Notifies all observers of a property change. </summary>
''' <param name="a_propertyName">   [String] the property naMe. </param>
Public Sub OnPropertyChanged(ByVal a_propertyName As String)

    RaiseEvent PropertyChanged( _
            cc_isr_Core.Factory.CreatePropertyChangedEventArgs(a_propertyName))
    ' allow the event to take place.
    DoEvents

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' TIMER CONTROL
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

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

    Dim p_dutNumber As Integer

    ' check if scan or single
    If This.AutoIncrementDutNumberEnabled Then
  
        ' in auto increment mode, use the target DUT number
        p_dutNumber = This.TargetDutNumber
      
    Else

        ' in non auto increment mode, use the target DUT number
        p_dutNumber = This.SelectedDutNumber

    End If
  
    ' allow other events to occur
    VBA.DoEvents

    Dim p_isDone As Boolean: p_isDone = This.StopRequested
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_success As Boolean: p_success = True

    p_success = This.K2700.TryMonitorExternalTrigger(p_dutNumber, p_isDone, p_details)
    
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
        If Not Me.TimerControlled Then _
            Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, _
                "Trigger monitoring done", False, Me.SingleReadEnabled
        
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
    Me.PauseRequested = False
    Me.TimerStarted = True
    Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring, _
        "Trigger monitoring commenced", True, Me.SingleReadEnabled
    If m_eventTimer Is Nothing Then
        Set m_eventTimer = cc_isr_Core.Factory.NewEventTimer.Initialize(a_interval)
        m_eventTimer.StartTimer
    Else
        m_eventTimer.ResumeTimer
    End If

End Sub

''' <summary>   Requests pausing the timer. </summary>
Public Sub PauseTimer()

    Me.PauseRequested = True

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
        Me.OnMeasurementStateChanged cc_isr_Tcp_Scpi.MeasurementModeOption.None, _
            "Trigger monitoring done", False, Me.SingleReadEnabled
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

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Restore Initial State
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initialize the state of the view model executables.  </summary>
Friend Sub InitializeExecutablesState()

    Dim p_connected As Boolean: p_connected = Me.Connected
    Me.OpenConnectionExecutable = Not p_connected
    Me.CloseConnectionExecutable = p_connected
    
    Me.ClearReadingsExecutable = True

    Me.ResetKnownStateExecutable = p_connected
    Me.ClearExecutionStateExecutable = p_connected
    Me.QueryExecutable = p_connected
    Me.ReadExecutable = p_connected
    Me.WriteExecutable = p_connected
    Me.QueryStatusByteExecutable = p_connected
    Me.ReadStandardEventStatusExecutable = p_connected
    Me.QueryInputsExecutable = p_connected
    Me.ImmediateTriggerOptionExecutable = p_connected
    Me.ExternalTriggerOptionExecutable = p_connected

    ' this gets enabled when entering the immediate mode
    Me.MeasureExecutable = False
    
    ' this gets enabled when entering the external trigger mode
    Me.StartMonitoringExecutable = False
    Me.StopMonitoringExecutable = False

    Dim p_gpibLanControlsExecutable As Boolean
    p_gpibLanControlsExecutable = This.Session.GpibLanControllerAttached And p_connected
    Me.SelectiveDeviceClearExecutable = p_gpibLanControlsExecutable
    Me.GoToLocalExecutable = p_gpibLanControlsExecutable
    Me.LocalLockoutExecutable = p_gpibLanControlsExecutable
    Me.AutoAssertTalkSetExecutable = p_gpibLanControlsExecutable
    Me.AutoAssertTalkGetExecutable = p_gpibLanControlsExecutable
    Me.ReadStatusByteExecutable = p_gpibLanControlsExecutable
    Me.ServiceRequestedExecutable = p_gpibLanControlsExecutable
    Me.GpibAddressGetExecutable = p_gpibLanControlsExecutable
    Me.GpibAddressSetExecutable = p_gpibLanControlsExecutable
    Me.GpibLanReadTimeoutGetExecutable = p_gpibLanControlsExecutable
    Me.GpibLanReadTimeoutSetExecutable = p_gpibLanControlsExecutable
    
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
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
Public Function TryReadSessionState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryReadSessionState"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    
    If p_success Then
    
        Dim p_reading As String: p_reading = This.Session.GpibAddressGetter
        
        p_success = 0 < This.Session.PrimaryGpibAddress
        If p_success Then
            Me.GpibAddress = This.Session.PrimaryGpibAddress
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

    ' report the socket address
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
    
    ' initialize the scan list for rear inputs.
    If VBA.vbNullString <> This.RearInputsSenseFunctionName Then
    
        ' set the instrument function scan list
        This.K2700.RouteSystem.DefineFunctionScanLists This.RearInputsSenseFunctionName
    
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
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
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
''' <param name="a_details">             [String] Details why the sense function needs to be restored. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreSenseFunction(ByVal a_senseFunctionName As String, _
    ByRef a_details As String) As Boolean

    If Not ((1 = InStr(1, This.RearInputsSenseFunctionName, a_senseFunctionName, vbTextCompare)) Or _
            (1 = InStr(1, This.FrontInputsSenseFunctionName, a_senseFunctionName, vbTextCompare))) Then
        a_details = "The K2700 device sense function '" & a_senseFunctionName & _
            "' is not expected; it must be either " & _
            This.RearInputsSenseFunctionName & " or " & _
            This.FrontInputsSenseFunctionName & "."
        ShouldRestoreSenseFunction = True
    Else
        ShouldRestoreSenseFunction = False
    End If

End Function

''' <summary>   Returns true and a message if the view model state does not match the expected restore state. </summary>
''' <remarks>   This is the expected state for making measurements. In future releases
''' we could define a set of states. </remarks>
''' <param name="a_details">   [Out, String] Details why the view model needs to be resotred
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
''' <param name="a_details">   [Out, String] Details why the view model needs to be resotred
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

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Socket Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

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
    
    Me.OnMeasurementStateChanged Me.MeasurementMode, "connection changed", Me.Measuring, Me.SingleReadEnabled

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

' + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.DutMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.DutReadingEventArgs"/>] event arguments. </param>
Private Sub m_k2700_DutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

    Const p_procedureName = "m_k2700_DutMeasured"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' issue a beep if reading beep is enabled
    If This.ReadingBeepEnabled Then VBA.Beep

    ' notify the observer of a DUT measurements
    OnDutMeasured a_eventArgs

    If This.AutoIncrementDutNumberEnabled Then

        ' increment the DUT number for the next reading
        
        Me.TargetDutNumber = IIf(a_eventArgs.DutNumber < Me.DutCount, _
            a_eventArgs.DutNumber + 1, 1)

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

