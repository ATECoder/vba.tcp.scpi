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
'''				2700 instrument demo. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Represents an observer pattern subject that can notify observers when a property value changes. </summary>
implements cc_isr_Core.INotifyPropertyChanged

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

Implements cc_isr_Winsock.IConnectable

' reference to the VI session
Private WithEvents m_viSession As cc_isr_Ieee488.ViSession
Attribute m_viSession.VB_VarHelpID = -1

Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private Type ThisData

    ' true when handling the timer event to prevent timer
    ' event handler overrun.
    HandlingTimerEvent As Boolean
    
    Stopwatch As Stopwatch
    
    K2700 As K2700
    
    Device As cc_isr_Ieee488.Device
    
    VISession As cc_isr_Ieee488.VISession
    
	Host As String
	Port As Integer
	SocketAddress As String
	SocketId As Integer
	ReceiveTimeout As Integer
	SessionReadTimeout As Integer

    ' true if a pause is requested so that the timer event is not handled
    ' and the timer is paused on the next event.
    PauseRequested As Boolean
	TimerInterval As Integer
   
	MessageToSend As String
	SentMessage As String
	ReceivedMessage As String
	ReceivedMessageLength As Integer
	ElapsedTime As Integer
	ElapsedTimeFormat As String = "0.0"
	ElapsedTimeCaption As String
	LastErrorMessage As String
	LastMessage As String
	SerialPollByte As Integer
	StatusByte As Integer
	StandardByte As Integer
	ServiceRequested As Boolean
	GpibAddress As Integer
	ReadAfterWriteEnabled As Boolean
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
	ExternalTrigMonitoringEnabled As Boolean
	MonitoringState As String
	AutoIncrementChannelNoEnabled As Boolean
	SingleReadEnabled As Boolean
	SenseFunctions As Collection
	ScpiCommands as Collection
	ImmediateTriggerModeEnabled As Boolean
	ExternalTriggerModeEnabled As Boolean
	AutoStatusRead As Boolean
	ReadingOffset As Double
	
End Type

Private This As ThisData

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.Stopwatch = Nothing
    
    If Not m_eventTimer Is Nothing Then
    
        m_eventTimer.StopTimer
      
        ' allow the timer to stop.
        cc_isr_Core_IO.CoreExtensions.Wait 0.0015 * Me.TimerInterval
        
        Set m_eventTimer = Nothing
        
    End If
    
    If Not m_k2700 Is Nothing Then m_k2700.Dispose
    Set m_k2700 = Nothing
    Set This.K2700 = Nothing
    
    Set This.Device = Nothing
    Set This.VISession = Nothing
	Set m_viSession = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Initialize this worksheet. </summary>
Public Sub Initialize()

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls False
    
    ' populate the combo box commands
    PopulateCommandsComboBox
    
    ' populate the function box commands
    PopulateFunctionComboBox
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize(New ErrTracer)
    Set m_k2700 = This.K2700
    Set This.Device = This.K2700.Device
    Set This.VISession = This.Device.VISession
	Set m_viSession = This.VISession 
    
    Set This.Stopwatch = cc_isr_Core_IO.Factory.NewStopwatch()
  

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' SETTINGS
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Tcp_Scpi.K2700"/>. </summary>
''' <value>   An <see cref="cc_isr_Tcp_Scpi.K2700"/>. </value>
Public Property Get K2700() As cc_isr_Tcp_Scpi.K2700
    Set K2700 = This.K2700
End Property

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.ViSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.ViSession"/>. </value>
Public Property Get ViSession() As cc_isr_Ieee488.ViSession
    Set ViSession = This.ViSession
End Property

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.Device"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.Device"/>. </value>
Public Property Get Device() As cc_isr_Ieee488.Device
    Set Device = This.Device
End Property

''' <summary>   Gets the host dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Integer]. </value>
Public Property Get Port() As Integer
    Port = This.Port
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_value As String)
	If This.SocketAddress <> a_value then
		This.SocketAddress = a_value
		Me.OnPropertyChanged "SocketAddress"
		Me.SocketAddressLength = VBA.Len(a_value)
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
    This.SocketId = a_value
End Property

''' <summary>   Gets the receive timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceiveTimeout() As Integer
    ReceiveTimeout = This.ReceiveTimeout
End Property

''' <summary>   Sets the receive timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let ReceiveTimeout(ByVal a_value As Integer)
	If This.ReceiveTimeout <> a_value then
		This.ReceiveTimeout = a_value
		Me.OnPropertyChanged "ReceiveTimeout"
	End If
End Property

''' <summary>   Gets the session read timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = This.SessionReadTimeout
End Property

''' <summary>   Sets the session read timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionReadTimeout(ByVal a_value As Integer)
    This.SessionReadTimeout = a_value
End Property

''' <summary>   Gets the Pause Requested condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get PauseRequested() As Boolean
    PauseRequested = This.PauseRequested
End Property

''' <summary>   Sets the Pause Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Pause Requested value. </value>
Public Property Let PauseRequested(ByVal a_value As Boolean)
	If This.PauseRequested <> a_value then
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
	If This.TimerInterval <> a_value then
		This.TimerInterval = a_value
		Me.OnPropertyChanged "TimerInterval"
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
	If This.MessageToSend <> a_value then
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
	If This.SentMessage <> a_value then
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
	If This.ReceivedMessage <> a_value then
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
	If This.ReceivedMessageLength <> a_value then
		This.ReceivedMessageLength = a_value
		Me.OnPropertyChanged "ReceivedMessageLength"
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
	If This.ElapsedTime <> a_value then
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
	If This.ElapsedTimeFormat <> a_value then
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
	If This.ElapsedTimeCaption <> a_value then
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
	If This.LastErrorMessage <> a_value then
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
	If This.LastMessage <> a_value then
		This.LastMessage = a_value
		Me.OnPropertyChanged "LastMessage"
	End If
End Property

''' <summary>   Gets the Serial Poll Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get SerialPollByte() As Integer
    SerialPollByte = This.SerialPollByte
End Property

''' <summary>   Sets the Serial Poll Byte. </summary>
''' <param name="a_value">   [Integer] the Serial Poll Byte value. </value>
Public Property Let SerialPollByte(ByVal a_value As Integer)
	If This.SerialPollByte <> a_value then
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
	If This.ServiceRequested <> a_value then
		This.ServiceRequested = a_value
		Me.OnPropertyChanged "ServiceRequested"
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
	If This.StatusByte <> a_value then
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
	If This.StandardByte <> a_value then
		This.StandardByte = a_value
		Me.OnPropertyChanged "StandardByte"
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
	If This.GpibAddress <> a_value then
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
	If This.ReadAfterWriteEnabled <> a_value then
		This.ReadAfterWriteEnabled = a_value
		Me.OnPropertyChanged "ReadAfterWriteEnabled"
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
	If This.TopCard <> a_value then
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
	If This.TopCardFunctionScanList <> a_value then
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
	If This.BottomCard <> a_value then
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
	If This.BottomCardFunctionScanList <> a_value then
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
	If This.FrontInputsRequired <> a_value then
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
	If This.FrontInputsValue <> a_value then
		This.FrontInputsValue = a_value
		Me.OnPropertyChanged "FrontInputsValue"
		me.FrontInputsCaption = IIf(a_value, "Front", "Rear")
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
	If This.FrontInputsCaption <> a_value then
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
	If This.ChannelCount <> a_value then
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
	If This.SelectedChannelNumber <> a_value then
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
	If This.CurrentChannelNumber <> a_value then
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
	If This.MeasuredChannelNumber <> a_value then
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
	If This.SenseFunctionName <> a_value then
		This.SenseFunctionName = a_value
		Me.OnPropertyChanged "SenseFunctionName"
		Me.OnSenseFunctionChanged
	End If
End Property

''' <summary>   Handles the sense function change. </summary>
Public Sub OnSenseFunctionChanged()

    Select Case Me.SenseFunctionName
    
        Case "CONT"
        
            Me.SenseFunctionCaption =  "Continuity"
            Me.SenseFunctionUnit =  "T/F"
        
        Case "CURR:AC"
        
            Me.SenseFunctionCaption =  "Current AC"
            Me.SenseFunctionUnit =  "A"
        
        Case "CURR:DC"
        
            Me.SenseFunctionCaption =  "Current DC"
            Me.SenseFunctionUnit =  "A"
        
        Case "FREQ"
        
            Me.SenseFunctionCaption =  "Frequency"
            Me.SenseFunctionUnit =  "Hz"
        
        Case "FRES"
        
            Me.SenseFunctionCaption =  "4-Wire Res."
            Me.SenseFunctionUnit =  VBA.Chr$(234)
        
        Case "PER"
        
            Me.SenseFunctionCaption =  "Period"
            Me.SenseFunctionUnit =  "s"
        
        Case "RES"
        
            Me.SenseFunctionCaption =  "2-Wire Res."
            Me.SenseFunctionUnit =  VBA.Chr$(234)
        
        Case "TEMP"
        
            Me.SenseFunctionCaption =  "Temperature"
            Me.SenseFunctionUnit =  "K"
        
        Case "VOLT:AC"
        
            Me.SenseFunctionCaption =  "Voltage AC"
            Me.SenseFunctionUnit =  "v"
        
        Case "VOLT:DC"
        
            Me.SenseFunctionCaption =  "Voltage DC"
            Me.SenseFunctionUnit =  "v"
        
        Case Else
        
            Me.SenseFunctionCaption =  "unknown"
            Me.SenseFunctionUnit =  "?"
    
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
	If This.SenseFunctionCaption <> a_value then
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
	If This.SenseFunctionUnit <> a_value then
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
	If This.MeasuredReading <> a_value then
		This.MeasuredReading = a_value
		Me.OnPropertyChanged "MeasuredReading"
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
	If This.MeasuredValueFormat <> a_value then
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
	If This.MeasuredValue <> a_value then
		This.MeasuredValue = a_value
		Me.OnPropertyChanged "MeasuredValue"
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
	If This.MonitoringState <> a_value then
		This.MonitoringState = a_value
		Me.OnPropertyChanged "MonitoringState"
	End If
End Property

''' <summary>   Gets the Sense Functions. </summary>
''' <value>   [Collection]. </value>
Public Property Get SenseFunctions() As Collection
    SenseFunctions = This.SenseFunctions
End Property

''' <summary>   Sets the Sense Functions. </summary>
''' <param name="a_value">   [Collection] the Sense Functions value. </value>
Public Property Let SenseFunctions(ByVal a_value As Collection)
	If This.SenseFunctions <> a_value then
		This.SenseFunctions = a_value
		Me.OnPropertyChanged "SenseFunctions"
	End If
End Property

''' <summary>   Gets the Scpi Commands. </summary>
''' <value>   [Collection]. </value>
Public Property Get ScpiCommands() As Collection
    ScpiCommands = This.ScpiCommands
End Property

''' <summary>   Sets the Scpi Commands. </summary>
''' <param name="a_value">   [Collection] the Scpi Commands value. </value>
Public Property Let ScpiCommands(ByVal a_value As Collection)
	If This.ScpiCommands <> a_value then
		This.ScpiCommands = a_value
		Me.OnPropertyChanged "ScpiCommands"
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
	If This.ExternalTrigMonitoringEnabled <> a_value then
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
	If This.AutoIncrementChannelNoEnabled <> a_value then
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
	If This.SingleReadEnabled <> a_value then
		This.SingleReadEnabled = a_value
		Me.OnPropertyChanged "SingleReadEnabled"
	End If
End Property

''' <summary>   Gets the Immediate trigger mode Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ImmediateTriggerModeEnabled() As Boolean
    ImmediateTriggerModeEnabled = This.ImmediateTriggerModeEnabled
End Property

''' <summary>   Sets the Immediate trigger mode Enabled. </summary>
''' <param name="a_value">   [Boolean] the Immediate trigger mode Enabled value. </value>
Public Property Let ImmediateTriggerModeEnabled(ByVal a_value As Boolean)
	If This.ImmediateTriggerModeEnabled <> a_value then
		This.ImmediateTriggerModeEnabled = a_value
		Me.OnPropertyChanged "ImmediateTriggerModeEnabled"
	End If
End Property

''' <summary>   Gets the External trigger mode Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTriggerModeEnabled() As Boolean
    ExternalTriggerModeEnabled = This.ExternalTriggerModeEnabled
End Property

''' <summary>   Sets the External trigger mode Enabled. </summary>
''' <param name="a_value">   [Boolean] the External trigger mode Enabled value. </value>
Public Property Let ExternalTriggerModeEnabled(ByVal a_value As Boolean)
	If This.ExternalTriggerModeEnabled <> a_value then
		This.ExternalTriggerModeEnabled = a_value
		Me.OnPropertyChanged "ExternalTriggerModeEnabled"
	End If
End Property

''' <summary>   Gets the Auto Status Read Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoStatusReadEnabled() As Boolean
    AutoStatusReadEnabled = This.AutoStatusReadEnabled
End Property

''' <summary>   Sets the Auto Status Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Status Read Enabled value. </value>
Public Property Let AutoStatusReadEnabled(ByVal a_value As Boolean)
	If This.AutoStatusReadEnabled <> a_value then
		This.AutoStatusReadEnabled = a_value
		Me.OnPropertyChanged "AutoStatusReadEnabled"
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
	If This.ReadingOffset <> a_value then
		This.ReadingOffset = a_value
		Me.OnPropertyChanged "ReadingOffset"
	End If
End Property

''' <summary>   Populates the collection of sense functions. </summary>
Public Sub PopulateSenseFunctions()

	Dim p_senseFunctions as New Collection
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

''' <summary>   Populates the Commands collection. </summary>
Public Sub PopulateScpiCommands()

	Dim p_scpiCommands as New Collection
    p_scpiCommands.add cc_isr_Ieee488.ClearExecutionStateCommand
    p_scpiCommands.add cc_isr_Ieee488.ClearExecutionStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.IdentityQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.OperationCompleteCommand
    p_scpiCommands.add cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.OptionsQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.ResetKnownStateCommand
    p_scpiCommands.add cc_isr_Ieee488.ResetKnownStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardEventEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.StandardEventEnableQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.StandardEventStatusQueryCommand
    p_scpiCommands.add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.ServiceRequestEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.ServiceRequestQueryCommand
    p_scpiCommands.add cc_isr_Ieee488.WaitCommand
    p_scpiCommands.add cc_isr_Ieee488.WaitCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
   
    p_scpiCommands.add "SYST:BEEP OFF"
    p_scpiCommands.add "SYST:CLE"
    p_scpiCommands.add "SYST:ERR?"
    p_scpiCommands.add "SYST:FRSW?"
    p_scpiCommands.add "FUNC 'RES',(@101:120)"
    p_scpiCommands.add "FUNC 'RES',(@201:220)"
    p_scpiCommands.add "FORM:ELEM READ"
    p_scpiCommands.add "STAT:PRES"
    p_scpiCommands.add "STAT:MEAS:ENAB 32"
    p_scpiCommands.add "FETCH?"
    p_scpiCommands.add "READ?"

End sub 

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Public Sub RestartStopwatch()
    
    Me.ElapsedTimeCaption = vbNullString
    This.Stopwatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Public Sub ReadStopwatch()
    
	me.ElapsedTime = This.Stopwatch.ElapsedMilliseconds

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Writes the <see cref="MessageToSend"/> and update elapsed times. </summary>
Public Sub WriteCommand() 

    Const p_procedureName = "WriteCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatch
    
    This.VISession.WriteLine me.MessageToSend
    
    Me.ReadStopwatch
	
	Me.SentMessage = me.MessageToSend

	If Me.AutoStatusReadEnabled Then
	
		If This.VISession.UsingGpibLan Then _
			Me.SerialPollByte = This.VISession.GpibLan.SerialPoll
	
		Dim p_isQuery As Boolean
		p_isQuery = cc_isr_Core.StringExtensions.EndsWith(p_messageToSend, "?")
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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
End Sub

''' <summary>   Query the instrument using the message to send and update then elapsed time. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Public Sub QueryCommand() 

    Const p_procedureName = "QueryCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatch
    
    Me.ReceivedMessage = This.VISession.QueryLine (Me.MessageToSend)
    
    Me.ReadStopwatch
	
	Me.SentMessage = me.MessageToSend
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
	
End Function

''' <summary>   Read from the instrument updating the elap[es times. </summary>
Public Sub ReadCommand()

    Const p_procedureName = "ReadCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatch
    
    Me.ReceivedMessage = This.VISession.Read()
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	

End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
''' <param name=a_connect>   [Boolean] true to connect; otherwise, disconnect. </param>
Public Sub ToggleConnectionCommand(ByVal a_connect as Boolean)

    Const p_procedureName = "ToggleConnectionCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Me.RestartStopwatch

    Dim p_host As String: p_host = Me.Host
    Dim p_port As Integer: p_port = Me.port
    Dim p_sessionReadTimeout As Integer: p_sessionReadTimeout = Me.SessionReadTimeout
      
    If Not a_connect And This.Connected Then
        
        This.K2700.CloseConnection
        
    ElseIf a_connect And Not This.Connected Then
    
        This.K2700.OpenConnection p_host, p_port
        
        This.VISession.SessionReadTimeout = p_sessionReadTimeout
        
    End If
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
	
End Sub

''' <summary>   Handles the Clear Known State button click event. </summary>
Public Sub ClearExecutionStateCommand()
    
    Const p_procedureName = "ClearExecutionStateCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.LastMessage = "Clearing execution state. "
    
    Me.RestartStopwatch
    
    This.Device.ClearExecutionState
    
    Me.ReadStopwatch
    
    Me.LastMessage = "Done clearing execution state. "
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
	
End Sub

''' <summary>   Returns true if the inputs are valid for front or rear settings. </summary>
Public Function ValidateInputs() As Boolean

    ' check if front or rear
    If Me.FrontReadValue And Not This.K2700.IsFrontInputs Then
    
        ' display the proper message
        Me.LastMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."
    
        ValidateInputs = False
        
    ElseIf Not Me.FrontReadValue And This.K2700.IsFrontInputs Then
    
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

	Dim p_success as boolean 

    ' exit if inputs are not in the correct settings
	
	p_success = Me.ValidateInputs()
	
    If p_success Then 
   
		' clear execution state when switching mode.
    
		This.K2700.Device.ClearExecutionState
   
		' Configure external triggering reading requiring the front panel
    
		This.K2700.ConfigureExternalTriggerReading Me.FunctionComboBox.Text, _
				Me.FrontReadValue
    
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
		ImmediateTriggerModeEnabled = false
		
		' enable toggling the external trigger monitoring.
		me.ExternalTriggerModeEnabled = true
	
	End If
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
	Dim p_success as Boolean
	
    ' exit if inputs are not in the correct settings

	p_success = Me.ValidateInputs()
	
    If  p_Success Then
    
		' clear execution state when switching mode.
    
		This.K2700.Device.ClearExecutionState
    
		' configure measuring a single value on each read.
		This.K2700.ConfigureSingleFunctionReading Me.SenseFunctionName
	
	End If
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStandardEventStatusCommand() 

    Const p_procedureName = "ReadStandardEventStatusCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.StandardByte = This.Device.QueryStandardEventsStatus
    
    Me.ReadStopwatch
	
	ReadStandardEventStatus = Me.StandardByte

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
    Me.RestartStopwatch
    
    This.Device.ResetKnownState
    
    Me.ReadStopwatch
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
	
End Sub

''' <summary>   Handles the Selective Device Clear button click event. </summary>
Public Sub SelectiveDeviceClearCommand() 

    Const p_procedureName = "SelectiveDeviceClearCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.SelectiveDeviceClear
    
    Me.ReadStopwatch
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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

    Me.RestartStopwatch

    Me.StatusByte = This.Device.QueryServiceRequestStatus()
    
    Me.ReadStopwatch
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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

    Me.RestartStopwatch

    Me.StatusByte = This.Device.ReadStatusByte( True ) 
    
    Me.ReadStopwatch
	
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB LAN Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Write a message to the device via the GPIB Lan. </summary>
Public Sub QueryGpibLanCommand() 

    Const p_procedureName = "QueryGpibLanCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    If 0 < This.VISession.GpibLan.SendToController(Me.MessageToSend) Then
        Me.ReceivedMessage = This.VISession.GpibLan.ReceiveFromDevice()
    End If
    
    Me.ReadStopwatch
    
	Me.SentMessage = me.MessageToSend

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
    Me.RestartStopwatch
    
    Me.MessageReceived = This.VISession.GpibLan.ReceiveFromDevice()
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Write a message to the device via the GPIB Lan. </summary>
Public Sub  WriteGpibLanCommand()

    Const p_procedureName = "WriteGpibLanCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.SendToController(me.MessageToSend)
    
    Me.ReadStopwatch
    
	Me.SentMessage = me.MessageToSend
	
	If Me.AutoStatusReadEnabled Then
	
		If This.VISession.UsingGpibLan Then _
			Me.SerialPollByte = This.VISession.GpibLan.SerialPoll
	
	End If
	

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Go To Local button click event. </summary>
Public Sub GoToLocalCommand()

    Const p_procedureName = "GoToLocalCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.GoToLocal
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutCommand()

    Const p_procedureName = "LocalLockoutCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.LocalLockout
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write set button click event. </summary>
Private Sub ReadAfterWriteSetCommand()

    Const p_procedureName = "ReadAfterWriteSetCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.ReadAfterWriteEnabledSetter Me.ReadAfterWriteEnabled
    
    Me.ReadStopwatch
    
    This.VISession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write get button click event. </summary>
Private Sub ReadAfterWriteGetCommand()

    Const p_procedureName = "ReadAfterWriteGetCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.ReadAfterWriteEnabled = This.VISession.GpibLan.ReadAfterWriteEnabledGetter
    
    Me.ReadStopwatch
    
    This.VISession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQQueryCommand()

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
    
    Me.RestartStopwatch
    
    Me.GpibAddress = This.VISession.GpibLan.GpibAddressGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.GpibAddressSetter Me.GpibAddress

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
    Me.RestartStopwatch
    
    Me.ReadTimeout = This.VISession.GpibLan.ReadTimeoutGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.ReadTimeoutSetter Me.ReadTimeout
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' 2700 Measurements Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsCommand()

    Const p_procedureName = "QueryInputsCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.FrontInputsValue = This.K2700.IsFrontInputs
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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

    If Me.AutoIncrementChannelNoEnabled Then
        
        ' Clear all cells
        Me.ClearReadings

        ' scan all channels
        For p_channelNumber = 1 To Me.ChannelCount
        
            ' allow events to occur
            VBA.DoEvents
            
            ' set the current channel number
            Me.CurrentChannelNumber = p_channelNumber
            
            Me.RestartStopwatch
            If Me.FrontReadValue Then
                This.K2700.TakeFrontReading p_channelNumber
            Else
                This.K2700.TakeRearReading p_channelNumber
            End If
            Me.ReadStopwatch
        
        Next p_channelNumber

    Else
    
        p_channelNumber = Me.SelectedChannelNumber
        
        Me.RestartStopwatch
        If Me.FrontReadValue Then
            This.K2700.TakeFrontReading p_channelNumber
        Else
            This.K2700.TakeRearReading p_channelNumber
        End If
        
        Me.ReadStopwatch
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
        SetReading p_channelNumber, Me.ReadingOffset
    Next p_channelNumber
  
End Sub

''' <summary> Toggles monitoring the instrument for external triggered measurements. </summary>
Public Sub StartMonitoringExternalTriggersCommand()

    Const p_procedureName = "StartMonitoringExternalTriggersCommand"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If Me.AutoIncrementChannelNoEnabled Then
        
        ' Clear all readings
        Me.ClearReadings
    
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
    This.VISession.WriteLine ":INIT; *CLS"
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
		
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
    
        ' EnableButtons
        Range(m_externalTriggerStatusCell) = "External Trigger off"
        
        Me.MonitorExtTrigToggleButton.Enabled = True
        Me.MeasureButton.Enabled = True
        
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
    
    ' enable all buttons
    Me.MonitorExtTrigToggleButton.Enabled = True
    Me.MeasureButton.Enabled = True
    
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_session_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        SetReading a_eventArgs.ChannelNumber, -1#
    Else
        SetReading a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.K2700.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName = "m_session_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls a_eventArgs.Connected

    If a_eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).Value = This.VISession.SocketAddress
        Range(m_socketIdCell).Value = This.VISession.SocketId
        
        If This.VISession.UsingGpibLan Then
            Me.ReadAfterWriteEnabled = This.VISession.GpibLan.ReadAfterWriteEnabledGetter
            If Me.ReadAfterWriteEnabled Then
                ' for the 2700 with needs this off
                This.VISession.GpibLan.ReadAfterWriteEnabledSetter False
                Me.ReadAfterWriteEnabled = This.VISession.GpibLan.ReadAfterWriteEnabledGetter
            End If
        End If
        
        This.VISession.GpibLan.SelectiveDeviceClear
        This.Device.ResetKnownState
        This.Device.ClearExecutionState
        
        Range(m_card1Cell).Value = This.K2700.CardName(1)
        Range(m_card2Cell).Value = This.K2700.CardName(2)
        
        This.K2700.DefineFunctionScanLists Me.FunctionComboBox.Text
        
        Range(m_card1ScanListCell).Value = This.K2700.FunctionScanList(1)
        Range(m_card2ScanListCell).Value = This.K2700.FunctionScanList(2)
    
        ' clear any errors
        This.Device.ClearExecutionState

    Else
    
        Range(m_socketAddressCell).Value = vbNullString
        Range(m_socketIdCell).Value = vbNullString
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
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
' Notify property changed
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Notifies all observers of a property change. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Public Sub OnPropertyChanged(ByVal a_propertyName as String)
	
    RaiseEvent PropertyChanged( _
            cc_isr_Core.Factory.CreatePropertyChangedEventArgs(a_propertyName))
	
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
    Set Connectable = This.Device.ViSession
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

    Set m_viSession = Nothing

    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set This.RouteSystem = Nothing
    Set This.ScpiSystem = Nothing
    Set This.ViSession = Nothing
    
    If Not This.Device Is Nothing Then
        This.Device.Dispose
    End If
    This.Device = Nothing
    
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
        
		If me.connected then
		
			If This.ViSession.UsingGpibLan Then
				If This.ViSession.GpibLan.ReadAfterWriteEnabledGetter Then
					' for the 2700 with needs this off
					This.ViSession.GpibLan.ReadAfterWriteEnabledSetter False
				End If
			End If
			
			Dim p_installedCardsCount As Integer
			p_installedCardsCount = Me.ReadCards()

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
'  IEEE 488 Session Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  CONTROLS AND CONTROL EVENTS TO MOVE to the View 
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Populates the Function combo box. </summary>
Public Sub PopulateFunctionComboBox()

    Const p_procedureName = "PopulateFunctionComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

	Dim p_fourWireResistanceIndex as Integer
    Me.FunctionComboBox.Clear
	Dim p_senseFunction As String
	for each p_senseFunction in This.K2700ViewModel.SenseFunctions
		Me.FunctionComboBox.AddItem p_senseFunction
		If "FRES" = p_senseFunction Then 
			p_fourWireResistanceIndex = me.FunctionComboBox.Count 
		End If
	Next

    Me.FunctionComboBox.ListIndex = p_fourWireResistanceIndex
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Populates the Commands combo box. </summary>
Public Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.CommandsComboBox.Clear
	Dim p_scpiCommand As String
	for each p_scpiCommand in This.K2700ViewModel.ScpiCommands
		Me.CommandsComboBox.AddItem p_scpiCommand
	Next
    
    Me.CommandsComboBox.ListIndex = 1
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
Private Sub ImmediateTriggerOptionButton_Click()

	This.K2700ViewModel.ConfigureImmediateTriggerReadingsCommand()

	' these should get handled on property changes.
	
    Me.EnableMeasureButton
    Me.EnableMonitorExtTrigToggleButton

End Sub

''' <summary>   Toggle the enable state of the controls. </summary>
''' <param name="a_connected">   [Boolean] True if connected; otherwise, false. </param>
Public Sub EnableControls(ByVal a_connected As Boolean)

    Me.ConnectToggleButton.Value = a_connected
    
    ' report the connection state
    Range(m_connectedCell).Value = a_connected
        
    Me.RSTButton.Enabled = a_connected
    Me.CLSButton.Enabled = a_connected
    Me.QueryButton.Enabled = a_connected
    Me.ReadButton.Enabled = a_connected
    Me.ReadStatusByteButton.Enabled = a_connected
    Me.ReadStandardEventStatusButton.Enabled = a_connected
    Me.WriteButton.Enabled = a_connected
    
    Me.QueryInputsButton.Enabled = a_connected
    
    Me.EnableMeasureButton
    Me.EnableMonitorExtTrigToggleButton
    
    If a_connected Then
        Me.EnableGpibLanControls This.VISession.UsingGpibLan
    Else
        Me.EnableGpibLanControls False
    End If

End Sub

''' <summary>   Toggle the enable state of the Measure Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Public Sub EnableMeasureButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeImmediateMeasurement
    End If
    p_enable = This.K2700ViewModel.ImmediateTriggerModeEnabled
    
    Me.ImmediateTriggerOptionButton.Enabled = p_enable
    Me.MeasureButton.Enabled = False

End Sub

''' <summary>   Toggle the enable state of the External Trig Toggle Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Public Sub EnableMonitorExtTrigToggleButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeExternalTriggers
    End If
    p_enable = This.K2700ViewModel.ExternalTriggerModeEnabled
    
    Me.ExternalTriggerOptionButton.Enabled = p_enable
    Me.MonitorExtTrigToggleButton.Enabled = False

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

	This.K2700ViewModel.ReadStandardEventStatusCommand
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()

	This.K2700ViewModel.ResetKnowStateCommand
    
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
	This.K2700ViewModel.SelectiveDeviceClearCommand
	
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStatusByteButton_Click()
    
	This.K2700ViewModel.QueryStatusByteCommand
    
End Sub

''' <summary>   Handles the Commands Combo Box change event. </summary>
Private Sub CommandsComboBox_Change()
	This.K2700ViewModel.MessageToSend = Me.CommandsComboBox.Text
End Sub

''' <summary>   Handles the Auto status read check box checked changed event. </summary>
Private Sub AutoStatusReadCheckBox_CheckedChanged()
	This.K2700ViewModel.AutoStatusReadEnabled = Me.AutoStatusReadCheckBox.Value
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

	This.K2700ViewModel.WriteCommand
    
End Sub


''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

	This.K2700ViewModel.ReadCommand
	
End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
	This.K2700ViewModel.QueryCommand

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB LAN
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Toggle the enabled state of the GPIB Lan controls. </summary>
''' <param name="a_usingGpibLanController">   [Boolean] True if using the GPIB-Lan controller. </param>
Public Sub EnableGpibLanControls(ByVal a_usingGpibLanController As Boolean)
    Me.SDCButton.Enabled = a_usingGpibLanController
    Me.GoToLocalButton.Enabled = a_usingGpibLanController
    Me.LocalLockoutButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteSetButton.Enabled = a_usingGpibLanController
    Me.ReadAfterWriteGetButton.Enabled = a_usingGpibLanController
    Me.SerialPollButton.Enabled = a_usingGpibLanController
    Me.SRQButton.Enabled = a_usingGpibLanController
    Me.GetGpibAddressButton.Enabled = a_usingGpibLanController
    Me.SetGpibAddressButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutGetButton.Enabled = a_usingGpibLanController
    Me.ReadTimeoutSetButton.Enabled = a_usingGpibLanController
End Sub

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()

	This.K2700ViewModel.GoToLocalCommand
	
end sub 

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()

	This.K2700ViewModel.LocalLockoutCommand

End Sub 


''' <summary>   Handles the read after write set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()

	This.K2700ViewModel.ReadAfterWriteSetCommand

End Sub 

''' <summary>   Handles the read after write get button click event. </summary>
Private Sub ReadAfterWriteGetButton_Click()

	This.K2700ViewModel.ReadAfterWriteGetCommand

End Sub 

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()

	This.K2700ViewModel.SerialPollCommand

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()

	This.K2700ViewModel.SRQQueryCommand

End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GetGpibAddressButton_Click()

	This.K2700ViewModel.GpibAddressGetCommand

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Private Sub SetGpibAddressButton_Click()

	This.K2700ViewModel.GpibAddressSetCommand

End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()

	This.K2700ViewModel.ReadTimeoutGetCommand

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()

	This.K2700ViewModel.ReadTimeoutSetCommand

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' 2700 Measurements Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the measurement display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Private Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)

    Range(m_channelNumberCell).Value = a_channelNo
    Range(m_functionValueCell).Value = a_value
    
    Dim p_rowNo As Integer: p_rowNo = m_measurementTopRow + a_channelNo - 1
    Range(m_measurementNoCol & CStr(p_rowNo)).Value = CStr(a_channelNo)
    Range(m_measurementValueCol & CStr(p_rowNo)).Value = Format(a_value, Me.ReadingFormat)
    
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

	This.K2700ViewModel.QueryInputsCommand

End Sub


''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()
  
    Const p_procedureName = "MeasureButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
	' TO-DO: See how this can be simplified.  
	' Can we toggle to enable status when the command is active.
	
    ' exit if inputs are not in the correct settings
    If Not Me.ValidateInputs() Then GoTo exit_Handler
    
    ' disable the relevant buttons.
    
    Me.MeasureButton.Enabled = False
    Me.ImmediateTriggerOptionButton.Enabled = False
    Me.ExternalTriggerOptionButton.Enabled = False
    
    ' make an immediate measure of one or all channels.
    
    This.K2700ViewModel.MeasureImmediatelyCommand

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.MeasureButton.Enabled = True
    Me.ImmediateTriggerOptionButton.Enabled = True
    Me.ExternalTriggerOptionButton.Enabled = True
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub MonitorExtTrigToggleButton_Click()

    Const p_procedureName = "MonitorExtTrigToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Me.ExternalTrigMonitoringEnabled Then
    
        Me.ImmediateTriggerOptionButton.Enabled = False
        Me.ExternalTriggerOptionButton.Enabled = False
    
        Me.StartMonitoringExternalTriggersCommand
    
    Else
    
        ' one external triggering is turned off, the timer
        ' is allowed to turn itself off
        Me.ImmediateTriggerOptionButton.Enabled = True
        Me.ExternalTriggerOptionButton.Enabled = True
    
    End If
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.EnableMeasureButton
    Me.EnableMonitorExtTrigToggleButton

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:

    ' stop the timer if the timer started.
    Me.PauseTimer
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub
