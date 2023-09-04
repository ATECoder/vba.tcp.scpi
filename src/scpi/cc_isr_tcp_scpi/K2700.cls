VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cc_isr_Winsock.IConnectable

' references the event handler of the VI session
Private WithEvents m_viSession As cc_isr_Ieee488.ViSession
Attribute m_viSession.VB_VarHelpID = -1

Private Type this_
    Device As cc_isr_Ieee488.Device
    ViSession As cc_isr_Ieee488.ViSession
    DeviceErrorTracer As IDeviceErrorTracer
    RouteSystem As RouteSystem
    ScpiSystem As ScpiSystem
    SenseSystem As K2700SenseSystem
    TriggerSystem As TriggerSystem
    ExtTrigInitiated As Boolean
End Type

Private This As this_

''' <summary>   The event that is raised upon completion of a channel measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="ChannelReadingEventArgs"/> event args. </param>
Public Event ChannelMeasured(ByVal a_eventArgs As ChannelReadingEventArgs)

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Initializes the k2700 . </summary>
Private Sub Class_Initialize()

    Set This.Device = cc_isr_Ieee488.Factory.NewDevice()
    Set This.ViSession = This.Device.ViSession
    
    ' make sure to turn off assering talk after write.
    This.ViSession.GpibLan.AssertTalkAfterWriteAllowed = False
    
    Set This.DeviceErrorTracer = New K2700DeviceErrorTracer
    This.DeviceErrorTracer.Initialize This.ViSession
    
    Set This.ScpiSystem = New ScpiSystem
    This.ScpiSystem.Initialize This.ViSession
    
    Dim p_senseSystem As SenseSystem
    Set p_senseSystem = New SenseSystem
    p_senseSystem.Initialize This.ViSession
    Set This.SenseSystem = New K2700SenseSystem
    This.SenseSystem.Initialize p_senseSystem
    
    Set This.RouteSystem = New RouteSystem
    This.RouteSystem.Initialize This.ViSession

    Set This.TriggerSystem = New TriggerSystem
    This.TriggerSystem.Initialize This.ViSession

End Sub

''' <summary>   Initializes and returns the <see cref="Ieee488Session"/>. </summary>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function Initialize() As K2700
   
    This.Device.Initialize
    
    Set m_viSession = This.Device.ViSession
    
    Set Initialize = Me
   
End Function

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.ViSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.ViSession"/>. </returns>
Public Property Get ViSession() As cc_isr_Ieee488.ViSession
    Set ViSession = This.Device.ViSession
End Property

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.Device"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.Device"/>. </returns>
Public Property Get Device() As cc_isr_Ieee488.Device
    Set Device = This.Device
End Property

''' <summary>   Returns a reference to the <see cref="IDeviceErrorTracer"/> subsystem. </summary>
''' <value>   An <see cref="IDeviceErrorTracer"/>. </returns>
Public Property Get DeviceErrorTracer() As IDeviceErrorTracer
    Set DeviceErrorTracer = This.DeviceErrorTracer
End Property

''' <summary>   Returns a reference to the <see cref="ScpiSystem"/> subsystem. </summary>
''' <value>   An <see cref="ScpiSystem"/>. </returns>
Public Property Get ScpiSystem() As ScpiSystem
    Set ScpiSystem = This.ScpiSystem
End Property

''' <summary>   Returns a reference to the <see cref="SenseSystem"/> subsystem. </summary>
''' <value>   An <see cref="SenseSystem"/>. </returns>
Public Property Get SenseSystem() As SenseSystem
    Set SenseSystem = This.SenseSystem
End Property

''' <summary>   Initializes the k2700 legacy mode. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_function">    [Optional, 'FRES'] The measurement function. </param>
''' <param name="a_timeoutMs">   [Optional, 3000] The session read timeout in milliseconds. </param>
''' <returns>   True if the device was initialized. </returns>
Public Function InitializeLegacy(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_function As String = "RES", _
        Optional ByVal a_timeoutMs As Long = 3000) As Boolean

    Me.OpenConnection a_host, a_port

    ' reset clear the device.
    If This.ViSession.UsingGpibLan Then This.ViSession.GpibLan.SelectiveDeviceClear
    This.ViSession.SessionReadTimeout = a_timeoutMs
    This.Device.ResetKnownState
    This.Device.ClearExecutionState
    
    ' read the instrument options and parse the card names and types.
    
    Me.ReadCards
    
    ' define function scan list for measurements on all cards.
   
    Me.DefineFunctionScanLists a_function
   
    ' turn on continuous auto range
    Me.ConfigureContinuousFunctionAutoRange a_function
    
    InitializeLegacy = True
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  2700 Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Take measurements
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Performs a single read from the front panel. </summary>
''' <param name="a_channelNumber">   [Integer] The channel number. </param>
''' <returns>   [String] the reading. </returns>
Public Function TakeFrontReading(Optional a_channelNumber As Integer = 0) As String

    This.ViSession.WriteLine ":INIT"
    This.ViSession.WriteLine ":READ?"
    
    Dim p_reading As String: p_reading = This.ViSession.Read
    
    OnChannelMeasured a_channelNumber, p_reading

    TakeFrontReading = p_reading

End Function

''' <summary>   Performs a single scan on the given channel number from the rear panel. </summary>
''' <remarks>
''' Use this method to get a single value from the instrument using the back panel switches
''' </remarks>
''' <param name="a_channelNumber">   The device channel number. </param>
''' <returns>    [String] the reading of empty if the channel number is out
''' of range for this device. </returns>
Public Function TakeRearReading(ByVal a_channelNumber As Integer) As String

    ' get the channel close command
    Dim p_routeCloseCommand As String
    On Error Resume Next
    p_routeCloseCommand = This.RouteSystem.GetRouteCloseCommand(a_channelNumber)
    On Error GoTo 0
    
    Dim p_reading As String
    p_reading = VBA.vbNullString
    
    If VBA.Len(p_routeCloseCommand) > 0 Then
    
        ' close the rear channel
        This.ViSession.WriteLine p_routeCloseCommand
        This.ViSession.WriteLine ":INIT"
        This.ViSession.WriteLine ":READ?"
        
        p_reading = This.ViSession.Read
        
        OnChannelMeasured a_channelNumber, p_reading
    
    End If

    TakeRearReading = p_reading

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Continuous and Immediate modes
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the instrument to continuous auto range reading. </summary>
''' <param name="a_function">    [String] The measurement function. </param>
Public Sub ConfigureContinuousFunctionAutoRange(ByVal a_function As String)

    This.Device.ResetKnownState
    This.Device.ClearExecutionState
    
    ' set function defaults
    ' This.ViSession.WriteLine ":RES:RANG:AUTO ON"
    ' This.ViSession.WriteLine ":RES:NPLC 1"
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.ViSession.WriteLine ":FUNC '" & a_function & "'"
    
    ' turn on continuous mode
    This.ViSession.WriteLine ":" & a_function & ":RANG:AUTO ON"
    This.ViSession.WriteLine ":" & a_function & ":NPLC 1"
    This.ViSession.WriteLine ":INIT:CONT ON"

End Sub

''' <summary>   Restore continuous reading. </summary>
''' <remarks>   This is typically called to restore continuous mode without
'''             chaning the function mode. </remarks>
Public Sub RestoreContinuousReading()

    This.ViSession.WriteLine ":INIT:CONT OFF"
    
    ' set the trigger source to immediate.
    This.TriggerSystem.SourceSetter TriggerSource.Immediate

    This.ViSession.WriteLine ":INIT:CONT ON"

End Sub

''' <summary>   Configures single function reading. </summary>
''' <param name="a_function">    [String] The measurement function. </param>
Public Sub ConfigureSingleFunctionReading(ByVal a_function As String)

    ' reset is not applied on each configuration because it takes too long to complete.
    ' This.ViSession.WriteLine "*RST"
    
    ' set the trigger source to immediate.
    ' This.ViSession.WriteLine ":TRIG:SOUR IMM"
    This.TriggerSystem.SourceSetter TriggerSource.Immediate
    This.ViSession.WriteLine ":INIT:CONT OFF"
    This.ViSession.WriteLine ":SAMP:COUN 1"
    This.ViSession.WriteLine ":TRIG:COUN 1"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.ViSession.WriteLine ":FUNC '" & a_function & "'"
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  External trigger source
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the instrument to detect the measurement available bit
'''             of the status register upon receiving an external trigger. </summary>
Public Sub EnableMeasurementAvailableMonitoring()

    ' setup the status register so that we get a RQS
    ' on measurement available
    
    This.ViSession.WriteLine "*ESE 0"
    This.ViSession.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    
    This.ViSession.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    
    This.ViSession.WriteLine "STAT:MEAS:ENAB 32"
    
#If False Then
    ' legacy code used the operation register going to idle to indicate that
    ' a measurement is available. This did not work when testing
    ' the VI code.
    
    ' enable idle operation SRQ
    This.ViSession.WriteLine ":STAT:OPER:ENAB 1024"
    This.ViSession.WriteLine "*CLS; *SRE 128"
#End If

    ' clear the service registered bit
    
    This.Device.ReadStatusByte True
    
    ' clear the flag indicating that the external
    ' trigger source was not initiated
    
    This.ExtTrigInitiated = False

End Sub

''' <summary> Sets the instrument to read using external trigger. </summary>
''' <remarks>   Typically, readings will be fetched when the instrument requests service. </remarks>
''' <param name="a_function">        [String] The measurement function. </param>
''' <param name="a_frontInputs ">    [Boolean] True if front inputs are required;
'''                                  otherwise, rear inputs are required. </param>
Public Sub ConfigureExternalTriggerReading(ByVal a_function As String, ByVal a_frontInputs As Boolean)

    ' make sure that the instrument is ready to accept
    ' input from the front panel
    
    If a_frontInputs And Not Me.IsFrontInputs Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ConfigurateExternalTriggerFrontReading", _
            " The instrument Inputs must be set for Front reading."
    ElseIf Not a_frontInputs And Me.IsFrontInputs Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ConfigurateExternalTriggerFrontReading", _
            " The instrument Inputs must be set for rear reading."
    End If
    
    ' Clear the device
    If This.ViSession.UsingGpibLan Then This.ViSession.GpibLan.SelectiveDeviceClear
    This.Device.ClearExecutionState
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"
    
    ' set external triggering
    This.ViSession.WriteLine ":INIT:CONT OFF"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.ViSession.WriteLine ":FUNC '" & a_function & "'"
    
    This.ViSession.WriteLine ":" & a_function & ":RANG:AUTO ON"
    This.ViSession.WriteLine ":" & a_function & ":NPLC 1"
    This.ViSession.WriteLine ":SAMP:COUN 1"
    This.ViSession.WriteLine ":TRIG:COUN 1"
    
    ' set the trigger source to external
    This.TriggerSystem.SourceSetter TriggerSource.External
    ' This.ViSession.WriteLine ":TRIG:SOUR EXT"
    
    ' enable the detection of externally triggered measurement available.
    
    Me.EnableMeasurementAvailableMonitoring
  
End Sub

''' <summary>   Monitors an external trigger taking a reading from the front panel. </summary>
''' <remarks>   Call this method from a timer event. </remarks>
''' <param name="a_channelNumber">   [Integer] The channel number. </param>
''' <param name="a_isDone">          [Boolean] True if done. </param>
Public Sub MonitorExternalTrigger(ByVal a_channelNumber As Integer, ByVal a_isDone As Boolean)

    ' allow other events to occur
    VBA.DoEvents
    
    If This.ExtTrigInitiated Then
    
        ' if awaiting the external trigger (external trigger source
        ' was initiated), check if service is requested
        ' check if we can use this instead:
        
        If 0 <> This.Device.ReadStatusByte(True) Then
      
            ' clear the trigger initiation status to indicate
            ' that it needs to be primed on the next timer event
            This.ExtTrigInitiated = False
        
            ' clear the serial poll bit of this device
            ' the Service request bit might already be cleared at this point
            This.Device.ReadStatusByte True
            
            ' fetch the data
            This.ViSession.WriteLine ":FETCH?"
          
            ' read the data
            Dim p_reading As String: p_reading = This.ViSession.Read()

            ' raise the channel measured event.
            OnChannelMeasured a_channelNumber, p_reading
          
        End If
    
    ElseIf Not a_isDone Then

        ' we need to initiate the next wait for an external trigger.
        
        ' toggle external trigger source.
        
        ' These commands are already set when the instrument is confitured.
        ' This.ViSession.WriteLine ":TRIG:SOUR IMM"
        ' This.ViSession.WriteLine ":TRIG:SOUR EXT"
        ' This.ViSession.WriteLine ":FRES:NPLC 1"
          
        ' clear the serial poll register
        
        This.Device.ReadStatusByte True
        
        ' initiate awaiting for the external trigger
        
        This.ViSession.WriteLine ":INIT; *CLS"
        
        ' and indicate that external triggering was initiated
        
        This.ExtTrigInitiated = True
    
    End If
    
    ' check if the option button is on
    If a_isDone Then
    
        ' if we are done,
        
        ' reset the instrument
          
        ' abort triggering
        
        If This.ViSession.UsingGpibLan Then This.ViSession.GpibLan.SelectiveDeviceClear
        
        ' should we do system:preset here?
        
        This.Device.ClearExecutionState
        
        Me.RestoreContinuousReading
        
        ' these were replaced with restoration of
        ' Continuous reading method above.
        ' This.ViSession.WriteLine ":INIT:CONT OFF"
        ' This.ViSession.WriteLine ":FUNC 'FRES'"
        ' This.ViSession.WriteLine ":TRIG:SOUR IMM"
        ' This.ViSession.WriteLine ":INIT:CONT ON"
        
        ' clear execution state and wait for operation completion.
        
        This.Device.ClearExecutionState
        
        ' and indicate that external triggering was not
        This.ExtTrigInitiated = False
            
    End If

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  System Subsystem
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the front (True) or rear (False) inputs. </summary>
''' <value>   True if front input; otherwise, rear input. </value>
Public Property Get IsFrontInputs() As Boolean

    IsFrontInputs = Me.ScpiSystem.QueryFrontSwitch()

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Route Subsystem
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the options value. </summary>
''' <value>   [String]. </value>
Public Property Get Options() As String
        
    If This.Device Is Nothing Then
        Options = VBA.vbNullString
    Else
        Options = This.Device.Options
    End If

End Property

''' <summary>   Gets the card name. </summary>
''' <returns>   [String]. </returns>
Public Function CardName(ByVal a_slotNumber As Integer) As String
    
    If This.RouteSystem.InstalledCards.Count >= a_slotNumber Then
        CardName = This.RouteSystem.InstalledCard(a_slotNumber).Name
    Else
        CardName = VBA.vbNullString
    End If

End Function

''' <summary>   Gets the card scan list. </summary>
''' <returns>   [String]. </returns>
Public Function FunctionScanList(ByVal a_slotNumber As Integer) As String
        
    If This.RouteSystem.InstalledCards.Count >= a_slotNumber Then
        FunctionScanList = This.RouteSystem.InstalledCard(a_slotNumber).FunctionScanList
    Else
        FunctionScanList = VBA.vbNullString
    End If

End Function

''' <summary>   Gets the options value and sets the cards. </summary>
''' <returns>   [Integer] the number of installed cards. </returns>
Public Function ReadCards() As Integer

    Dim p_7700CardCount As Integer
    p_7700CardCount = This.RouteSystem.Populate7700Cards()
    
    Dim p_installedCardsCount As Integer
    If p_7700CardCount > 0 Then
    
        p_installedCardsCount = This.RouteSystem.PopulateCards(This.Device.QueryOptions)
        
    End If
    ReadCards = p_installedCardsCount
    
End Function

''' <summary>   Sets the complete scan lists for measurements on all cards. </summary>
''' <param name="a_function">    The measurement function. </param>
Public Sub DefineFunctionScanLists(ByVal a_function As String)

    This.RouteSystem.BuildFunctionScanLists a_function
    
    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.RouteSystem.InstalledCards
        This.ViSession.WriteLine p_multiplexerCard.FunctionScanList
    Next

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Trigger Subsystem
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets true if the instrument is configured for taking
'''             immeidate measurments. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CanTakeImmediateMeasurement() As Boolean
    
    CanTakeImmediateMeasurement = Me.Connected And _
        This.TriggerSystem.Source = cc_isr_Tcp_Scpi.TriggerSource.Immediate
                                    
End Property

''' <summary>   Gets true if the instrument is configured for external
'''             triggers. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CanTakeExternalTriggers() As Boolean
    
    CanTakeExternalTriggers = Me.Connected And _
        This.TriggerSystem.Source = cc_isr_Tcp_Scpi.TriggerSource.External
                                    
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' channel Measured
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ChannelMeasured"/> event. </summary>
''' <para name="a_channelNumber">   [integer] The channel number. </param>
''' <para name="a_reading">         [double] The reading. </param>
Private Sub OnChannelMeasured(ByVal a_channelNumber As Integer, ByVal a_reading As String)
    RaiseEvent ChannelMeasured( _
            Factory.NewChannelReadingEventArgs.Initialize(a_channelNumber, _
                                                          a_reading))
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

    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set m_viSession = Nothing
    Set This.RouteSystem = Nothing
    Set This.ScpiSystem = Nothing
    Set This.ViSession = Nothing
    
    If Not This.Device Is Nothing Then
        This.Device.Dispose
    End If
    Set This.Device = Nothing
    
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
    
    Const p_procedureName = "m_viSession_ConnectionChanged"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        Dim p_installedCardsCount As Integer
        p_installedCardsCount = Me.ReadCards()
    
    End If
       
    OnConnectionChanged Me.Connected
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook

    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.ViSession.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_viSession_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
    IConnectable_OnConnectionChanging a_eventArgs
End Sub


