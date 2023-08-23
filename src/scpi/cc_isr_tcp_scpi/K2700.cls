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

Private Const MODULE_NAME As String = "K2700"
Private Const Mux40Cards As String = "7702,7708"
Private Const Mux20Cards As String = "7700,7710"

' reference to the VI session
Private WithEvents m_viSession As cc_isr_Ieee488.ViSession
Attribute m_viSession.VB_VarHelpID = -1

Private Type ThisData
    Device As cc_isr_Ieee488.Device
    ViSession As cc_isr_Ieee488.ViSession
    RouteSystem As RouteSystem
    ScpiSystem As ScpiSystem

    Options As String
    HasMux40A As Boolean
    HasMux40B As Boolean
    HasMux20A As Boolean
    HasMux20B As Boolean
    Card1 As String
    Card2 As String
    Card1ScanList As String
    Card2ScanList As String

    ''' <summary> True if the DMM external trigger mode was initiated. </summary>
    ExtTrigInitiated As Boolean
End Type

Private This As ThisData

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
    Set m_viSession = This.Device.ViSession
    
    Set This.ScpiSystem = New ScpiSystem
    This.ScpiSystem.Initialize This.ViSession
    
    Set This.RouteSystem = New RouteSystem
    This.RouteSystem.Initialize This.ViSession


End Sub

''' <summary>   Initializes and returns the <see cref="IErrTracer"/> for the <see cref="Ieee488Session"/>. </summary>
''' <param name="a_tracer">   <see cref="IErrTracer"/> implementation. </param>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function Initialize(ByVal a_tracer As IErrTracer) As K2700
   
    This.Device.Initialize a_tracer
    
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

''' <summary>   Returns a reference to the <see cref="ScpiSystem"/> subsystem. </summary>
''' <value>   An <see cref="ScpiSystem"/>. </returns>
Public Property Get ScpiSystem() As ScpiSystem
    Set ScpiSystem = This.ScpiSystem
End Property

''' <summary>   Initializes the k2700 legacy mode. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, 3000] The session read timeout in milliseconds. </param>
''' <returns>   True if the device was initialized. </returns>
Public Function InitializeLegacy(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 3000) As Boolean

    Me.OpenConnection a_host, a_port

    ' reset clear the device.
    If This.ViSession.UsingGpibLan Then This.ViSession.GpibLan.SelectiveDeviceClear
    This.ViSession.SessionReadTimeout = a_timeoutMs
    This.Device.ResetKnownState
    This.Device.ClearExecutionState
    
    ' read the instrument options and parse the card names and types.
    
    Me.ReadCards
    
    ' define scan list for resistance measurements on all cards.
   
    Me.DefineResistanceScanLists
   
    ' turn on continuous resistance auto range
    Me.ContinuousResistanceAutoRange
    
    InitializeLegacy = True
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  2700 Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Make a resistance measurement
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Performs a single read from the front panel. </summary>
''' <param name="a_resistanceNo">   [Integer] The channel number. </param>
''' <returns>   [String] the reading. </returns>
Public Function TakeFrontReading(Optional a_resistanceNo As Integer = 0) As String

    This.ViSession.WriteLine ":INIT"
    This.ViSession.WriteLine ":READ?"
    
    Dim p_reading As String: p_reading = This.ViSession.Read
    
    OnChannelMeasured a_resistanceNo, p_reading

    TakeFrontReading = p_reading

End Function

''' <summary>   Performs a single scan on the given channel number from the rear panel. </summary>
''' <remarks>
''' Use this method to get a single value from the instrument using the back panel switches
''' </remarks>
''' <param name="a_channelNumber">   The channel number. </param>
''' <returns>    [String] the reading of empty if the channel number is out
''' of range for this device. </returns>
Public Function TakeRearReading(ByVal a_channelNumber As Integer) As String

    ' get the channels to close
    Dim p_resistanceScanList As String
    On Error Resume Next
    p_resistanceScanList = Me.BuildScanList(a_channelNumber)
    On Error GoTo 0
    
    Dim p_reading As String
    p_reading = VBA.vbNullString
    
    If VBA.Len(p_resistanceScanList) > 0 Then
    
        ' get the channel list
        This.ViSession.WriteLine ":ROUT:CLOS " & p_resistanceScanList
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

''' <summary>   Sets the instrument to continuous auto range resistance measurement. </summary>
Public Function ContinuousResistanceAutoRange()

    ' set resistance defaults
    This.ViSession.WriteLine ":RES:RANG:AUTO ON"
    This.ViSession.WriteLine ":RES:NPLC 1"
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"
    
    ' turn on continuous mode
    This.ViSession.WriteLine ":FUNC 'FRES'"
    This.ViSession.WriteLine ":FRES:RANG:AUTO ON"
    This.ViSession.WriteLine ":FRES:NPLC 1"
    This.ViSession.WriteLine ":INIT:CONT On"

End Function

''' <summary>   Configures single resistance reading. </summary>
Public Function ConfigureSingleResistanceReading()

    ' set the device to measure 4 wire resistance
    ' This.ViSession.WriteLine "*RST"
    This.ViSession.WriteLine ":TRIG:SOUR IMM"
    This.ViSession.WriteLine ":INIT:CONT OFF"
    This.ViSession.WriteLine ":SAMP:COUN 1"
    This.ViSession.WriteLine ":TRIG:COUN 1"
    This.ViSession.WriteLine ":FUNC 'RES'"
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  External trigger mode
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary> Sets the instrument to read front inputs resistance using external trigger. </summary>
''' <remarks>   Typically, readings will be fetched when the instrument requests service. </remarks>
Public Sub SetExternalTriggerResistanceMode()

    ' make sure that the instrument is ready to accept
    ' input from the front panel
    
    If Not Me.IsFrontInputs Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SetExternalTriggerResistanceMode", _
            " The instrument Inputs must be set for Front inputs"
    End If
    
    ' Clear the device
    If This.ViSession.UsingGpibLan Then This.ViSession.GpibLan.SelectiveDeviceClear
    This.Device.ClearExecutionState
    
    ' set reading format
    This.ViSession.WriteLine ":FORM:ELEM READ"
    
    ' set external triggering
    This.ViSession.WriteLine ":INIT:CONT OFF"
    This.ViSession.WriteLine ":FUNC 'FRES'"
    This.ViSession.WriteLine ":FRES:RANG:AUTO ON"
    This.ViSession.WriteLine ":FRES:NPLC 1"
    This.ViSession.WriteLine ":SAMP:COUN 1"
    This.ViSession.WriteLine ":TRIG:COUN 1"
    This.ViSession.WriteLine ":TRIG:SOUR EXT"
    
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
    
    ' clear the flag indicating that the external
    ' trigger mode was not initiated
    This.ExtTrigInitiated = False
  
End Sub

''' <summary>   Monitors an external trigger taking a reading from the front panel. </summary>
''' <remarks>   Call this method from a timer event. </remarks>
''' <param name="a_channelNumber">   [Integer] The channel number. </param>
''' <param name="a_isDone">          [Boolean] True if done. </param>
Public Sub MonitorExternalTrigger(ByVal a_channelNumber As Integer, ByVal a_isDone As Boolean)

    ' allow other events to occur
    DoEvents
    
    If This.ExtTrigInitiated Then
    
        ' if awaiting the external trigger (external trigger mode
        ' was initiated), check if service is requested
        ' check if we can use this instead:
        
        If This.Device.QueryServiceRequestStatus <> 0 Then
      
            ' clear the trigger initiation status to indicate
            ' that it needs to be primed on the next timer event
            This.ExtTrigInitiated = False
        
            ' clear the serial poll bit of this device
            This.Device.QueryServiceRequestStatus
            
            ' fetch the data
            This.ViSession.WriteLine ":FETCH?"
          
            ' read the data
            Dim p_reading As String: p_reading = This.ViSession.Read()

            ' raise the channel measured event.
            OnChannelMeasured a_channelNumber, p_reading
          
        End If
    
    ElseIf Not a_isDone Then

        ' we need to initiate the next wait for an external trigger.
        
        ' toggle external trigger mode.
        
        This.ViSession.WriteLine ":TRIG:SOUR IMM"
        This.ViSession.WriteLine ":TRIG:SOUR EXT"
        
        This.ViSession.WriteLine ":FRES:NPLC 1"
          
        ' clear the serial poll register
        
        This.Device.QueryServiceRequestStatus
        
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
        
        ' set the device to measure 4 wire resistance
        
        This.ViSession.WriteLine ":INIT:CONT OFF"
        This.ViSession.WriteLine ":FUNC 'FRES'"
        This.ViSession.WriteLine ":TRIG:SOUR IMM"
    
        ' restore continuous reading
        
        This.ViSession.WriteLine ":INIT:CONT ON"
        
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
''' <value> [String]. </value>
Public Property Get Options() As String
        
    Options = This.Options

End Property

''' <summary>   Gets the top card. </summary>
''' <value> [String]. </value>
Public Property Get Card1() As String
        
    Card1 = This.Card1

End Property

''' <summary>   Sets the top card. </summary>
''' <param name="a_value">   [String] The name of the top card. </param>
''' <value> [String]. </value>
Public Property Let Card1(ByVal a_value As String)
    This.Card1 = a_value
    This.HasMux20A = 0 < InStr(1, Mux20Cards, a_value, vbTextCompare)
    This.HasMux40A = 0 < InStr(1, Mux40Cards, a_value, vbTextCompare)
End Property

''' <summary>   Gets the top card scan list. </summary>
''' <value> [String]. </value>
Public Property Get Card1ScanList() As String
        
    Card1ScanList = This.Card1ScanList

End Property

''' <summary>   Gets the bottom card. </summary>
''' <value> [String]. </value>
Public Property Get Card2() As String
        
    Card2 = This.Card2

End Property

''' <summary>   Sets the bottom card. </summary>
''' <param name="a_value">   [String] The name of the bottom card. </param>
''' <value> [String]. </value>
Public Property Let Card2(ByVal a_value As String)
    This.Card2 = a_value
    This.HasMux40B = 0 < InStr(1, Mux40Cards, a_value, vbTextCompare)
    This.HasMux20B = 0 < InStr(1, Mux20Cards, a_value, vbTextCompare)
End Property

''' <summary>   Gets the bottom card scan list. </summary>
''' <value> [String]. </value>
Public Property Get Card2ScanList() As String
        
    Card2ScanList = This.Card2ScanList

End Property

''' <summary>   Returns true if the instrument as a 40 channel multiplexer
''' such as the 7702 board in the first position. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasMux40A() As Boolean
        
    HasMux40A = This.HasMux40A

End Property

''' <summary>   Returns true if the instrument as a 40 channel multiplexer
''' such as the 7702 board in the second position. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasMux40B() As Boolean

    HasMux40B = This.HasMux40B

End Property

''' <summary>   Returns true if the instrument as a 20 channel multiplexer
''' such as the 7700 or 7710 boards in the first position. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasMux20A() As Boolean

    HasMux20A = This.HasMux20A

End Property

''' <summary>   Returns true if the instrument as a 20 channel multiplexer
''' such as the 7700 or 7710 boards in the second position. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasMux20B() As Boolean

    HasMux20B = This.HasMux20B

End Property

''' <summary>   Returns true if the instrument as any mux cards. </summary>
''' <value>   [Boolean]. </value>
Public Property Get HasMuxCards() As Boolean

    HasMuxCards = Me.HasMux20A Or Me.HasMux20B Or Me.HasMux40A Or Me.HasMux40B

End Property

''' <summary>   Gets the options value and sets the cards. </summary>
Public Function ReadCards()

    This.Options = This.Device.QueryOptions
    
    Me.Card1 = StringExtensions.Pop(This.Options, ",")
    Me.Card2 = StringExtensions.Pop(This.Options, ",")
    
End Function

''' <summary>   Sets the complete scan lists for resistance measurements on all cards. </summary>
Public Function DefineResistanceScanLists()

    If Me.HasMux40A Then
        This.Card1ScanList = ":FUNC 'RES',(@101:140)"
        This.ViSession.WriteLine This.Card1ScanList
    End If
    
    If Me.HasMux40B Then
        This.Card2ScanList = ":FUNC 'RES',(@201:240)"
        This.ViSession.WriteLine This.Card2ScanList
    End If
    
    If Me.HasMux20A Then
        This.Card1ScanList = ":FUNC 'RES',(@101:120)"
        This.ViSession.WriteLine This.Card1ScanList
    End If
    
    If Me.HasMux20B Then
        This.Card2ScanList = ":FUNC 'RES',(@201:220)"
        This.ViSession.WriteLine This.Card2ScanList
    End If

End Function

''' <summary>   Returns the channel list. </summary>
''' <param name="a_resistanceNumber">   [Integer] The channel number. </param>
''' <returns>  The channel list to close. </returns>
Public Function BuildChannelList(ByVal a_resistanceNumber As Integer) As String

    Dim p_routeCommand As String
    
    ' Check if we need to close the bottom or top channel.
    If a_resistanceNumber <= 20 Then
    
        If Me.HasMux40A Then
          
            p_routeCommand = ":ROUT:MULT:CLOS (@144,145)"
        
        ElseIf Me.HasMux20A Then
          
            p_routeCommand = ":ROUT:MULT:CLOS (@124,125)"
        
        End If
    
    ElseIf a_resistanceNumber <= 40 Then
      
        If Me.HasMux40A Then
        
            ' Get channel string
            p_routeCommand = ":ROUT:MULT:CLOS (@144,145)"
        
        ElseIf Me.HasMux20B Then
          
            ' Get channel string
            p_routeCommand = ":ROUT:MULT:CLOS (@224,225)"
        
        End If
      
    ElseIf a_resistanceNumber > 40 Then
    
        If Me.HasMux40B Then
            p_routeCommand = ":ROUT:MULT:CLOS (@244,245)"
        End If
    
    End If
        
    ' set the return value
    BuildChannelList = p_routeCommand

End Function

''' <summary>   Returns the scan list for a specific resistance. </summary>
''' <param name="a_resistanceNo">   [Integer] The channel number. </param>
''' <returns>  The scan list. </returns>
Public Function BuildScanList(ByVal a_resistanceNo As Integer) As String

    Dim p_scanList As String
    Dim p_resistaceChannel As String
    
    ' Check if we need to close the bottom or top channel.
    If a_resistanceNo <= 20 Then
    
        ' Get channel string
        If Me.HasMux40A Or Me.HasMux20A Then
        
            p_resistaceChannel = "1" & Right$("0" & Format$(a_resistanceNo), 2)
            p_scanList = "(@" & p_resistaceChannel & ")"
        
        End If
      
    ElseIf a_resistanceNo <= 40 Then
      
        If Me.HasMux40A Then
        
            ' Get channel string
            p_resistaceChannel = "1" & Right$("0" & Format$(a_resistanceNo), 2)
            p_scanList = "(@" & p_resistaceChannel & ")"
        
        ElseIf Me.HasMux20B Then
          
            ' Get channel string
            p_resistaceChannel = "2" & Right$("0" & Format$(a_resistanceNo - 20), 2)
            
            p_scanList = "(@" & p_resistaceChannel & ")"
        
        End If
      
    ElseIf a_resistanceNo > 40 Then
    
        ' Get channel string
        p_resistaceChannel = "2" & Right$("0" & Format$(a_resistanceNo - 40), 2)
        
        If Me.HasMux40B Then
            p_scanList = "(@" & p_resistaceChannel & ")"
        End If
    
    End If
        
    ' set the return value
    BuildScanList = p_scanList

End Function



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

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_ieee488Session_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub

''' <summary>   Handles the <see cref="cc_isr_Ieee488.Ieee488Session.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_ieee488Session_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

End Sub


