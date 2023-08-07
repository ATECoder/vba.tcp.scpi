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

' reference to the IEEE488 session
Private WithEvents m_ieee488Session As Ieee488Session
Attribute m_ieee488Session.VB_VarHelpID = -1

Private Type ThisData

    ' reference to the SCPI System subsystem class
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

''' <summary>   The event that is raised upon completion of a a resistance measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="ResistanceNumberEventArgs"/> event args. </param>
Public Event ResistanceNumberMeasured(ByVal a_eventArgs As ResistanceNumberEventArgs)

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Initializes the k2700 . </summary>
Private Sub Class_Initialize()

    Set m_ieee488Session = cc_isr_Ieee488.Factory.NewIeee488Session
    m_ieee488Session.Initialize New ErrTracer

End Sub

''' <summary>   Initializes and resturns the <see cref="IErrTracer"/> for the <see cref="Ieee488Session"/>. </summary>
''' <param name="a_tracer">   <see cref="IErrTracer"/> implementation. </param>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function Initialize(ByVal a_tracer As IErrTracer) As K2700
   
    m_ieee488Session.Initialize a_tracer
    
    Set Initialize = Me
   
End Function

''' <summary>   Returns a reference to the <see cref="Ieee488Session"/>. </summary>
''' <value>   An <see cref="Ieee488Session"/>. </returns>
Public Property Get Session() As Ieee488Session
    Set Session = m_ieee488Session
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
    If Me.Session.VISession.UsingGpibLan Then Me.Session.VISession.GpibLan.SelectiveDeviceClear
    Me.Session.VISession.SessionReadTimeout = a_timeoutMs
    Me.Session.ResetKnownState
    Me.Session.ClearExecutionState
    
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

''' <summary>   Gets the options value and sets the cards. </summary>
Public Function ReadCards()

    This.Options = Me.Session.QueryOptions
    
    Me.Card1 = StringExtensions.Pop(This.Options, ",")
    Me.Card2 = StringExtensions.Pop(This.Options, ",")
    
End Function

''' <summary>   Sets the complete scan lists for resistance measurements on all cards. </summary>
Public Function DefineResistanceScanLists()

    If Me.HasMux40A Then
        This.Card1ScanList = ":FUNC 'RES',(@101:140)"
        Me.Session.WriteLine This.Card1ScanList
    End If
    
    If Me.HasMux40B Then
        This.Card2ScanList = ":FUNC 'RES',(@201:240)"
        Me.Session.WriteLine This.Card2ScanList
    End If
    
    If Me.HasMux20A Then
        This.Card1ScanList = ":FUNC 'RES',(@101:120)"
        Me.Session.WriteLine This.Card1ScanList
    End If
    
    If Me.HasMux20B Then
        This.Card2ScanList = ":FUNC 'RES',(@201:220)"
        Me.Session.WriteLine This.Card2ScanList
    End If

End Function

''' <summary>   Sets the instrument to continuous auto range resistance measurement. </summary>
Public Function ContinuousResistanceAutoRange()

    ' set resistance defaults
    Me.Session.WriteLine ":RES:RANG:AUTO ON"
    Me.Session.WriteLine ":RES:NPLC 1"
    
    ' set reading format
    Me.Session.WriteLine ":FORM:ELEM READ"
    
    ' turn on continuous mode
    Me.Session.WriteLine ":FUNC 'FRES'"
    Me.Session.WriteLine ":FRES:RANG:AUTO ON"
    Me.Session.WriteLine ":FRES:NPLC 1"
    Me.Session.WriteLine ":INIT:CONT On"

End Function

''' <summary>   Gets the options value. </summary>
''' <value> A String. </value>
Public Property Get Options() As String
        
    Options = This.Options

End Property

''' <summary>   Gets the top card. </summary>
''' <value> A String. </value>
Public Property Get Card1() As String
        
    Card1 = This.Card1

End Property

''' <summary>   Sets the top card. </summary>
''' <param name="a_value">   [String] The name of the top card. </param>
''' <value> A String. </value>
Public Property Let Card1(ByVal a_value As String)
    This.Card1 = a_value
    This.HasMux20A = 0 < InStr(1, Mux20Cards, a_value, vbTextCompare)
    This.HasMux40A = 0 < InStr(1, Mux40Cards, a_value, vbTextCompare)
End Property

''' <summary>   Gets the top card scan list. </summary>
''' <value> A String. </value>
Public Property Get Card1ScanList() As String
        
    Card1ScanList = This.Card1ScanList

End Property

''' <summary>   Gets the bottom card. </summary>
''' <value> A String. </value>
Public Property Get Card2() As String
        
    Card2 = This.Card2

End Property

''' <summary>   Sets the bottom card. </summary>
''' <param name="a_value">   [String] The name of the bottom card. </param>
''' <value> A String. </value>
Public Property Let Card2(ByVal a_value As String)
    This.Card2 = a_value
    This.HasMux40B = 0 < InStr(1, Mux40Cards, a_value, vbTextCompare)
    This.HasMux20B = 0 < InStr(1, Mux20Cards, a_value, vbTextCompare)
End Property

''' <summary>   Gets the bottom card scan list. </summary>
''' <value> A String. </value>
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

''' <summary>   Returns the channel list. </summary>
''' <param name="a_resistanceNumber">   [Integer] The resistance number. </param>
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
''' <param name="a_resistanceNo">   [Integer] The resistance number. </param>
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

''' <summary>   Gets the front (True) or rear (False) inputs. </summary>
''' <value>   True if front input; otherwise, rear input. </value>
Public Property Get IsFrontInputs() As Boolean

    IsFrontInputs = 0 < CInt(Me.Session.QueryLine("SYST:FRSW?"))

End Property

''' <summary>   Configures single resistance reading. </summary>
Public Function ConfigureSingleResistanceReading()

    ' set the device to measure 4 wire resistance
    ' Me.Session.WriteLine "*RST"
    Me.Session.WriteLine ":TRIG:SOUR IMM"
    Me.Session.WriteLine ":INIT:CONT OFF"
    Me.Session.WriteLine ":SAMP:COUN 1"
    Me.Session.WriteLine ":TRIG:COUN 1"
    Me.Session.WriteLine ":FUNC 'RES'"
    
    ' set reading format
    Me.Session.WriteLine ":FORM:ELEM READ"

End Function

''' <summary>   Performs a single read from the front panel. </summary>
''' <param name="a_resistanceNo">   [Integer] The resistance number. </param>
''' <returns>   [Double] the measured resistance. </returns>
Public Function ReadFrontResistance(Optional a_resistanceNo As Integer = 0) As Double

    Me.Session.WriteLine ":INIT"
    Me.Session.WriteLine ":READ?"
    
    Dim p_reading As String: p_reading = Me.Session.Read
    
    ' convert the measured resistance to double
    Dim p_value As Double: p_value = CDbl(p_reading)
    
    OnResistanceNumberMeasured a_resistanceNo, p_value

    ReadFrontResistance = p_value

End Function

''' <summary>   Performs a single scan on the given resistance from the rear panel. </summary>
''' <remarks>
''' Use this method to get a single value from the instrument using the back panel switches
''' </remarks>
''' <param name="a_resistanceNo">   The resistance number. </param>
''' <returns>    The measured resistance of -1 if the resistance number is out
''' of range for this device. </returns>
Public Function ReadRearResistance(ByVal a_resistanceNo As Integer) As Double

    ' get the channels to close
    Dim p_resistanceScanList As String: p_resistanceScanList = Me.BuildScanList(a_resistanceNo)
    
    If Len(p_resistanceScanList) > 0 Then
    
        ' get the channel list
        Me.Session.WriteLine ":ROUT:CLOS " & p_resistanceScanList
        Me.Session.WriteLine ":INIT"
        Me.Session.WriteLine ":READ?"
        
        Dim p_reading As String: p_reading = Me.Session.Read
        
        ' convert the measured resistance to double
        Dim p_value As Double: p_value = CDbl(p_reading)
        
        OnResistanceNumberMeasured a_resistanceNo, p_value
    
        ReadRearResistance = p_value
    
    Else
        ReadRearResistance = -1
    
    End If

End Function

''' <summary> Sets the instrument to read front inputs resistance using external trigger. </summary>
''' <remarks>   Typically, readings will be fetched when the instrument requests service. </remarks>
Public Sub SetExternalTriggerResistanceMode()

    ' make sure that the instrument is ready to accept
    ' input from the front panel
    
    If Not Me.IsFrontInputs Then
        cc_isr_Core.UserDefinedErrors.RaiseError UserDefinedErrors.InvalidOperationError, _
            "cc_isr_TCP_VI.K2700.SetExternalTriggerResistanceMode", _
            " The instrument Inputs must be set for Front inputs"
    End If
    
    ' Clear the device
    If Me.Session.VISession.UsingGpibLan Then Me.Session.VISession.GpibLan.SelectiveDeviceClear
    Me.Session.ClearExecutionState
    
    ' set reading format
    Me.Session.WriteLine ":FORM:ELEM READ"
    
    ' set external triggering
    Me.Session.WriteLine ":INIT:CONT OFF"
    Me.Session.WriteLine ":FUNC 'FRES'"
    Me.Session.WriteLine ":FRES:RANG:AUTO ON"
    Me.Session.WriteLine ":FRES:NPLC 1"
    Me.Session.WriteLine ":SAMP:COUN 1"
    Me.Session.WriteLine ":TRIG:COUN 1"
    Me.Session.WriteLine ":TRIG:SOUR EXT"
    
    ' setup the status register so that we get a RQS
    ' on measurement available
    
    Me.Session.WriteLine "*ESE 0"
    Me.Session.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    
    Me.Session.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    
    Me.Session.WriteLine "STAT:MEAS:ENAB 32"
    
#If False Then
    ' legacy code used the operation register going to idle to indicate that
    ' a measurment is available. This did not work when testing
    ' the VI code.
    
    ' enable idle operation SRQ
    Me.Session.WriteLine ":STAT:OPER:ENAB 1024"
    Me.Session.WriteLine "*CLS; *SRE 128"
#End If
    
    ' clear the flag indicating that the external
    ' trigger mode was not initiated
    This.ExtTrigInitiated = False
  
End Sub

''' <summary>   Monitors an external trigger reading resistances from the front panel. </summary>
''' <remarks>   Call this method from a timer event. </remarks>
''' <param name="a_resistanceNo">   [Integer] The resistance number. </param>
''' <param name="a_isDone">   [Boolean] True if done. </param>
Public Sub MonitorExternalTrigger(ByVal a_resistanceNo As Integer, ByVal a_isDone As Boolean)

    ' allow other events to occur
    DoEvents
    
    If This.ExtTrigInitiated Then
    
        ' if awaiting the external trigger (external trigger mode
        ' was initiated), check if service is requested
        ' check if we can use this instead:
        
        If Me.Session.QueryServiceRequestStatus <> 0 Then
      
            ' clear the trigger initiation status to indicate
            ' that it needs to be primed on the next timer event
            This.ExtTrigInitiated = False
        
            ' clear the serial poll bit of this device
            Me.Session.QueryServiceRequestStatus
            
            ' fetch the data
            Me.Session.WriteLine ":FETCH?"
          
            ' read the data
            Dim p_resistance As Double: p_resistance = CDbl(Me.Session.Read)

            ' raise the resistance measured event.
            OnResistanceNumberMeasured a_resistanceNo, p_resistance
          
        End If
    
    ElseIf Not a_isDone Then

        ' we need to initiate the next wait for an external trigger.
        
        ' toggle external trigger mode.
        
        Me.Session.WriteLine ":TRIG:SOUR IMM"
        Me.Session.WriteLine ":TRIG:SOUR EXT"
        
        Me.Session.WriteLine ":FRES:NPLC 1"
          
        ' clear the serial poll register
        
        Me.Session.QueryServiceRequestStatus
        
        ' initiate awaiting for the external trigger
        
        Me.Session.WriteLine ":INIT; *CLS"
        
        ' and indicate that external triggering was initiated
        This.ExtTrigInitiated = True
    
    End If
    
    ' check if the option button is on
    If a_isDone Then
    
        ' if we are done,
        
        ' reset the instrument
          
        ' abort triggering
        
        If Me.Session.VISession.UsingGpibLan Then Me.Session.VISession.GpibLan.SelectiveDeviceClear
        
        ' should we do system:preset here?
        
        Me.Session.ClearExecutionState
        
        ' set the device to measure 4 wire resistance
        
        Me.Session.WriteLine ":INIT:CONT OFF"
        Me.Session.WriteLine ":FUNC 'FRES'"
        Me.Session.WriteLine ":TRIG:SOUR IMM"
    
        ' restore continuous reading
        
        Me.Session.WriteLine ":INIT:CONT ON"
        
        ' clear execution state and wait for operation completion.
        
        Me.Session.ClearExecutionState
        
        ' and indicate that external triggering was not
        This.ExtTrigInitiated = False
            
    End If

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Resistance Number Measured
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ResistanceNumberMeasured"/> event. </summary>
''' <para name="a_resistanceNo">   [integer] The resistance number. </param>
''' <para name="a_resistance">     [double] The resistance. </param>
Private Sub OnResistanceNumberMeasured(ByVal a_resistanceNo As Integer, ByVal a_resistance As Double)
    RaiseEvent ResistanceNumberMeasured( _
            cc_isr_TCP_VI.Factory.NewResistanceNumberEventArgs.Initialize(a_resistanceNo, a_resistance))
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
    Set Connectable = m_ieee488Session
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
    
    Set m_ieee488Session = Nothing
    
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
    
    If a_isConnected Then
    
        Set This.ScpiSystem = New ScpiSystem
        This.ScpiSystem.Initialize Me.Session
        
    Else
    
        Set This.ScpiSystem = Nothing
    
    End If
    
    IConnectable_OnConnectionChanged cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize(a_isConnected)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanged"/> event after the
''' connection was established and ended. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionchangedEventArgs"/>]. </param>
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


