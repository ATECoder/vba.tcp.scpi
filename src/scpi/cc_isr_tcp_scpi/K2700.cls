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

' references the event handler of the Socket
Private WithEvents m_socket As IPv4StreamSocket
Attribute m_socket.VB_VarHelpID = -1

Private Type this_

    ' session settings
    Session As cc_isr_Ieee488.TcpSession
    GpibLanControllerPort As Long
    ReadAfterWriteDelay As Long
    Termination As String
    
    ' device settings
    Device As cc_isr_Ieee488.Device
    FormatSystem As FormatSystem
    RouteSystem As RouteSystem
    ScpiSystem As ScpiSystem
    SenseSystem As K2700SenseSystem
    TriggerSystem As TriggerSystem
    ExtTrigInitiated As Boolean
    
End Type

Private This As this_

''' <summary>   The event that is raised upon completion of a DUT measurement. </summary>
''' <param name="a_eventArgs">   An <see cref="DutReadingEventArgs"/> event args. </param>
Public Event DutMeasured(ByVal a_eventArgs As DutReadingEventArgs)

''' <summary>   Initializes the k2700 . </summary>
Private Sub Class_Initialize()

    ' predefined values
    This.GpibLanControllerPort = 1234
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    
    Set This.Device = cc_isr_Ieee488.Factory.NewDevice()
    Set This.Session = This.Device.Session
    
    ' construct the subsystems
    Set This.FormatSystem = New FormatSystem
    Set This.RouteSystem = New RouteSystem
    Set This.ScpiSystem = New ScpiSystem
    Set This.SenseSystem = New K2700SenseSystem
    Set This.TriggerSystem = New TriggerSystem
    
End Sub

''' <summary>   Initializes and returns the <see cref="Ieee488Session"/>. </summary>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function Initialize() As K2700
   
    This.Device.ReadAfterWriteDelay = This.ReadAfterWriteDelay
    This.Device.GpibLanControllerPort = This.GpibLanControllerPort
    This.Device.Termination = This.Termination
   
    ' this also initializes the session.
    This.Device.Initialize
    
    ' Set the event handler after initializing the device, which initializes the Session.
    ' This ensures that the connection events are handled in order: Session, Device, 2700.
    Set m_socket = This.Session.Socket
    
    ' initialize the subsystems.
    This.FormatSystem.Initialize This.Session
    
    This.RouteSystem.Initialize This.Session
    
    This.ScpiSystem.Initialize This.Session
    
    Dim p_senseSystem As SenseSystem
    Set p_senseSystem = New SenseSystem
    p_senseSystem.Initialize This.Session
    This.SenseSystem.Initialize p_senseSystem
    
    This.TriggerSystem.Initialize This.Session
    
    Set Initialize = Me
   
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Session Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.TcpSession"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.TcpSession"/>. </returns>
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
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    This.GpibLanControllerPort = a_value
    If Not This.Device Is Nothing Then _
        This.Device.GpibLanControllerPort = a_value
End Property

''' <summary>   Gets the Read After Write Delay in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = This.ReadAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Read After Write Delay Delay in milliseconds. </param>
Public Property Let ReadAfterWriteDelay(ByVal a_value As Integer)
    This.ReadAfterWriteDelay = a_value
    If Not This.Device Is Nothing Then _
        This.Device.ReadAfterWriteDelay = a_value
End Property

''' <summary>   Gets the termination to append to output messages. </summary>
''' <value>   [String]. </value>
Public Property Get Termination() As String
    Termination = This.Termination
End Property

''' <summary>   Sets the termination to append to output messages. </summary>
''' <param name="a_value">   [String] the termination. </param>
Public Property Let Termination(ByVal a_value As String)
    This.Termination = a_value
    If Not This.Device Is Nothing Then _
        This.Device.Termination = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device Settings
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +


''' <summary>   Returns a reference to the <see cref="cc_isr_Ieee488.Device"/>. </summary>
''' <value>   An <see cref="cc_isr_Ieee488.Device"/>. </returns>
Public Property Get Device() As cc_isr_Ieee488.Device
    Set Device = This.Device
End Property

''' <summary>   Returns a reference to the <see cref="FormatSystem"/> subsystem. </summary>
''' <value>   An <see cref="FormatSystem"/>. </returns>
Public Property Get FormatSystem() As FormatSystem
    Set FormatSystem = This.FormatSystem
End Property

''' <summary>   Returns a reference to the <see cref="RouteSystem"/> subsystem. </summary>
''' <value>   An <see cref="RouteSystem"/>. </returns>
Public Property Get RouteSystem() As RouteSystem
    Set RouteSystem = This.RouteSystem
End Property

''' <summary>   Returns a reference to the <see cref="ScpiSystem"/> subsystem. </summary>
''' <value>   An <see cref="ScpiSystem"/>. </returns>
Public Property Get ScpiSystem() As ScpiSystem
    Set ScpiSystem = This.ScpiSystem
End Property

''' <summary>   Returns a reference to the <see cref="K2700SenseSystem"/> subsystem. </summary>
''' <value>   [<see cref="K2700SenseSystem"/>]. </returns>
Public Property Get SenseSystem() As K2700SenseSystem
    Set SenseSystem = This.SenseSystem
End Property

''' <summary>   Returns a reference to the <see cref="TriggerSystem"/> subsystem. </summary>
''' <value>   [<see cref="TriggerSystem"/>]. </returns>
Public Property Get TriggerSystem() As TriggerSystem
    Set TriggerSystem = This.TriggerSystem
End Property

''' <summary>   Initializes the k2700 legacy mode. </summary>
''' <param name="a_address">    [String] An IPv4 dotted-decimal host address and post. </param>
''' <param name="a_details">    [Out, String] Details the failure information if any. </param>
''' <param name="a_function">   [Optional, String, 'RES'] The measurement function. </param>
''' <param name="a_timeout">    [Optional, Long, 3000] The session read timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the device was initialized. </returns>
Public Function InitializeLegacy(ByVal a_address As String, _
        ByRef a_details As String, _
        Optional ByVal a_function As String = "RES", _
        Optional ByVal a_timeout As Long = 3000) As Boolean

    Const p_procedureName = "InitializeLegacy"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_success = This.Session.Socket.TryOpenConnection(a_address, a_timeout, p_details)
    
    If p_success Then
        
        Dim p_tries As Integer
        Dim p_repeatCount As Integer: p_repeatCount = 3
        
        p_success = This.Device.TryResetClears(p_details, p_tries, p_repeatCount)
        If p_tries > 1 Then Debug.Print "Reset clear try count= "; p_tries
    
        If Not p_success Then _
            p_details = "Failed resetting the IEEE488 device to its known state; " & p_details
        
    End If
    
    If p_success Then
        
        ' reset clear the device.
        ' this is issued after *RST because SDC fails to synchronize if the instrument is in
        ' it trigger mode (on the 2700).
        p_success = This.Session.TrySelectiveDeviceClear(p_details, True, 500)
    
    End If
    
    If p_success Then
    
        p_success = This.Device.TryClearExecutionState(a_details)
        
        If Not p_success Then _
            p_details = "Failed clearing the IEEE488 device execution state; " & p_details
    
    End If
    
    If p_success Then
    
        ' read the instrument options and parse the card names and types.
        
        Me.RouteSystem.Read7700Cards This.Device.QueryOptions
        
        ' define function scan list for measurements on all cards.
        
        Me.RouteSystem.DefineFunctionScanLists a_function
        
        ' turn on continuous auto range
        p_success = Me.TryConfigureContinuousFunctionAutoRange(a_function, p_details)
        
        If Not p_success Then _
            a_details = "Failed configuring continuous mode; " & p_details
        
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed configuring continous trigger mode; " & p_details

    InitializeLegacy = p_success
    
End Function

''' <summary>   Closes the connection and releases the <see cref="Ieee488Session">Session</see>. </summary>
Public Sub Dispose()
    
    If Not Me.Connectable Is Nothing Then
        If Me.Connectable.Connected Then
            Me.Connectable.CloseConnection
        End If
    End If

    If Not Me.Connectable Is Nothing Then
        Me.Connectable.Dispose
    End If
    
    Set m_socket = Nothing
    Set This.FormatSystem = Nothing
    Set This.RouteSystem = Nothing
    Set This.ScpiSystem = Nothing
    Set This.SenseSystem = Nothing
    Set This.TriggerSystem = Nothing
    Set This.Session = Nothing
    
    If Not This.Device Is Nothing Then
        This.Device.Dispose
    End If
    Set This.Device = Nothing
    
End Sub

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
''' <param name="a_reading">     [Out, String] The reading. </param>
''' <param name="a_details">     [Out, String] Details the failure information if any. </param>
''' <param name="a_dutNumber">   [Optional, Integer, 0] The DUT number. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryTakeFrontReading(ByRef a_reading As String, ByRef a_details As String, _
    Optional a_dutNumber As Integer = 0) As Boolean

    Const p_procedureName = "TryTakeFrontReading"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    a_reading = VBA.vbNullString

    ' trigger and fetch a reading
    
    Dim p_command As String: p_command = ":INIT;*WAI;:READ?"
    p_length = This.Session.TryQueryLine(p_command, a_reading, p_details)
    p_success = 0 < p_length
    If p_success Then
        OnDutMeasured a_dutNumber, a_reading
    Else
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " failed taking a front reading using the command '" & p_command & "; " & p_details
    End If
    
    TryTakeFrontReading = p_success

End Function

''' <summary>   Performs a single scan on the given device-level channel number from the rear panel. </summary>
''' <remarks>
''' Use this method to get a single value from the instrument using the back panel switches
''' </remarks>
''' <param name="a_deviceChannelNumber">   The device-level channel number. </param>
''' <param name="a_reading">               [Out, String] The reading or empty if
''                                         the channel number is out of range for this device. </param>
''' <param name="a_details">               [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryTakeRearReading(ByVal a_deviceChannelNumber As Integer, _
    ByRef a_reading As String, ByRef a_details As String) As Boolean

    Const p_procedureName = "TryTakeRearReading"
    
    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_command As String
    
    ' get the channel close command
    Dim p_routeCloseCommand As String
    On Error Resume Next
    
    p_routeCloseCommand = This.RouteSystem.GetRouteCloseCommand(a_deviceChannelNumber)
    
    On Error GoTo 0
    
    a_reading = VBA.vbNullString
    
    If VBA.Len(p_routeCloseCommand) > 0 Then
    
        ' close the rear channel
        p_command = p_routeCloseCommand & ";*WAI;*OPC?"
        p_length = This.Session.TryQueryLine(p_command, a_reading, p_details)
        p_success = 0 < p_length
        If Not p_success Then _
            p_details = "Failed closing channel #" & VBA.CStr(a_deviceChannelNumber) & _
                " using the command '" & p_command & "; " & p_details
        
        If p_success Then
        
            ' trigger and fetch a reading
            
            p_command = ":INIT;*WAI;:READ?"
            p_length = This.Session.TryQueryLine(p_command, a_reading, p_details)
            p_success = 0 < p_length
            If p_success Then
                OnDutMeasured a_deviceChannelNumber, a_reading
            Else
                p_details = "Failed taking a rear reading using the command '" & p_command & "; " & p_details
            End If
        
        End If
        
    End If

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryTakeRearReading = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Continuous and Immediate modes
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Configure continuous trigger mode for the specified function
''' abdpower line cycles. </summary>
''' <param name="a_function">          [String] The measurement function. </param>
''' <param name="a_details">           [Out, String] Details the failure information if any. </param>
''' <param name="a_powerLineCycles">   [Optional, Double, 1] number of power line cycle for
'''                                    the reading aperture. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureContinuousFunctionAutoRange(ByVal a_function As String, _
    ByRef a_details As String, Optional a_powerLineCycles As Double = 1#) As Boolean

    Const p_procedureName = "TryConfigureContinuousFunctionAutoRange"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString

    ' set function defaults
    ' This.Session.WriteLine ":RES:RANG:AUTO ON"
    ' This.Session.WriteLine ":RES:NPLC 1"
    
    ' set reading format
    This.Session.WriteLine ":FORM:ELEM READ"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.Session.WriteLine ":FUNC '" & a_function & "'"
    
    ' turn on continuous mode
    This.Session.WriteLine ":" & a_function & ":RANG:AUTO ON"
    This.Session.WriteLine ":" & a_function & ":NPLC " & Format(a_powerLineCycles, "#.###")
    
    Dim p_command As String: p_command = ":INIT:CONT ON;*CLS"
    p_success = This.Device.TryApplyCommandAwaitOperationCompleted(p_command, p_details)
    If Not p_success Then _
        p_details = "Failed initiating continuous function mode; " & p_details
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryConfigureContinuousFunctionAutoRange = p_success

End Function

''' <summary>   Restore the continuous trigger mode. </summary>
''' <remarks>   This is typically called to restore continuous mode without
''' changing the function mode. </remarks>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryRestoreContinuousReading(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryRestoreContinuousReading"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    This.Session.WriteLine ":INIT:CONT OFF"
    
    ' set the trigger source to immediate.
    This.TriggerSystem.SourceSetter TriggerSourceOption.Immediate

    Dim p_command As String: p_command = ":INIT:CONT ON"
    p_success = This.Device.TryApplyCommandAwaitOperationCompleted(p_command, p_details)
    If Not p_success Then _
        p_details = "Failed restoring continuous reading mode; " & p_details
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryRestoreContinuousReading = p_success

End Function

''' <summary>   Configures single function reading. </summary>
''' <param name="a_function">   [String] The measurement function. </param>
''' <param name="a_details">    [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureSingleFunctionReading(ByVal a_function As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryRestoreContinuousReading"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' reset is not applied on each configuration because it takes too long to complete.
    ' This.Session.WriteLine "*RST"
    
    ' set the trigger source to immediate.
    ' This.Session.WriteLine ":TRIG:SOUR IMM"
    This.TriggerSystem.SourceSetter cc_isr_Tcp_Scpi.TriggerSourceOption.Immediate
    This.Session.WriteLine ":INIT:CONT OFF"
    This.Session.WriteLine ":SAMP:COUN 1"
    This.Session.WriteLine ":TRIG:COUN 1"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.Session.WriteLine ":FUNC '" & a_function & "'"
    
    ' set reading format
    
    Dim p_command As String: p_command = ":FORM:ELEM READ"
    p_success = This.Device.TryApplyCommandAwaitOperationCompleted(p_command, p_details)
    If Not p_success Then _
        p_details = "Failed configuring single function reading mode; " & p_details
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    
    TryConfigureSingleFunctionReading = p_success

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  External trigger source
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the condition indicating that the external trigger was initiated
''' at which state the trigger monitoring loop is looking to receive a triggered
''' reading. </summary>
Public Property Get ExtTrigInitiated() As Boolean
    ExtTrigInitiated = This.ExtTrigInitiated
End Property

''' <summary>   Sets the instrument to detect the measurement available bit
'''             of the status register upon receiving an external trigger. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryEnableMeasurementAvailableMonitoring(ByRef a_details As String) As Boolean

    Const p_procedureName = "TryEnableMeasurementAvailableMonitoring"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_statusByte As Integer
    
    ' setup the status register so that we get a RQS
    ' on measurement available
    
    This.Session.WriteLine "*ESE 0"
    This.Session.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    This.Session.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    Dim p_command As String: p_command = "STAT:MEAS:ENAB 32"
    p_success = This.Device.TryApplyCommandAwaitOperationCompleted(p_command, p_details)
    If Not p_success Then _
        p_details = "Failed enabling measurement available monitoring; " & p_details
    
#If False Then
    ' legacy code used the operation register going to idle to indicate that
    ' a measurement is available. This did not work when testing
    ' the VI code.
    
    ' enable idle operation SRQ
    This.Session.WriteLine ":STAT:OPER:ENAB 1024"
    This.Session.WriteLine "*CLS; *SRE 128"
#End If

    ' clear the service registered bit
    
    p_statusByte = This.Device.TryReadStatusByte(p_details, True)
    
    ' clear the flag indicating that the external
    ' trigger source was not initiated
    
    This.ExtTrigInitiated = False

    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details

    TryEnableMeasurementAvailableMonitoring = p_success
    
End Function

''' <summary> Sets the instrument to read using external trigger. </summary>
''' <remarks>   Typically, readings will be fetched when the instrument requests service. </remarks>
''' <param name="a_function">   [String] The measurement function. </param>
''' <param name="a_details">    [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureExternalTriggerReading(ByVal a_function As String, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryConfigureExternalTriggerReading"

    Dim p_success As Boolean: p_success = True
    a_details = VBA.vbNullString

    ' set reading format
    This.Session.WriteLine ":FORM:ELEM READ"
    
    ' set external triggering
    This.Session.WriteLine ":INIT:CONT OFF"
    
    ' set the sense function
    This.SenseSystem.SenseSystem.SenseFunctionSetter a_function
    ' This.Session.WriteLine ":FUNC '" & a_function & "'"
    
    This.Session.WriteLine ":" & a_function & ":RANG:AUTO ON"
    This.Session.WriteLine ":" & a_function & ":NPLC 1"
    This.Session.WriteLine ":SAMP:COUN 1"
    This.Session.WriteLine ":TRIG:COUN 1"
    
    ' set the trigger source to external
    This.TriggerSystem.SourceSetter TriggerSourceOption.External
    ' This.Session.WriteLine ":TRIG:SOUR EXT"
    
    ' enable the detection of externally triggered measurement available.
    
    p_success = Me.TryEnableMeasurementAvailableMonitoring(a_details)
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & a_details
    
    TryConfigureExternalTriggerReading = p_success
    
End Function

''' <summary>   Monitors an external trigger taking a reading from the front panel. </summary>
''' <remarks>   Call this method from a timer event.
''' Note that each trigger causes the instrument to place a measurement to be fetched. Debouncing
''' is required to ensure that only a single reading is taken upon each trigger.
''' For example, here is a sequence of triggered readings where a single button press causes
''' three triggers. In this case, the reading was processed if the status byte was non-zero.
''' The new code will not process these readings. Notice, that the 5t-th reading is empty.
''' <code>
''' Awaiting triggers...
''' Status byte:  65 ; SRQ: True; Cleared status byte:  1
''' Reading: '+1.00118553E+02'.
'''  1 : 100.118553
''' Status byte:  1 ; SRQ: False; Cleared status byte:  1
''' Reading: '+1.00118332E+02'.
'''  2 : 100.118332
''' Status byte:  1 ; SRQ: False; Cleared status byte:  1
''' Reading: '+1.00118118E+02'.
''' 3 : 100.118118
''' Status byte:  1 ; SRQ: False; Cleared status byte:  1
''' Reading: '+1.00117813E+02'.
'''  4 : 100.117813
''' Status byte:  1 ; SRQ: False; Cleared status byte:  0
''' Reading: ''.
'''  5 : -1
''' </code>
''' Test 10 TestTriggerPollingShouldRead passed. Elapsed time: 20856.2 ms.
''' </remarks>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number. </param>
''' <param name="a_isDone">                [Boolean] True if done. </param>
''' <param name="a_details">               [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryMonitorExternalTrigger(ByVal a_deviceChannelNumber As Integer, ByVal a_isDone As Boolean, _
    ByRef a_details As String) As Boolean

    Const p_procedureName = "TryMonitorExternalTrigger"

    Dim p_success As Boolean: p_success = True: a_details = VBA.vbNullString
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_length As Long
    Dim p_statusByte As Integer
    Dim p_clearedStatusByte As Integer
    Dim p_clearedStatusByte2 As Integer
    Dim p_isServiceRequested As Boolean
    
    Dim p_command As String

    ' allow other events to occur
    VBA.DoEvents
    
    If a_isDone Then
    
        ' if we are done,
        
        ' and indicate that external triggering was not initiated
        This.ExtTrigInitiated = False
    
        ' reset the instrument
          
        ' abort triggering
        
        ' reset known state.
        p_success = This.Device.TryResetClear(p_details)
        
        If Not p_success Then
            p_details = "Failed resetting and clearing the instrument when ending external trigger monitoring; " & p_details
        End If
        
        If p_success Then
        
            p_success = Me.TryRestoreContinuousReading(p_details)
            ' these were replaced with restoration of
            ' Continuous reading method above.
            ' This.Session.WriteLine ":INIT:CONT OFF"
            ' This.Session.WriteLine ":FUNC 'FRES'"
            ' This.Session.WriteLine ":TRIG:SOUR IMM"
            ' This.Session.WriteLine ":INIT:CONT ON"
            
        Else
            p_details = "Failed clearing execution state when ending external trigger monitoring; " & p_details
        End If
        
        If Not p_success Then _
            p_details = "Failed restoring continuous reading when ending external trigger monitoring; " & p_details
        
        ' clear execution state and wait for operation completion.
        ' This.Device.ClearExecutionState
        
    ElseIf This.ExtTrigInitiated Then
    
        ' if awaiting the external trigger (external trigger source
        ' was initiated), check if service is requested
        ' check if we can use this instead:
        p_statusByte = This.Device.TryReadStatusByte(p_details)
        
        If 0 > p_statusByte Then
        
            ' ignore error reading status byte
            
        ElseIf 0 <> p_statusByte Then
      
            ' clear the serial poll bit of this device
            ' the Service request bit might already be cleared at this point
            p_clearedStatusByte = This.Device.TryReadStatusByte(p_details)
            
            ' this solved the multiple reading issues to a large extent.
            ' now each new triggered reading has the RQS bit enabled.
            p_length = This.Session.TryWriteLine("*CLS", p_details)
            
            p_isServiceRequested = cc_isr_Ieee488.ServiceRequestFlags.RequestingService = _
                (cc_isr_Ieee488.ServiceRequestFlags.RequestingService And p_statusByte)
            
            Debug.Print "Status byte: "; p_statusByte; "; SRQ: "; p_isServiceRequested; "; Cleared status byte: "; p_clearedStatusByte
            
            ' fetch and read the data
            
            If p_isServiceRequested Then
            
                ' clear the trigger initiation status to indicate
                ' that it needs to be primed on the next timer event
                This.ExtTrigInitiated = False
            
                p_command = ":FETCH?"
                Dim p_reading As String
                p_length = This.Session.TryQueryLine(p_command, p_reading, p_details)
                p_success = 0 <= p_length
                If p_success Then
                
                    Debug.Print "Reading: '"; p_reading; "'."
                
                    ' raise the channel measured event.
                    If 0 < VBA.Len(p_reading) Then _
                        OnDutMeasured a_deviceChannelNumber, p_reading
                    
                Else
                
                    a_details = "Failed fetching a reading; " & p_details
                
                End If
            
            End If
          
                
        End If
    
    Else

        ' we need to initiate the next wait for an external trigger.
        
        ' toggle external trigger source.
        
        ' These commands are already set when the instrument is confitured.
        'This.Session.WriteLine ":TRIG:SOUR IMM"
        'This.Session.WriteLine ":TRIG:SOUR EXT"
        'This.Session.WriteLine ":FRES:NPLC 1"
          
        ' clear the serial poll register
        
        p_clearedStatusByte2 = This.Device.TryReadStatusByte(p_details)
        
        ' initiate awaiting for the external trigger. This cannot include an opc query.
        
        p_command = ":INIT;*CLS"
        p_length = This.Session.TryWriteLine(p_command, p_details)
        p_success = 0 < p_length
        If p_success Then
            ' and indicate that external triggering was initiated
            This.ExtTrigInitiated = True
        Else
            p_details = "Failed preparing for the next external trigger; " & p_details
        End If
    
    End If
    
    If Not p_success Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; " & p_details
    

    TryMonitorExternalTrigger = p_success
    
End Function

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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Trigger Subsystem
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets true if the instrument is configured for taking
'''             immediate measurments. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CanTakeImmediateMeasurement() As Boolean
    
    CanTakeImmediateMeasurement = Me.Connected And _
        This.TriggerSystem.Source = cc_isr_Tcp_Scpi.TriggerSourceOption.Immediate
                                    
End Property

''' <summary>   Gets true if the instrument is configured for external
'''             triggers. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CanTakeExternalTriggers() As Boolean
    
    CanTakeExternalTriggers = Me.Connected And _
        This.TriggerSystem.Source = cc_isr_Tcp_Scpi.TriggerSourceOption.External
                                    
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' channel Measured
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="DutMeasured"/> event. </summary>
''' <remarks>   Lacking an internal definition for unit testing, this method
''' is made public so as the unit test the propagation of measured values and
''' device-level channel numbers. </remarks>
''' <para name="a_dutNumber">   [integer] The DUT number equaling the device-level
'''                             channel number. </param>
''' <para name="a_reading">     [double] The reading. </param>
Public Sub OnDutMeasured(ByVal a_dutNumber As Integer, ByVal a_reading As String)
    RaiseEvent DutMeasured( _
            Factory.NewDutReadingEventArgs.Initialize(a_dutNumber, _
                                                          a_reading))
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the <see cref="cc_isr_Winsock.IConnectable"/>
''' <see cref="Session"/see>.<see cref="TcpSession.Socket"/see>. </summary>
''' <value>   [<see cref="cc_isr_Winsock.IConnectable"/>]. </value>
Public Property Get Connectable() As cc_isr_Winsock.IConnectable
    If This.Session Is Nothing Then
        Set Connectable = Nothing
    ElseIf This.Session.Connectable Is Nothing Then
        Set Connectable = Nothing
    Else
        Set Connectable = This.Session.Connectable
    End If
End Property

''' <summary>   Returns the connection state of the <see cref="Session">Tcp Session</see>. </summary>
''' <value>   [Boolean] True if the <see cref="Session">Tcp Session</see> is connected; otherwise, False. </value>
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Restore Initial State
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Configures the 2700 Multimeter scanner initial state. </summary>
''' <remarks>   Call this method upon connecting and restoring Initial State. </remarks>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryConfigureInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"

    Dim p_success As Boolean: p_success = True
    a_details = VBA.vbNullString
    
    If p_success Then
    
        Me.RouteSystem.Read7700Cards This.Device.QueryOptions
    
    End If
    
    TryConfigureInitialState = p_success

End Function

''' <summary>   Checks if the 2700 needs to restore to ites Initial State. </summary>
''' <param name="a_details">   [Out, String] Details the information of why the instrument
'''                            should restore state. </param>
''' <returns>   [Boolean]. </returns>
Public Function ShouldRestoreInitialState(ByRef a_details As String) As Boolean

    a_details = VBA.vbNullString
    If This.Device Is Nothing Then
        a_details = "The IEEE 488 device instance is nothing."
        ShouldRestoreInitialState = True
    ElseIf This.Device.Session Is Nothing Then
        a_details = "The Tcp Session instance is nothing."
        ShouldRestoreInitialState = True
    Else
        ShouldRestoreInitialState = This.Session.ShouldRestoreInitialState(a_details)
    End If

End Function

''' <summary>   Restores the instrument to its initial state in case the instrument power
''' was toggled or a recovery from error is needed. </summary>
''' <remarks>   This method attempts to restore the instrument to a initial state in case it is
''' powered off while the program is active as follows:
''' - The user defined errors are cleared;
''' - The class is initialized if the session is nothing;
''' - The device is connected if it is not connected;
''' - The GPIB-Lan device is restored to its initial state;
''' - Execution state is cleared.
''' - the cards are read.
''' </remarks>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean] true if restored. </returns>
Public Function TryRestoreInitialState(ByRef a_details As String) As Boolean

    Const p_procedureName As String = "TryConfigureInitialState"
    
    Dim p_success As Boolean: p_success = True
    a_details = VBA.vbNullString
    
    If p_success And This.Device Is Nothing Then
    
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            " unable to restore state because the IEEE 488 device was disposed. Please restart the program."
        p_success = False
    
    End If
    
    If p_success And This.Session Is Nothing Then
    
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
            "; unable to restore state because the Tcp Session was disposed. Please restart the program."
        p_success = False
    
    End If
    
    If p_success Then
    
        If This.Session Is Nothing Then
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; unable to restore state because the Tcp Session was disposed. Please restart the program."
            p_success = False
        End If
    
    End If
    
    If p_success And m_socket Is Nothing Then
    
        Me.Initialize
        
    End If
    
    If p_success And Not (Me.Connected Or This.Session.Socket.CanConnect) Then
        
        ' if not connected and not 'can connect' it means that the
        ' socket was disposed and needs to be rebuilt.
        Me.Initialize
            
        p_success = This.Session.Socket.CanConnect
        
        If Not p_success Then
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                    " failed initializing the socket; the socket id is invalid ('" & _
                VBA.CStr(cc_isr_Winsock.ws32_INVALID_SOCKET) & "')."
        End If
    
    End If
    
    If p_success Then
        p_success = This.Device.TryRestoreInitialState(a_details)
        If Not p_success Then _
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & a_details
    End If
    
    If p_success Then
    
        ' configure the initial state
        p_success = Me.TryConfigureInitialState(a_details)
    
        If Not p_success Then _
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                "; " & a_details
    
    End If

    TryRestoreInitialState = p_success

End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  IPV4 Socket Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4StreamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
    
    Const p_procedureName = "m_socket_ConnectionChanged"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        Dim p_details As String
         
        ' configure the initial state
        If Not Me.TryConfigureInitialState(p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Failed configuring the initial state; " & p_details
    
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

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4StreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub


