VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_sessionReadTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_socketAddressCell As String = "C2"
Private Const m_socketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "J2"
Private Const m_receivedMessageCell As String = "K2"
Private Const m_elapsedTimeCell As String = "L2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B33"

Private Const m_serialPollCell As String = "E11"
Private Const m_statusByteCell As String = "F11"
Private Const m_standardByteCell As String = "H11"

Private Const m_srqCell As String = "C10"
Private Const m_gpibAddressCell As String = "C12"
Private Const m_readTimeoutCell As String = "C15"
Private Const m_readAfterWriteEnabledCell As String = "C18"

Private Const m_timerIntervalCell As String = "C22"
Private Const m_card1Cell As String = "C23"
Private Const m_card2Cell As String = "C24"
Private Const m_card1ScanListCell As String = "C25"
Private Const m_card2ScanListCell As String = "C26"
Private Const m_inputsCell As String = "C27"
Private Const m_resistanceNoCell As String = "C28"
Private Const m_resistanceCell As String = "C29"
Private Const m_resistanceCountCell As String = "C30"
Private Const m_externalTriggerStatusCell As String = "C31"

Private Const m_resistanceNoTopRow As Integer = 2
Private Const m_resistanceNoCol As String = "M"
Private Const m_resistanceCol As String = "N"
Private Const m_resistanceFormat As String = "0.00"

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1

Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private Type ThisData

    ' true if a pause is requested so that the timer event is not handled
    ' and the timer is paused on the next event.
    PauseRequested As Boolean
    
    ' true when handling the timer event to prevent timer
    ' event handler overrun.
    HandlingTimerEvent As Boolean
    
    ExtTrigInitiated  As Boolean
    
    CurrentResistanceNo As Integer
    
    Stopwatch As Stopwatch
    
    K2700 As K2700
    
    Device As cc_isr_Ieee488.Device
    
    VISession As cc_isr_Ieee488.VISession
    
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
    
    This.Device = Nothing
    This.VISession = Nothing
    
    If Not m_k2700 Is Nothing Then m_k2700.Dispose
    Set m_k2700 = Nothing
    Set This.K2700 = Nothing
    
    
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
    
    Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize(New ErrTracer)
    Set m_k2700 = This.K2700
    Set This.Device = This.K2700.Device
    Set This.VISession = This.Device.VISession
    
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

''' <summary>   Toggle the enable state of the controls. </summary>
''' <param name="a_connected">   [Boolean] True if connected; otherwise, false. </param>
Friend Sub EnableControls(ByVal a_connected As Boolean)

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
    
    Me.EnableReadResistanceButton
    Me.EnableExtTrigToggleButton
    
    If a_connected Then
        Me.EnableGpibLanControls This.VISession.UsingGpibLan
    Else
        Me.EnableGpibLanControls False
    End If

End Sub

''' <summary>   Toggle the enable state of the Read Resistance Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Friend Sub EnableReadResistanceButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeImmediateMeasurement
    End If
    
    Me.ImmediateTriggerOptionButton.Enabled = p_enable
    Me.ReadResistanceButton.Enabled = False

End Sub

''' <summary>   Toggle the enable state of the External Trig Toggle Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Friend Sub EnableExtTrigToggleButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeExternalTriggers
    End If
    
    Me.ExternalTriggerOptionButton.Enabled = p_enable
    Me.ExtTrigToggleButton.Enabled = False

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.CommandsComboBox.Clear
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.IdentityQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompleteCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OptionsQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardEventEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventStatusQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.ServiceRequestEnableCommand, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Core.StringExtensions.StringFormat( _
                                cc_isr_Ieee488.StandardServiceEnableCommand, &H7F, &H7F) & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand & ";" & _
                                cc_isr_Ieee488.OperationCompletedQueryCommand
   
    Me.CommandsComboBox.AddItem "SYST:BEEP OFF"
    Me.CommandsComboBox.AddItem "SYST:CLE"
    Me.CommandsComboBox.AddItem "SYST:ERR?"
    Me.CommandsComboBox.AddItem "SYST:FRSW?"
    Me.CommandsComboBox.AddItem "FUNC 'RES',(@101:120)"
    Me.CommandsComboBox.AddItem "FUNC 'RES',(@201:220)"
    Me.CommandsComboBox.AddItem "FORM:ELEM READ"
    Me.CommandsComboBox.AddItem "STAT:PRES"
    Me.CommandsComboBox.AddItem "STAT:MEAS:ENAB 32"
    Me.CommandsComboBox.AddItem "FETCH?"
    Me.CommandsComboBox.AddItem "READ?"
    
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

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopwatch()
    
    Range(m_elapsedTimeCell).Value = vbNullString
    This.Stopwatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopwatch()
    
    Range(m_elapsedTimeCell).Value = Format(This.Stopwatch.ElapsedMilliseconds, m_elapsedTimeFormat)

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE 488 Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Write line handling an error messages. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Sub WriteLine(ByVal a_message As String)

    Const p_procedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.WriteLine a_message
    
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

''' <summary>   Query a message from the instrument handling any errors. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Function QueryLine(ByVal a_message As String) As String

    Const p_procedureName = "QueryLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    QueryLine = This.VISession.QueryLine(a_message)
    
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

''' <summary>   Read from the instrument handling any errors. </summary>
Friend Function Read() As String

    Const p_procedureName = "Read"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Read = This.VISession.Read()
    
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

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const p_procedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch

    Dim p_host As String: p_host = Me.Host
    Dim p_port As Integer: p_port = Me.port
    Dim p_sessionReadTimeout As Integer: p_sessionReadTimeout = Me.SessionReadTimeout
      
    If Not ConnectToggleButton.Value And This.K2700.Connected Then
        
        This.K2700.CloseConnection
        
    ElseIf ConnectToggleButton.Value And Not This.K2700.Connected Then
    
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

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    
    Const p_procedureName = "CLSButton_Click"
    
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

''' <summary>   Turns on external trigger mode. </summary>
Private Sub ExternalTriggerOptionButton_Click()

    Const p_procedureName = "ExternalTriggerOptionButton_Click"
    
    If Not ExternalTriggerOptionButton.Value Then Exit Sub
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
   
    This.VISession.WriteLine ":TRIG:SOUR IMM"
    
    This.VISession.WriteLine ":TRIG:SOUR EXT"
    
    ' clear the serial poll register
    
    This.VISession.WriteLine "*ESE 0"
    This.VISession.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    This.VISession.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    This.VISession.WriteLine "STAT:MEAS:ENAB 32"
    
    This.Device.QueryServiceRequestStatus
    
    ' initiate awaiting for the external trigger
    
    This.VISession.WriteLine ":INIT; *CLS"
   
    ' disable the immediate reading mode buttons
    
    ReadResistanceButton.Enabled = False
    
   
    ' enable toggling the external trigger monitoring.
    
    Me.ExtTrigToggleButton.Enabled = True
    
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

    Const p_procedureName = "ImmediateTriggerOptionButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' configure measuring a single resistance on each read.
    This.K2700.ConfigureSingleResistanceReading

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.EnableReadResistanceButton
    Me.EnableExtTrigToggleButton

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

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    Const p_procedureName = "ReadStandardEventStatusButton"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Me.StandardByte = This.Device.QueryStandardEventsStatus
    
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

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    
    Const p_procedureName = "RSTButton_Click"
    
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

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const p_procedureName = "SDCButton_Click"
    
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

''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    
    Const p_procedureName = "ReadStatusByteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.VISession.UsingGpibLan Then
    
        Me.RestartStopwatch
        
        Range(m_serialPollCell).Value = This.VISession.GpibLan.SerialPoll
        
        Me.ReadStopwatch
        
    End If
    
    Me.RestartStopwatch
    
    Me.StatusByte = This.Device.QueryServiceRequestStatus
    
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

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    Const p_procedureName = "WriteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.Value
    
    Me.RestartStopwatch
    
    Me.WriteLine p_messageToSend
    
    Me.ReadStopwatch
    
    Me.SentMessage = p_messageToSend
    If This.VISession.UsingGpibLan And Me.AutoStatusReadCheckBox.Value Then
        Range(m_serialPollCell).Value = This.VISession.GpibLan.SerialPoll
    End If
    
    Dim p_isQuery As Boolean: p_isQuery = cc_isr_Core.StringExtensions.EndsWith(p_messageToSend, "?")
    If Not p_isQuery And Me.AutoStatusReadCheckBox.Value Then
        Me.StatusByte = This.Device.QueryServiceRequestStatus
        Me.StandardByte = This.Device.QueryStandardEventsStatus
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

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const p_procedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Dim p_messageReceived As String
    
    Me.RestartStopwatch
    
    p_messageReceived = Me.Read
    
    Me.ReadStopwatch
    
    Me.ReceivedMessage = p_messageReceived
   
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

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const p_procedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReceivedMessage = vbNullString
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.Value
    Dim p_messageReceived As String
    
    Me.RestartStopwatch
    
    p_messageReceived = Me.QueryLine(p_messageToSend)
    
    Me.ReadStopwatch
    
    Me.ReceivedMessage = p_messageReceived
    Me.SentMessage = p_messageToSend
    
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

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get Host() As String
    Host = Range(m_inputHostCell).Value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get port() As Integer
    port = CInt(Range(m_inputPortCell).Value)
End Property

''' <summary>   Get the receive timeout. </summary>
Friend Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = CInt(Range(m_sessionReadTimeoutCell).Value)
End Property

''' <summary>   Gets the received message. </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).Value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let ReceivedMessage(ByVal a_value As String)
    Range(m_receivedMessageCell).Value = a_value
    Range(m_receivedMessageLengthCell).Value = Len(a_value)
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
Friend Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = Range(m_readAfterWriteEnabledCell).Value
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
Friend Property Let ReadAfterWriteEnabled(ByVal a_value As Boolean)
    Range(m_readAfterWriteEnabledCell).Value = a_value
End Property

''' <summary>   Gets the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).Value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal a_value As String)
    Range(m_sentMessageCell).Value = a_value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).Value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal a_value As Integer)
    Range(m_statusByteCell).Value = a_value
End Property

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_standardByteCell).Value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal a_value As Integer)
    Range(m_standardByteCell).Value = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).Value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
Friend Property Get LastMessage() As String
    LastMessage = Range(m_lastMessageCell).Value
End Property

''' <summary>   Sets the Last Message. </summary>
Friend Property Let LastMessage(ByVal a_value As String)
    Range(m_lastMessageCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB LAN
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Toggle the enabled state of the GPIB Lan controls. </summary>
''' <param name="a_usingGpibLanController">   [Boolean] True if using the GPIB-Lan controller. </param>
Friend Sub EnableGpibLanControls(ByVal a_usingGpibLanController As Boolean)
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' GPIB LAN Control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Write a message to the device via the GPIB Lan. </summary>
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Function QueryGpibLan(ByVal a_message As String) As String

    Const p_procedureName = "QueryGpibLan"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    If 0 < This.VISession.GpibLan.SendToController(a_message) Then
        QueryGpibLan = This.VISession.GpibLan.ReceiveFromDevice()
    End If
    
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

''' <summary>   Read from the device using the GPIB LAN. </summary>
Friend Function ReadGpibLan() As String

    Const p_procedureName = "ReadGpibLan"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    ReadGpibLan = This.VISession.GpibLan.ReceiveFromDevice()
    
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
''' <param name="a_message">   [String] to send to the instrument. </param>
Friend Function WriteGpibLan(ByVal a_message As String) As Long

    Const p_procedureName = "WriteLine"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    WriteGpibLan = This.VISession.GpibLan.SendToController(a_message)
    
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

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()

    Const p_procedureName = "GoToLocalButton_Click"
    
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
Private Sub LocalLockoutButton_Click()

    Const p_procedureName = "LocalLockoutButton_Click"
    
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
Private Sub ReadAfterWriteSetButton_Click()

    Const p_procedureName = "ReadAfterWriteSetButton_Click"
    
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
Private Sub ReadAfterWriteGetButton_Click()

    Const p_procedureName = "ReadAfterWriteGetButton_Click"
    
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

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()

    Const p_procedureName = "SerialPollButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_serialPollCell).Value = This.VISession.GpibLan.SerialPoll
    
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

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()

    Const p_procedureName = "SRQButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_srqCell).Value = This.VISession.GpibLan.ServiceRequested
    
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

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GetGpibAddressButton_Click()

    Const p_procedureName = "GetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_gpibAddressCell).Value = This.VISession.GpibLan.GpibAddressGetter

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
Private Sub SetGpibAddressButton_Click()

    Const p_procedureName = "SetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.GpibAddressSetter CInt(Range(m_gpibAddressCell).Value)

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
Private Sub ReadTimeoutGetButton_Click()

    Const p_procedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_readTimeoutCell).Value = This.VISession.GpibLan.ReadTimeoutGetter

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
Private Sub ReadTimeoutSetButton_Click()

    Const p_procedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    This.VISession.GpibLan.ReadTimeoutSetter (CInt(Range(m_readTimeoutCell).Value))
    
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

''' <summary> Get the timer interval. </summary>
Friend Property Get TimerInterval() As Integer
    TimerInterval = CInt(Range(m_timerIntervalCell).Value)
End Property

''' <summary> Get the resistance count. </summary>
Friend Property Get ResistanceCount() As Integer
    ResistanceCount = CInt(Range(m_resistanceCountCell).Value)
End Property

''' <summary>   Indicates that front inputs are required for this mode. </summary>
Friend Property Get FrontReadRequired() As Boolean
    FrontReadRequired = Me.FrontRearToggleButton.Value
End Property

''' <summary>   Gets the external trigger status. </summary>
Friend Property Get ExternalTrigEnabled() As Boolean
    ExternalTrigEnabled = Me.ExtTrigToggleButton.Value
End Property

''' <summary>   Enables or disables external triggering. </summary>
Friend Property Let ExternalTrigEnabled(ByVal a_value As Boolean)
    Me.ExtTrigToggleButton.Value = a_value
End Property

''' <summary>   Gets the auto increment resistance number status. </summary>
Friend Property Get AutoIncrementResistanceNoEnabled() As Boolean
    AutoIncrementResistanceNoEnabled = Me.AutoIncrementToggleButton.Value
End Property

''' <summary>   Gets the single read enabled. </summary>
Friend Property Get SingeReadEnabled() As Boolean
    SingeReadEnabled = Me.SingleAllToggleButton.Value
End Property

''' <summary>   Gets the resistance number. </summary>
Friend Property Get SelectedResistanceNumber() As Integer
    SelectedResistanceNumber = CInt(Range(m_resistanceNoCell).Value)
End Property

''' <summary>   Sets the resistance display. </summary>
''' <param name="resistanceNo">      [single] the resistance number. </param>
''' <param name="resistanceValue">   [double] the resistance. </param>
Private Sub SetResistance(ByVal a_resistanceNo As Integer, ByVal a_resistance As Double)

    Range(m_resistanceNoCell).Value = a_resistanceNo
    Range(m_resistanceCell).Value = a_resistance
    
    Dim p_rowNo As Integer: p_rowNo = m_resistanceNoTopRow + a_resistanceNo - 1
    Range(m_resistanceNoCol & CStr(p_rowNo)).Value = "R" & CStr(a_resistanceNo)
    Range(m_resistanceCol & CStr(p_rowNo)).Value = Format(a_resistance, m_resistanceFormat)
    
End Sub

''' <summary>   Handles the front/read toggle button click event. </summary>
Private Sub FrontRearToggleButton_Click()
    FrontRearToggleButton.Caption = IIf(FrontRearToggleButton.Value, "Front", "Rear")
End Sub

''' <summary>   Handles the auto-increment toggle button click event. </summary>
Private Sub AutoIncrementToggleButton_Click()
    AutoIncrementToggleButton.Caption = IIf(AutoIncrementToggleButton.Value, "Auto", "Manual")
End Sub

''' <summary>   Handles the single/all button click event. </summary>
Private Sub SingleAllToggleButton_Click()
    SingleAllToggleButton.Caption = IIf(SingleAllToggleButton.Value, "Single", "All")
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

    Const p_procedureName = "QueryInputsButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopwatch
    
    Range(m_inputsCell).Value = IIf(This.K2700.IsFrontInputs, "Front", "Rear")
    
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

''' <summary>   Handles the read resistance button click event. Reads a single or all
''' resistances from the front or read panel. </summary>
Private Sub ReadResistanceButton_Click()
    
    Const p_procedureName = "ReadResistanceButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.ReadResistanceButton.Enabled = False
    Me.ImmediateTriggerOptionButton.Enabled = False
    Me.ExternalTriggerOptionButton.Enabled = False
    
    ' reinitialize if device not defined
    If Not This.K2700.Connected Then
        ' This.K2700.InitializeLegacy Me.host, Me.port, Me.ReceiveTimeout
    End If
    
    If Me.FrontReadRequired And Not This.K2700.IsFrontInputs Then
        Me.LastErrorMessage = "Please release the front panel INPUTS button to its FRONT (F) position"
        GoTo exit_Handler
    ElseIf Not Me.FrontReadRequired And This.K2700.IsFrontInputs Then
        Me.LastErrorMessage = "Please depress the front panel INPUTS button to its REAR (R) position"
        GoTo exit_Handler
    End If

    Dim p_resistanceNo As Integer

    If Me.AutoIncrementResistanceNoEnabled Then
        
        ' Clear all cells
        Me.ClearReadings 0

        ' scan all the resistance
        For p_resistanceNo = 1 To Me.ResistanceCount
        
            ' allow events to occur
            DoEvents
            
            ' set the current resistance
            This.CurrentResistanceNo = p_resistanceNo
            
            Me.RestartStopwatch
            If Me.FrontReadRequired Then
                This.K2700.TakeFrontReading p_resistanceNo
            Else
                This.K2700.TakeRearReading p_resistanceNo
            End If
            Me.ReadStopwatch
        
        Next p_resistanceNo

    Else
        p_resistanceNo = Me.SelectedResistanceNumber
        Me.RestartStopwatch
        If Me.FrontReadRequired Then
            This.K2700.TakeFrontReading p_resistanceNo
        Else
            This.K2700.TakeRearReading p_resistanceNo
        End If
        ' Range("C22").value = This.K2700.TakeRearReading(ResistanceNumber)
        Me.ReadStopwatch
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.ReadResistanceButton.Enabled = True
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

''' <summary>   clears the data acquisition readings. </summary>
''' <param name="offsetResistance">    The offset resistance value to use for clearing the readings. </param>
Public Sub ClearReadings(ByVal a_offsetResistance As Single)

    Dim p_resistanceNumber As Integer
    For p_resistanceNumber = 1 To Me.ResistanceCount
        VBA.DoEvents
        SetResistance p_resistanceNumber, a_offsetResistance
    Next p_resistanceNumber
  
End Sub

''' <summary> Sets the instrument and program to get data using external trigger. </summary>
''' <param name="timerInterval"> The polling timer interval. </param>
Private Sub SetExternalTriggerMode(Optional ByVal a_timerInterval As Integer = 100)

    Const p_procedureName = "SetExternalTriggerMode"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' check if doing an external trigger reading with or without auto increment
    If Me.ExternalTrigEnabled Then
  
        ' clear the triggered flag
        This.ExtTrigInitiated = False
        
        ' make sure that the instrument is ready to accept
        ' input from the front panel
     
        ' check if front or read
        If Me.FrontReadRequired And Not This.K2700.IsFrontInputs Then
        
            ' if rear, toggle back the button
            Me.ExternalTrigEnabled = False
            
            ' display the proper message
            Me.LastErrorMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."
              
            ' and get out
            GoTo exit_Handler
              
        ElseIf Not Me.FrontReadRequired And This.K2700.IsFrontInputs Then
        
        
            ' if rear, toggle back the button
            Me.ExternalTrigEnabled = False
            
            ' display the proper message
            Me.LastErrorMessage = "Please depress the front panel INPUTS button to its REAR (R) position and try again."
              
            ' and get out
            GoTo exit_Handler
      
        End If
  
        Range(m_externalTriggerStatusCell) = "External Trigger On"

        ' set the instrument for externally triggered measurements
        This.K2700.SetExternalTriggerResistanceMode
    
        If Me.AutoIncrementResistanceNoEnabled Then
        
            ' set the first resistance for manual scan
            This.CurrentResistanceNo = 1
        
        Else
        
            ' set the single resistance for manual scan
            This.CurrentResistanceNo = Me.SelectedResistanceNumber
        
        End If
       
        ' start the timer
        Me.StartTimer a_timerInterval
        
    Else
  
        ' once external triggering is turned off, the timer is allowed to
        ' to turn itself off
  
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

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub ExtTrigToggleButton_Click()

    Const p_procedureName = "ExtTrigToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Me.ExternalTrigEnabled Then
    
        Me.ImmediateTriggerOptionButton.Enabled = False
        Me.ExternalTriggerOptionButton.Enabled = False
        
        If Me.AutoIncrementResistanceNoEnabled Then
            
            ' Clear all readings
            Me.ClearReadings 0
        
        End If
    
        Range(m_externalTriggerStatusCell) = "External Trigger On"
        
        ' set triggered mode
        SetExternalTriggerMode Me.TimerInterval
    
    Else
    
        ' one external triggering is turned off, the timer
        ' is allowed to turn itself off
        Me.ImmediateTriggerOptionButton.Enabled = True
        Me.ExternalTriggerOptionButton.Enabled = True
    
    End If
 
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.EnableReadResistanceButton
    Me.EnableExtTrigToggleButton

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
''' 2023-06-23 David Hary 1.0.8574 Use the device event handler to update the resistances.
''' </remarks>
Private Sub HandleTimerEvent()

    Const p_procedureName = "HandleTimerEvent"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.PauseRequested Or This.HandlingTimerEvent Then Exit Sub
   
    ' prevent reentry
    This.HandlingTimerEvent = True
    
    Dim p_resistanceNo As Integer
    
    ' check if scan or single
    If AutoIncrementResistanceNoEnabled Then
      
        ' select the current resistance
        p_resistanceNo = This.CurrentResistanceNo
          
    Else
    
        ' get the cell number
        p_resistanceNo = Me.SelectedResistanceNumber
    
    End If
      
    ' allow other events to occur
    DoEvents
    
    Dim p_isDone As Boolean: p_isDone = Not Me.ExternalTrigEnabled
    
    ' stop the timer.
    If p_isDone Then Me.PauseTimer
    
    This.K2700.MonitorExternalTrigger p_resistanceNo, p_isDone
    
    ' if we ended manual mode, enable the rest of the buttons
    If p_isDone Then
    
        ' EnableButtons
        Range(m_externalTriggerStatusCell) = "External Trigger off"
        
        Me.ExtTrigToggleButton.Enabled = True
        Me.ReadResistanceButton.Enabled = True
        
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
    Me.ExtTrigToggleButton.Enabled = True
    Me.ReadResistanceButton.Enabled = True
    
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
Friend Sub StartTimer(Optional ByVal a_interval As Long = 100)
    
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
Friend Sub PauseTimer()
    
    This.PauseRequested = True

End Sub

''' <summary>   Aborts the timer. </summary>
Friend Sub AbortTimer()

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
        SetResistance a_eventArgs.ChannelNumber, -1#
    Else
        SetResistance a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
    End If
    
    If Me.AutoIncrementResistanceNoEnabled Then
    
        ' increment the resistance for the next reading
        This.CurrentResistanceNo = a_eventArgs.ChannelNumber + 1
        If This.CurrentResistanceNo > Me.ResistanceCount Then
            This.CurrentResistanceNo = 1
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
        
        This.K2700.DefineResistanceScanLists
        
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

