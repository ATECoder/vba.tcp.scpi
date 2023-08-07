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
Private Const m_lastErrorMessageCell As String = "B4"

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

' true if a pause is requested so that the timer event is not handled
' and the timer is paused on the next event.
Private m_pauseRequested As Boolean

' true when handling the timer event to prevent timer
' event handler overrun.
Private m_handlingTimerEvent As Boolean

Private m_extTrigInitiated  As Boolean

Private m_currentResistanceNo As Integer

Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private m_stopWatch As StopWatch

''' <summary>   Gets the singleton instance of the 2700 virtual instrument. </summary>
''' <value>   The singleton instance of the 2700 virtual instrument. </value>
Public Property Get K2700() As cc_isr_TCP_VI.K2700
    If m_k2700 Is Nothing Then
        Set m_k2700 = cc_isr_TCP_VI.Factory.NewK2700
    End If
    Set K2700 = m_k2700
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.Ieee488Session
    Set Session = Me.K2700.Session
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </value>
Public Property Get StopWatch() As cc_isr_Core.StopWatch
    If m_stopWatch Is Nothing Then
        Set m_stopWatch = cc_isr_Core.Factory.NewStopWatch
    End If
    Set StopWatch = m_stopWatch
End Property

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set m_stopWatch = Nothing
    
    If Not m_eventTimer Is Nothing Then
      m_eventTimer.StopTimer
      Set m_eventTimer = Nothing
    End If
    
    If Not m_k2700 Is Nothing Then m_k2700.Dispose
    Set m_k2700 = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Toggle the enable state of the controls. </summary>
''' <param name="a_connected">   [Boolean] True if connected; otehrwise, false. </param>
Friend Sub EnableControls(ByVal a_connected As Boolean)

    Me.ConnectToggleButton.value = a_connected
    
    ' report the connection state
    Range(m_connectedCell).value = a_connected
        
    Me.RSTButton.Enabled = a_connected
    Me.CLSButton.Enabled = a_connected
    Me.QueryButton.Enabled = a_connected
    Me.ReadButton.Enabled = a_connected
    Me.ReadStatusByteButton.Enabled = a_connected
    Me.ReadStandardEventStatusButton.Enabled = a_connected
    Me.WriteButton.Enabled = a_connected
    
    Me.InitExtTrigButton.Enabled = a_connected
    Me.ReadCardsButton.Enabled = a_connected
    Me.SetScansButton.Enabled = a_connected
    Me.QueryInputsButton.Enabled = a_connected
    
    Me.EnableReadResistanceButton
    Me.EnableExtTrigToggleButton
    
    If a_connected Then
        Me.EnableGpibLanControls Me.Session.VISession.UsingGpibLan
    Else
        Me.EnableGpibLanControls False
    End If

End Sub

''' <summary>   Toggle the enable state of the Read Resistane Button based on the status
''' of the K2700 and the availablity of MUX cards. </summary>
Friend Sub EnableReadResistanceButton()
    
    Dim p_enable As Boolean
    p_enable = Not K2700 Is Nothing
    If p_enable Then
        p_enable = Me.K2700.Connected
    End If
    If p_enable Then
        p_enable = Me.K2700.HasMuxCards
    End If
    
    Me.ReadResistanceButton.Enabled = p_enable

End Sub

''' <summary>   Toggle the enable state of the External Trig Toggle Button based on the status
''' of the K2700 and the availablity of MUX cards. </summary>
Friend Sub EnableExtTrigToggleButton()
    
    Dim p_enable As Boolean
    p_enable = Not K2700 Is Nothing
    If p_enable Then
        p_enable = Me.K2700.Connected
    End If
    If p_enable Then
        p_enable = Me.K2700.HasMuxCards
    End If
    
    Me.ExtTrigToggleButton.Enabled = p_enable

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
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopWatch()
    
    Range(m_elapsedTimeCell).value = vbNullString
    Me.StopWatch.Restart

End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopWatch()
    
    Range(m_elapsedTimeCell).value = Format(Me.StopWatch.ElapsedMilliseconds, m_elapsedTimeFormat)

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
    
    Me.RestartStopWatch
    
    Me.Session.WriteLine a_message
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
    
    Me.RestartStopWatch
    
    QueryLine = Me.Session.QueryLine(a_message)
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Read from the instrument handling any errors. </summary>
Friend Function Read() As String

    Const p_procedureName = "Read"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Read = Me.Session.Read()
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    Const p_procedureName = "ConnectToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch

    Dim p_host As String: p_host = Me.host
    Dim p_port As Integer: p_port = Me.port
    Dim p_sessionReadTimeout As Integer: p_sessionReadTimeout = Me.SessionReadTimeout
      
    If Not ConnectToggleButton.value And Me.K2700.Connected Then
        
        Me.K2700.CloseConnection
        
    ElseIf ConnectToggleButton.value And Not Me.K2700.Connected Then
    
        Me.K2700.OpenConnection p_host, p_port
        
        Me.K2700.Session.VISession.SessionReadTimeout = p_sessionReadTimeout
        
    End If
    
    Me.ReadStopWatch
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    
    Const p_procedureName = "CLSButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ClearExecutionState
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Initialize External Trigger Button click event. </summary>
''' <remarks>   Sets the instrument to turn on the RQS bit on external trigger. </remarks>
Private Sub InitExtTrigButton_Click()


    Const p_procedureName = "ReadStandardEventStatusButton"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
   
    Me.Session.WriteLine ":TRIG:SOUR IMM"
    
    Me.Session.WriteLine ":TRIG:SOUR EXT"
    
    ' clear the serial poll register
    
    Me.Session.WriteLine "*ESE 0"
    Me.Session.WriteLine "*SRE 63"
    
    ' clear the measurement event enable register
    Me.Session.WriteLine "STAT:PRES"
    
    ' enable the reading available bit of the measurement event register.
    Me.Session.WriteLine "STAT:MEAS:ENAB 32"
    
    Me.Session.QueryServiceRequestStatus
    
    ' initiate awaiting for the external trigger
    
    Me.Session.WriteLine ":INIT; *CLS"
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    Const p_procedureName = "ReadStandardEventStatusButton"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.StandardByte = Me.Session.QueryStandardEventsStatus
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    
    Const p_procedureName = "RSTButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.ResetKnownState
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    Const p_procedureName = "SDCButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.SelectiveDeviceClear
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    
    Const p_procedureName = "ReadStatusByteButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Me.Session.VISession.UsingGpibLan Then
    
        Me.RestartStopWatch
        
        Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
        
        Me.ReadStopWatch
        
    End If
    
    Me.RestartStopWatch
    
    Me.StatusByte = Me.Session.QueryServiceRequestStatus
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.value
    
    Me.RestartStopWatch
    
    Me.WriteLine p_messageToSend
    
    Me.ReadStopWatch
    
    Me.SentMessage = p_messageToSend
    If Me.Session.VISession.UsingGpibLan And Me.AutoStatusReadCheckBox.value Then
        Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
    End If
    
    Dim p_isQuery As Boolean: p_isQuery = cc_isr_Core.StringExtensions.EndsWith(p_messageToSend, "?")
    If Not p_isQuery And Me.AutoStatusReadCheckBox.value Then
        Me.StatusByte = Me.Session.QueryServiceRequestStatus
        Me.StandardByte = Me.Session.QueryStandardEventsStatus
    End If
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    Const p_procedureName = "ReadButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.receivedMessage = vbNullString
    Dim p_messageReceived As String
    
    Me.RestartStopWatch
    
    p_messageReceived = Me.Read
    
    Me.ReadStopWatch
    
    Me.receivedMessage = p_messageReceived
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    Const p_procedureName = "QueryButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.receivedMessage = vbNullString
    Me.SentMessage = vbNullString
    Dim p_messageToSend As String: p_messageToSend = Me.CommandsComboBox.value
    Dim p_messageReceived As String
    
    Me.RestartStopWatch
    
    p_messageReceived = Me.QueryLine(p_messageToSend)
    
    Me.ReadStopWatch
    
    Me.receivedMessage = p_messageReceived
    Me.SentMessage = p_messageToSend
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
Friend Property Get host() As String
    host = Range(m_inputHostCell).value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get port() As Integer
    port = CInt(Range(m_inputPortCell).value)
End Property

''' <summary>   Get the receive timeout. </summary>
Friend Property Get SessionReadTimeout() As Integer
    SessionReadTimeout = CInt(Range(m_sessionReadTimeoutCell).value)
End Property

''' <summary>   Gets the received message. </summary>
Friend Property Get receivedMessage() As String
    receivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let receivedMessage(ByVal a_value As String)
    Range(m_receivedMessageCell).value = a_value
    Range(m_receivedMessageLengthCell).value = Len(a_value)
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
Friend Property Get ReadAfterWriteEnabled() As Boolean
    ReadAfterWriteEnabled = Range(m_readAfterWriteEnabledCell).value
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
Friend Property Let ReadAfterWriteEnabled(ByVal a_value As Boolean)
    Range(m_readAfterWriteEnabledCell).value = a_value
End Property

''' <summary>   Gets the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal a_value As String)
    Range(m_sentMessageCell).value = a_value
End Property

''' <summary>   Gets the status byte. </summary>
Friend Property Get StatusByte() As Integer
    StatusByte = Range(m_statusByteCell).value
End Property

''' <summary>   Sets the status byte. </summary>
Friend Property Let StatusByte(ByVal a_value As Integer)
    Range(m_statusByteCell).value = a_value
End Property

''' <summary>   Gets the Standard byte. </summary>
Friend Property Get StandardByte() As Integer
    StandardByte = Range(m_standardByteCell).value
End Property

''' <summary>   Sets the Standard byte. </summary>
Friend Property Let StandardByte(ByVal a_value As Integer)
    Range(m_standardByteCell).value = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).value = a_value
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

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()

    Const p_procedureName = "GoToLocalButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.GoToLocal
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()

    Const p_procedureName = "LocalLockoutButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.LocalLockout
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()

    Const p_procedureName = "ReadAfterWriteSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.ReadAfterWriteEnabledSetter Me.ReadAfterWriteEnabled
    
    Me.ReadStopWatch
    
    Me.Session.VISession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the read after write get button click event. </summary>
Private Sub ReadAfterWriteGetButton_Click()

    Const p_procedureName = "ReadAfterWriteGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.ReadAfterWriteEnabled = Me.Session.VISession.GpibLan.ReadAfterWriteEnabledGetter
    
    Me.ReadStopWatch
    
    Me.Session.VISession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()

    Const p_procedureName = "SerialPollButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_serialPollCell).value = Me.Session.VISession.GpibLan.SerialPoll
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()

    Const p_procedureName = "SRQButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_srqCell).value = Me.Session.VISession.GpibLan.ServiceRequested
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GetGpibAddressButton_Click()

    Const p_procedureName = "GetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_gpibAddressCell).value = Me.Session.VISession.GpibLan.GpibAddressGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Private Sub SetGpibAddressButton_Click()

    Const p_procedureName = "SetGpibAddressButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.GpibAddresssetter CInt(Range(m_gpibAddressCell).value)

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()

    Const p_procedureName = "ReadTimeoutGetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_readTimeoutCell).value = Me.Session.VISession.GpibLan.ReadTimeoutGetter

    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()

    Const p_procedureName = "ReadTimeoutSetButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.Session.VISession.GpibLan.ReadTimeoutSetter (CInt(Range(m_readTimeoutCell).value))
    
    Me.ReadStopWatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
Friend Property Get TimerInerval() As Integer
    TimerInerval = CInt(Range(m_timerIntervalCell).value)
End Property

''' <summary> Get the resistance count. </summary>
Friend Property Get ResistanceCount() As Integer
    ResistanceCount = CInt(Range(m_resistanceCountCell).value)
End Property

''' <summary>   Indicates that front inputs are required for this mode. </summary>
Friend Property Get FrontReadRequired() As Boolean
    FrontReadRequired = Me.FrontRearToggleButton.value
End Property

''' <summary>   Gets the external trigger status. </summary>
Friend Property Get ExternalTrigEnabled() As Boolean
    ExternalTrigEnabled = Me.ExtTrigToggleButton.value
End Property

''' <summary>   Enables or disables external triggering. </summary>
Friend Property Let ExternalTrigEnabled(ByVal a_value As Boolean)
    Me.ExtTrigToggleButton.value = a_value
End Property

''' <summary>   Gets the auto increment resistance number status. </summary>
Friend Property Get AutoIncrementResistanceNoEnabled() As Boolean
    AutoIncrementResistanceNoEnabled = Me.AutoIncrementToggleButton.value
End Property

''' <summary>   Gets the single read enabled. </summary>
Friend Property Get SingeReadEnabled() As Boolean
    SingeReadEnabled = Me.SingleAllToggleButton.value
End Property

''' <summary>   Gets the resistance number. </summary>
Friend Property Get SelectedResistanceNumber() As Integer
    SelectedResistanceNumber = CInt(Range(m_resistanceNoCell).value)
End Property

''' <summary>   Sets the resistance display. </summary>
''' <param name="resistanceNo">      [single] the resistance number. </param>
''' <param name="resistanceValue">   [double] the resistance. </param>
Private Sub SetResistance(ByVal a_resistanceNo As Integer, ByVal a_resistance As Double)

    Range(m_resistanceNoCell).value = a_resistanceNo
    Range(m_resistanceCell).value = a_resistance
    
    Dim p_rowNo As Integer: p_rowNo = m_resistanceNoTopRow + a_resistanceNo - 1
    Range(m_resistanceNoCol & CStr(p_rowNo)).value = "R" & CStr(a_resistanceNo)
    Range(m_resistanceCol & CStr(p_rowNo)).value = Format(a_resistance, m_resistanceFormat)
    
End Sub

''' <summary>   Handles the front/read toggle button click event. </summary>
Private Sub FrontRearToggleButton_Click()
    FrontRearToggleButton.Caption = IIf(FrontRearToggleButton.value, "Front", "Rear")
End Sub

''' <summary>   Handles the auto-increment toggle button click event. </summary>
Private Sub AutoIncrementToggleButton_Click()
    AutoIncrementToggleButton.Caption = IIf(AutoIncrementToggleButton.value, "Auto", "Manual")
End Sub

''' <summary>   Handles the single/all button click event. </summary>
Private Sub SingleAllToggleButton_Click()
    SingleAllToggleButton.Caption = IIf(SingleAllToggleButton.value, "Single", "All")
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

    Const p_procedureName = "QueryInputsButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Range(m_inputsCell).value = IIf(Me.K2700.IsFrontInputs, "Front", "Rear")
    
    Me.ReadStopWatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the read cards button click event. </summary>
Private Sub ReadCardsButton_Click()
    
    Const p_procedureName = "ReadCardsButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.K2700.ReadCards
    
    Me.ReadStopWatch
    
    Range(m_card1Cell).value = Me.K2700.Card1
    Range(m_card2Cell).value = Me.K2700.Card2

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.EnableReadResistanceButton
    Me.EnableExtTrigToggleButton

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
    Me.ExtTrigToggleButton.Enabled = False
    
    ' reinitialize if device not defined
    If Not Me.K2700.Connected Then
        ' Me.K2700.InitializeLegacy Me.host, Me.port, Me.ReceiveTimeout
    End If
    
    If Me.FrontReadRequired And Not Me.K2700.IsFrontInputs Then
        Me.LastErrorMessage = "Please release the front panel INPUTS button to its FRONT (F) position"
        GoTo exit_Handler
    ElseIf Not Me.FrontReadRequired And Me.K2700.IsFrontInputs Then
        Me.LastErrorMessage = "Please depress the front panel INPUTS button to its REAR (R) position"
        GoTo exit_Handler
    End If

    ' configure measuring a single resistance on each read.
    Me.K2700.ConfigureSingleResistanceReading

    Dim p_resistanceNo As Integer

    If Me.AutoIncrementResistanceNoEnabled Then
        
        ' Clear all cells
        Me.ClearReadings 0

        ' scan all the resistance
        For p_resistanceNo = 1 To Me.ResistanceCount
        
            ' allow events to occur
            DoEvents
            
            ' set the current resistance
            m_currentResistanceNo = p_resistanceNo
            
            Me.RestartStopWatch
            If Me.FrontReadRequired Then
                Me.K2700.ReadFrontResistance p_resistanceNo
            Else
                Me.K2700.ReadRearResistance p_resistanceNo
            End If
            Me.ReadStopWatch
        
        Next p_resistanceNo

    Else
        p_resistanceNo = Me.SelectedResistanceNumber
        Me.RestartStopWatch
        If Me.FrontReadRequired Then
            Me.K2700.ReadFrontResistance p_resistanceNo
        Else
            Me.K2700.ReadRearResistance p_resistanceNo
        End If
        ' Range("C22").value = Me.K2700.ReadRearResistance(ResistanceNumber)
        Me.ReadStopWatch
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.ReadResistanceButton.Enabled = True
    Me.ExtTrigToggleButton.Enabled = True
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Set Scans button click event. Sets and displays the scan list
''' for the cards. </summary>
Private Sub SetScansButton_Click()

    Const p_procedureName = "SetScansButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.RestartStopWatch
    
    Me.K2700.DefineResistanceScanLists
    
    Me.ReadStopWatch
    
    Range(m_card1ScanListCell).value = Me.K2700.Card1ScanList
    Range(m_card2ScanListCell).value = Me.K2700.Card2ScanList
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   clears the data acquisition readings. </summary>
''' <param name="offsetResistance">    The offset resistance value to use for clearing the readings. </param>
Public Sub ClearReadings(ByVal a_offsetResistance As Single)

    Dim p_resistanceNumber As Integer
    For p_resistanceNumber = 1 To Me.ResistanceCount
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
        m_extTrigInitiated = False
        
        ' make sure that the instrument is ready to accept
        ' input from the front panel
     
        ' check if front or read
        If Me.FrontReadRequired And Not Me.K2700.IsFrontInputs Then
        
            ' if rear, toggle back the button
            Me.ExternalTrigEnabled = False
            
            ' display the proper message
            Me.LastErrorMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."
              
            ' and get out
            GoTo exit_Handler
              
        ElseIf Not Me.FrontReadRequired And Me.K2700.IsFrontInputs Then
        
        
            ' if rear, toggle back the button
            Me.ExternalTrigEnabled = False
            
            ' display the proper message
            Me.LastErrorMessage = "Please depres the front panel INPUTS button to its REAR (R) position and try again."
              
            ' and get out
            GoTo exit_Handler
      
        End If
  
        Range(m_externalTriggerStatusCell) = "External Trigger On"

        ' set the instrument for externally triggered measurements
        Me.K2700.SetExternalTriggerResistanceMode
    
        If Me.AutoIncrementResistanceNoEnabled Then
        
            ' set the first resistance for manual scan
            m_currentResistanceNo = 1
        
        Else
        
            ' set the single resistance for manual scan
            m_currentResistanceNo = Me.SelectedResistanceNumber
        
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
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
    
        Me.ExtTrigToggleButton.Enabled = True
        Me.ReadResistanceButton.Enabled = False
        
        If Me.AutoIncrementResistanceNoEnabled Then
            
            ' Clear all readings
            Me.ClearReadings 0
        
        End If
    
        Range(m_externalTriggerStatusCell) = "External Trigger On"
        
        ' set triggered mode
        SetExternalTriggerMode Me.TimerInerval
    
    Else
    
      ' one external triggering is turned off, the timer
      ' is allowed to turn itself off
    
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
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
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
    
    If m_pauseRequested Or m_handlingTimerEvent Then Exit Sub
   
    ' prevent reentry
    m_handlingTimerEvent = True
    
    Dim p_resistanceNo As Integer
    
    ' check if scan or single
    If AutoIncrementResistanceNoEnabled Then
      
        ' select the current resistance
        p_resistanceNo = m_currentResistanceNo
          
    Else
    
        ' get the cell number
        p_resistanceNo = Me.SelectedResistanceNumber
    
    End If
      
    ' allow other events to occur
    DoEvents
    
    Dim p_isDone As Boolean: p_isDone = Not Me.ExternalTrigEnabled
    
    ' stop the timer.
    If p_isDone Then Me.PauseTimer
    
    Me.K2700.MonitorExternalTrigger p_resistanceNo, p_isDone
    
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
    m_handlingTimerEvent = False

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
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Starts the event timer. </summary>
''' <param name="a_interval">   [Long] Specifies the timer interval in milliseconds. </param>
Friend Sub StartTimer(Optional ByVal a_interval As Long = 100)
    
    m_handlingTimerEvent = False
    m_pauseRequested = False
    If m_eventTimer Is Nothing Then
        Set m_eventTimer = cc_isr_Core.Factory.NewEventTimer.Initialize(a_interval)
        m_eventTimer.StartTimer
    Else
        m_eventTimer.ResumeTimer
    End If

End Sub

''' <summary>   Requests pausing the timer. </summary>
Friend Sub PauseTimer()
    
    m_pauseRequested = True

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
    If m_pauseRequested Then
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

''' <summary>   Handles the K2700 <see cref="cc_isr_TCP_VI.ResistanceNumberMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_TCP_VI.ResistanceNumberEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ResistanceNumberMeasured(ByVal a_eventArgs As cc_isr_TCP_VI.ResistanceNumberEventArgs)

    Const p_procedureName = "m_session_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    SetResistance a_eventArgs.ResistanceNumber, a_eventArgs.Resistance
    
    If Me.AutoIncrementResistanceNoEnabled Then
    
        ' increment the resistance for the next reading
        m_currentResistanceNo = a_eventArgs.ResistanceNumber + 1
        If m_currentResistanceNo > Me.ResistanceCount Then
            m_currentResistanceNo = 1
        End If
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_TCP_VI.K2700.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName = "m_session_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.EnableControls a_eventArgs.Connected

    If a_eventArgs.Connected Then
    
        ' report the socket number
        Range(m_socketAddressCell).value = Me.Session.VISession.socketAddress
        Range(m_socketIdCell).value = Me.Session.SocketId
        
        Me.Session.VISession.GpibLan.SelectiveDeviceClear
        Me.Session.ResetKnownState
        Me.Session.ClearExecutionState
        
        If Me.Session.VISession.UsingGpibLan Then
            Me.ReadAfterWriteEnabled = Me.Session.VISession.GpibLan.ReadAfterWriteEnabledGetter
        End If
    
    Else
    
        Range(m_socketAddressCell).value = vbNullString
        Range(m_socketIdCell).value = vbNullString
        
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    Me.LastErrorMessage = ThisWorkbook.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the <see cref="cc_isr_TCP_VI.K2700.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event arguments. </param>
Private Sub m_k2700_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub





