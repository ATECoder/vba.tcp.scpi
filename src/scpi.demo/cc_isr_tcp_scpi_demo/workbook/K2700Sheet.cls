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

Private Const m_readingFormatCell As String = "C21"
Private Const m_timerIntervalCell As String = "C22"
Private Const m_card1Cell As String = "C23"
Private Const m_card2Cell As String = "C24"
Private Const m_card1ScanListCell As String = "C25"
Private Const m_card2ScanListCell As String = "C26"
Private Const m_inputsCell As String = "C27"
Private Const m_channelNumberCell As String = "C28"
Private Const m_functionNameCell As String = "B29"
Private Const m_functionValueCell As String = "C29"
Private Const m_functionUnitCell As String = "D29"
Private Const m_channelCountCell As String = "C30"
Private Const m_externalTriggerStatusCell As String = "C31"

Private Const m_measurementTopRow As Integer = 2
Private Const m_measurementNoCol As String = "M"
Private Const m_measurementValueCol As String = "N"

' defines the standard event timer.
Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1

Private WithEvents m_k2700 As K2700
Attribute m_k2700.VB_VarHelpID = -1

Private Type ThisData
    
    Initialized As Boolean

    ' true if a pause is requested so that the timer event is not handled
    ' and the timer is paused on the next event.
    PauseRequested As Boolean
    
    ' true when handling the timer event to prevent timer
    ' event handler overrun.
    HandlingTimerEvent As Boolean
    
    CurrentChannelNumber As Integer
    
    Stopwatch As Stopwatch
    
    K2700 As K2700
    
    Device As cc_isr_Ieee488.Device
    
    ViSession As cc_isr_Ieee488.ViSession
    
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
    This.ViSession = Nothing
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Me.EnableMeasureButton
    Me.EnableMonitorExtTrigToggleButton
    
    If a_connected Then
        Me.EnableGpibLanControls This.ViSession.UsingGpibLan
    Else
        Me.EnableGpibLanControls False
    End If

End Sub

''' <summary>   Toggle the enable state of the Measure Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Friend Sub EnableMeasureButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeImmediateMeasurement
    End If
    
    Me.ImmediateTriggerOptionButton.Enabled = p_enable
    Me.MeasureButton.Enabled = False

End Sub

''' <summary>   Toggle the enable state of the External Trig Toggle Button based on the status
''' of the K2700 and the availability of MUX cards. </summary>
Friend Sub EnableMonitorExtTrigToggleButton()
    
    Dim p_enable As Boolean
    p_enable = Not This.K2700 Is Nothing
    If p_enable Then
        p_enable = This.K2700.CanTakeExternalTriggers
    End If
    
    Me.ExternalTriggerOptionButton.Enabled = p_enable
    Me.MonitorExtTrigToggleButton.Enabled = False

End Sub

''' <summary>   Populates the Function combo box. </summary>
Friend Sub PopulateFunctionComboBox()

    Const p_procedureName = "PopulateFunctionComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.FunctionComboBox.Clear
    Me.FunctionComboBox.AddItem "CONT"
    Me.FunctionComboBox.AddItem "CURR:AC"
    Me.FunctionComboBox.AddItem "CURR:DC"
    Me.FunctionComboBox.AddItem "FREQ"
    Me.FunctionComboBox.AddItem "FRES"
    Me.FunctionComboBox.AddItem "PER"
    Me.FunctionComboBox.AddItem "RES"
    Me.FunctionComboBox.AddItem "TEMP"
    Me.FunctionComboBox.AddItem "VOLT:AC"
    Me.FunctionComboBox.AddItem "VOLT:DC"

    Me.FunctionComboBox.ListIndex = 5
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.WriteLine a_message
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    QueryLine = This.ViSession.QueryLine(a_message)
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Read = This.ViSession.Read()
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    Dim p_port As Integer: p_port = Me.Port
    Dim p_sessionReadTimeout As Integer: p_sessionReadTimeout = Me.SessionReadTimeout
      
    If Not ConnectToggleButton.Value And This.K2700.Connected Then
        
        This.K2700.CloseConnection
        
    ElseIf ConnectToggleButton.Value And Not This.K2700.Connected Then
    
        This.K2700.OpenConnection p_host, p_port
        
        This.ViSession.SessionReadTimeout = p_sessionReadTimeout
        
    End If
    
    Me.ReadStopwatch
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Returns true if the inputs are valid for front or rear settings. </summary>
Friend Function ValidateInputs() As Boolean

    ' check if front or rear
    If Me.FrontReadRequired And Not This.K2700.IsFrontInputs Then
    
        ' display the proper message
        Me.LastMessage = "Please release the front panel INPUTS button to its FRONT (F) position and try again."
    
        ValidateInputs = False
        
    ElseIf Not Me.FrontReadRequired And This.K2700.IsFrontInputs Then
    
        ' display the proper message
        Me.LastMessage = "Please depress the front panel INPUTS button to its REAR (R) position and try again."
        
        ValidateInputs = False
    
    Else
    
        ValidateInputs = True
    
    End If

End Function

''' <summary>   Turns on external trigger mode. </summary>
Private Sub ExternalTriggerOptionButton_Click()

    Const p_procedureName = "ExternalTriggerOptionButton_Click"
    
    If Not ExternalTriggerOptionButton.Value Then Exit Sub
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
   
    ' exit if inputs are not in the correct settings
    If Not Me.ValidateInputs() Then GoTo exit_Handler
   
    ' clear execution state when switching mode.
    
    This.K2700.Device.ClearExecutionState
   
    ' Configure external triggering reading requiring the front panel
    
    This.K2700.ConfigurateExternalTriggerReading Me.FunctionComboBox.Text, Me.FrontReadRequired
    
    ' this is done in the previous command.
    ' This.VISession.WriteLine ":TRIG:SOUR IMM"
    ' This.VISession.WriteLine ":TRIG:SOUR EXT"
    
    ' clear the serial poll register
    ' these are also part of the previus command.
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
    
    MeasureButton.Enabled = False
    
    ' enable toggling the external trigger monitoring.
    
    Me.MonitorExtTrigToggleButton.Enabled = True
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    

End Sub

''' <summary>   Selects the measurement function. </summary>
Private Sub FunctionComboBox_Change()

    Dim p_measurement As String
    p_measurement = Me.FunctionComboBox.Text
    Select Case p_measurement
    
        Case "CONT"
        
            FunctionName = "Continuity"
            FunctionUnit = "T/F"
        
        Case "CURR:AC"
        
            FunctionName = "Current AC"
            FunctionUnit = "A"
        
        Case "CURR:DC"
        
            FunctionName = "Current DC"
            FunctionUnit = "A"
        
        Case "FREQ"
        
            FunctionName = "Frequency"
            FunctionUnit = "Hz"
        
        Case "FRES"
        
            FunctionName = "4-Wire Res."
            FunctionUnit = VBA.Chr$(234)
        
        Case "PER"
        
            FunctionName = "Period"
            FunctionUnit = "s"
        
        Case "RES"
        
            FunctionName = "2-Wire Res."
            FunctionUnit = VBA.Chr$(234)
        
        Case "TEMP"
        
            FunctionName = "Temperature"
            FunctionUnit = "K"
        
        Case "VOLT:AC"
        
            FunctionName = "Voltage AC"
            FunctionUnit = "v"
        
        Case "VOLT:DC"
        
            FunctionName = "Voltage DC"
            FunctionUnit = "v"
        
        Case Else
        
            FunctionName = "unknown"
            FunctionUnit = "?"
    
    End Select
    
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
Private Sub ImmediateTriggerOptionButton_Click()

    Const p_procedureName = "ImmediateTriggerOptionButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' exit if inputs are not in the correct settings
    If Not Me.ValidateInputs() Then GoTo exit_Handler
    
    ' clear execution state when switching mode.
    
    This.K2700.Device.ClearExecutionState
    
    ' configure measuring a single value on each read.
    This.K2700.ConfigureSingleFunctionReading Me.FunctionComboBox.Text

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.EnableMeasureButton
    Me.EnableMonitorExtTrigToggleButton

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.SelectiveDeviceClear
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    If This.ViSession.UsingGpibLan Then
    
        Me.RestartStopwatch
        
        Range(m_serialPollCell).Value = This.ViSession.GpibLan.SerialPoll
        
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    If This.ViSession.UsingGpibLan And Me.AutoStatusReadCheckBox.Value Then
        Range(m_serialPollCell).Value = This.ViSession.GpibLan.SerialPoll
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
Friend Property Get Port() As Integer
    Port = CInt(Range(m_inputPortCell).Value)
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
    
    If 0 < This.ViSession.GpibLan.SendToController(a_message) Then
        QueryGpibLan = This.ViSession.GpibLan.ReceiveFromDevice()
    End If
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    ReadGpibLan = This.ViSession.GpibLan.ReceiveFromDevice()
    
    Me.ReadStopwatch

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    WriteGpibLan = This.ViSession.GpibLan.SendToController(a_message)
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.GoToLocal
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.LocalLockout
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.ReadAfterWriteEnabledSetter Me.ReadAfterWriteEnabled
    
    Me.ReadStopwatch
    
    This.ViSession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Me.ReadAfterWriteEnabled = This.ViSession.GpibLan.ReadAfterWriteEnabledGetter
    
    Me.ReadStopwatch
    
    This.ViSession.GpibLan.DisableReadAfterWriteOnWrite = Me.ReadAfterWriteEnabled

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Range(m_serialPollCell).Value = This.ViSession.GpibLan.SerialPoll
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Range(m_srqCell).Value = This.ViSession.GpibLan.ServiceRequested
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Range(m_gpibAddressCell).Value = This.ViSession.GpibLan.GpibAddressGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.GpibAddressSetter CInt(Range(m_gpibAddressCell).Value)

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    Range(m_readTimeoutCell).Value = This.ViSession.GpibLan.ReadTimeoutGetter

    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    
    This.ViSession.GpibLan.ReadTimeoutSetter (CInt(Range(m_readTimeoutCell).Value))
    
    Me.ReadStopwatch
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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

''' <summary> Get the channel count. </summary>
Friend Property Get ChannelCount() As Integer
    ChannelCount = CInt(Range(m_channelCountCell).Value)
End Property

''' <summary>   Indicates that front inputs are required for this mode. </summary>
Friend Property Get FrontReadRequired() As Boolean
    FrontReadRequired = Me.InputsCheckBox.Value
End Property

''' <summary>   Gets the external trigger monitoring status. </summary>
Friend Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = Me.MonitorExtTrigToggleButton.Value
End Property

''' <summary>   Enables or disables external triggering. </summary>
Friend Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    Me.MonitorExtTrigToggleButton.Value = a_value
End Property

''' <summary>   Gets the auto increment channel number enable option. </summary>
Friend Property Get AutoIncrementChannelNoEnabled() As Boolean
    AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
End Property

''' <summary>   Gets the single read enabled. </summary>
Friend Property Get SingeReadEnabled() As Boolean
    SingeReadEnabled = Me.SingleAllCheckBox.Value
End Property

''' <summary>   Gets the selected channel number. </summary>
Friend Property Get SelectedChannelNumber() As Integer
    SelectedChannelNumber = CInt(Range(m_channelNumberCell).Value)
End Property

''' <summary>   Gets the function name from the function name cell. </summary>
Friend Property Get FunctionName() As String
    FunctionName = Range(m_functionNameCell).Value
End Property

''' <summary>   Sets the function name. </summary>
Friend Property Let FunctionName(ByVal a_value As String)
    Range(m_functionNameCell).Value = a_value
End Property

''' <summary>   Gets the function Unit from the function Unit cell. </summary>
Friend Property Get FunctionUnit() As String
    FunctionUnit = Range(m_functionUnitCell).Value
End Property

''' <summary>   Sets the function Unit. </summary>
Friend Property Let FunctionUnit(ByVal a_value As String)
    Range(m_functionUnitCell).Value = a_value
End Property

''' <summary>   Gets the Reading Format from the Reading Format cell. </summary>
Friend Property Get ReadingFormat() As String
    ReadingFormat = Range(m_readingFormatCell).Value
End Property

''' <summary>   Sets the Reading Format. </summary>
Friend Property Let ReadingFormat(ByVal a_value As String)
    Range(m_readingFormatCell).Value = a_value
End Property

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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Takes a single or multiple readings from the front or rear panel. </summary>
Friend Sub MeasureImmediately()

    ' reinitialize if device not defined
    If Not This.K2700.Connected Then
        ' This.K2700.InitializeLegacy Me.host, Me.port, Me.ReceiveTimeout
    End If
    
    Dim p_channelNumber As Integer

    If Me.AutoIncrementChannelNoEnabled Then
        
        ' Clear all cells
        Me.ClearReadings

        ' scan all channels
        For p_channelNumber = 1 To Me.ChannelCount
        
            ' allow events to occur
            VBA.DoEvents
            
            ' set the current channel number
            This.CurrentChannelNumber = p_channelNumber
            
            Me.RestartStopwatch
            If Me.FrontReadRequired Then
                This.K2700.TakeFrontReading p_channelNumber
            Else
                This.K2700.TakeRearReading p_channelNumber
            End If
            Me.ReadStopwatch
        
        Next p_channelNumber

    Else
    
        p_channelNumber = Me.SelectedChannelNumber
        
        Me.RestartStopwatch
        If Me.FrontReadRequired Then
            This.K2700.TakeFrontReading p_channelNumber
        Else
            This.K2700.TakeRearReading p_channelNumber
        End If
        
        Me.ReadStopwatch
        
    End If

End Sub

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()
  
    Const p_procedureName = "MeasureButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' exit if inputs are not in the correct settings
    If Not Me.ValidateInputs() Then GoTo exit_Handler
    
    ' disable the relevant buttons.
    
    Me.MeasureButton.Enabled = False
    Me.ImmediateTriggerOptionButton.Enabled = False
    Me.ExternalTriggerOptionButton.Enabled = False
    
    ' make an immediate measure of one or all channels.
    
    Me.MeasureImmediately

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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Clears the readings. </summary>
''' <param name="a_offsetValue">    [Optional, Double, 0]T he offset value to use for
'''                                 clearing the readings. </param>
Public Sub ClearReadings(Optional ByVal a_offsetValue As Double = 0#)

    Dim p_channelNumber As Integer
    For p_channelNumber = 1 To Me.ChannelCount
        VBA.DoEvents
        SetReading p_channelNumber, a_offsetValue
    Next p_channelNumber
  
End Sub

''' <summary> Toggles monitoring the instrument for external triggered measurements. </summary>
''' <param name="timerInterval"> The polling timer interval. </param>
Friend Sub StartMonitoringExternalTriggers(Optional ByVal a_timerInterval As Integer = 100)

    If Me.AutoIncrementChannelNoEnabled Then
        
        ' Clear all readings
        Me.ClearReadings
    
    End If

    Range(m_externalTriggerStatusCell) = "External Trigger On"
    
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
    Me.StartTimer a_timerInterval
        
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub MonitorExtTrigToggleButton_Click()

    Const p_procedureName = "MonitorExtTrigToggleButton_Click"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Me.ExternalTrigMonitoringEnabled Then
    
        Me.ImmediateTriggerOptionButton.Enabled = False
        Me.ExternalTriggerOptionButton.Enabled = False
    
        Me.StartMonitoringExternalTriggers Me.TimerInterval
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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
        Range(m_socketAddressCell).Value = This.ViSession.SocketAddress
        Range(m_socketIdCell).Value = This.ViSession.SocketId
        
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
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

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Worksheet events
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Initialize this worksheet if not initialized. </summary>
Private Sub Worksheet_Activate()

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not This.Initialized Then
    
        Me.EnableControls False
        
        ' populate the combo box commands
        PopulateCommandsComboBox
        
        ' populate the function box commands
        PopulateFunctionComboBox
        
        Set This.K2700 = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize(New ErrTracer)
        Set m_k2700 = This.K2700
        Set This.Device = This.K2700.Device
        Set This.ViSession = This.Device.ViSession
        
        Set This.Stopwatch = cc_isr_Core_IO.Factory.NewStopwatch()
    
    End If
    
    This.Initialized = True

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.name, ThisWorkbook
    
    ' display the error message
    Me.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


