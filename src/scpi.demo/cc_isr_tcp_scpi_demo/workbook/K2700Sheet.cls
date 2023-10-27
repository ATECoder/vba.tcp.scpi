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

Private Const m_sessionTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_socketAddressCell As String = "C2"
Private Const m_socketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "J2"
Private Const m_receivedMessageCell As String = "K2"
Private Const m_elapsedTimeCell As String = "L2"
Private Const m_elapsedTimeFormatCell As String = "L3"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B33"

Private Const m_serialPollCell As String = "E11"
Private Const m_statusByteCell As String = "F11"
Private Const m_standardByteCell As String = "H11"

Private Const m_srqCell As String = "C10"
Private Const m_gpibAddressCell As String = "C12"
Private Const m_readTimeoutCell As String = "C15"
Private Const m_autoAssertTalkCell As String = "C18"

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

Private Const m_readingOffsetCell As String = "L4"
Private Const m_measurementTopRow As Integer = 2
Private Const m_measurementNoCol As String = "M"
Private Const m_measurementValueCol As String = "N"

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_
    Initialized As Boolean
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    FocusedControlName As String
    DesignMode As Boolean
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Construction and cleanup
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize() As K2700Sheet
    
    Set Initialize = Me

End Function

''' <summary>   Initializes the views. </summary>
Private Sub InitializeViews()

    If Not K2700Observer.Initialized Then
    
        Dim p_viewModel As cc_isr_Tcp_Scpi.K2700ViewModel
        
        Set p_viewModel = cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        
        ' initialize the views before initializing the view model
        ' the view model gets initialized upon connecting.
        DataView.Initialize p_viewModel
        UserView.Initialize p_viewModel
    
    End If
    
    UserView.AutoSampleFrontInputs = Me.FrontAutoScanCheckBox.Value

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the socket address. </summary>
''' <value>   [String]. </value
Public Property Get SocketAddress() As String
    SocketAddress = VBA.CStr(Range(m_socketAddressCell).Value)
End Property

Public Property Let SocketAddress(ByVal a_value As String)
    Range(m_socketAddressCell).Value = a_value
End Property

''' <summary>   Gets or sets the session timeout. </summary>
''' <value>   [Long]. </value>
Public Property Get SessionTimeout() As Long
    SessionTimeout = VBA.CLng(Range(m_sessionTimeoutCell).Value)
End Property

Public Property Let SessionTimeout(ByVal a_value As Long)
    Range(m_sessionTimeoutCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Toggle connection button click event. </summary>
Friend Sub ConnectToggleButton_Click()

    If This.ViewModel Is Nothing Then
    
        Set This.ViewModel = cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        
        ' initialize the observer before initializing the view mode
        ' but after the observer setting are set. The observer initial
        ' settings are then applied to the view model.
        Set This.Observer = K2700Observer.Initialize(This.ViewModel)
        
    End If
    
    ' issue the open connection command. This initializes the view model.
    This.ViewModel.OpenConnectionCommand


    If ConnectToggleButton.Value And This.ViewModel.CanConnect Then
    
        ' the view mode is initialized when connected.
        This.ViewModel.OpenConnectionCommand
    
    ElseIf Not ConnectToggleButton.Value Then
    
        If Not This.ViewModel Is Nothing Then
            If This.ViewModel.Connected Then _
                This.ViewModel.CloseConnectionCommand
            This.ViewModel.Dispose
        End If
    
    End If
    
End Sub

Private Sub ConnectToggleButton_GotFocus()
    This.FocusedControlName = Me.ConnectToggleButton.Name
End Sub

Private Sub ConnectToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
    
    Dim p_details As String
    This.ViewModel.ClearExecutionStateCommand p_details
    Range(m_lastErrorMessageCell).Value = p_details
    
End Sub

Private Sub CommandsComboBox_GotFocus()
    This.FocusedControlName = Me.CommandsComboBox.Name
    DoEvents
End Sub

Private Sub CommandsComboBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the Commands Combo Box change event. </summary>
Private Sub CommandsComboBox_Change()
    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.CommandsComboBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.MessageToSend = Me.CommandsComboBox.Text
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.MessageToSend = Me.CommandsComboBox.Text

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.MessageToSend = Me.CommandsComboBox.Text

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Public Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.CommandsComboBox.Clear
    Dim p_item As Variant
    Dim p_scpiCommand As String
    For Each p_item In This.ViewModel.ScpiCommands
        p_scpiCommand = VBA.CStr(p_item)
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    
    This.ViewModel.QueryCommand

End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()

    This.ViewModel.ReadCommand
    
End Sub

Private Sub ReadStatusAfterWriteCheckBox_GotFocus()
    This.FocusedControlName = Me.ReadStatusAfterWriteCheckBox.Name
    DoEvents
End Sub

Private Sub ReadStatusAfterWriteCheckBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the reading status after write changed event. </summary>
Private Sub ReadStatusAfterWriteCheckBox_Change()

    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.ReadStatusAfterWriteCheckBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
    
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStatusByteButton_Click()
    
    This.ViewModel.QueryStatusByteCommand
    
End Sub


''' <summary>   Handles the Read Standard Event Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()

    This.ViewModel.ReadStandardEventStatusCommand
    
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()

    Dim p_details As String
    This.ViewModel.ResetKnownStateCommand p_details
    Range(m_lastErrorMessageCell).Value = p_details
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()

    This.ViewModel.WriteCommand
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the primary Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Integer
    TimerInterval = VBA.CInt(Range(m_gpibAddressCell).Value)
End Property

Public Property Let PrimaryGpibAddress(ByVal a_value As Integer)
    Range(m_gpibAddressCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()

    This.ViewModel.GoToLocalCommand
    
End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GpibAddressGetButton_Click()

    This.ViewModel.GpibAddressGetCommand

End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Private Sub GpibAddressSetButton_Click()

    This.ViewModel.GpibAddressSetCommand

End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()

    This.ViewModel.LocalLockoutCommand

End Sub

''' <summary>   Handles the read after write set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()

    This.ViewModel.AutoAssertTalkSetCommand

End Sub

''' <summary>   Handles the read after write get button click event. </summary>
Private Sub ReadAfterWriteGetButton_Click()

    This.ViewModel.AutoAssertTalkGetCommand

End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()

    This.ViewModel.GpibLanReadTimeoutGetCommand

End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()

    This.ViewModel.GpibLanReadTimeoutSetCommand

End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    
    This.ViewModel.SelectiveDeviceClearCommand
    
End Sub

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()

    This.ViewModel.SerialPollCommand

End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()

    This.ViewModel.ServiceRequestedQueryCommand

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary> Get or sets the timer interval. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = VBA.CInt(Range(m_timerIntervalCell).Value)
End Property

Public Property Let TimerInterval(ByVal a_value As Integer)
    Range(m_timerIntervalCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub AutoIncrementCheckBox_GotFocus()
    This.FocusedControlName = Me.AutoIncrementCheckBox.Name
    DoEvents
End Sub

Private Sub AutoIncrementCheckBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the auto increment check box change and click events. </summary>
Private Sub AutoIncrementCheckBox_Change()

    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.AutoIncrementCheckBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
    
End Sub

Private Sub ClearReadingsButton_Click()
    This.ViewModel.ClearReadingsCommand
End Sub

Private Sub ExternalTriggerOptionButton_GotFocus()
    This.FocusedControlName = Me.ExternalTriggerOptionButton.Name
    DoEvents
End Sub

Private Sub ExternalTriggerOptionButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the external trigger option button change and click events. </summary>
Private Sub ExternalTriggerOptionButton_Change()

    On Error GoTo err_Handler:
    Dim p_details As String

    If This.FocusedControlName = Me.ExternalTriggerOptionButton.Name Then

        On Error GoTo 0

        ' Commands for click
        If Me.ExternalTriggerOptionButton.Value Then _
            This.ViewModel.ConfigureExternalTriggerReadingCommand p_details
        
        Range(m_lastErrorMessageCell).Value = p_details
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    
End Sub

Private Sub FunctionComboBox_GotFocus()
    This.FocusedControlName = Me.FunctionComboBox.Name
    DoEvents
End Sub

Private Sub FunctionComboBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Selects the measurement function. </summary>
Private Sub FunctionComboBox_Change()
    
    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.FunctionComboBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.SenseFunctionName = Me.FunctionComboBox.Text
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.SenseFunctionName = Me.FunctionComboBox.Text

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.SenseFunctionName = Me.FunctionComboBox.Text
    

End Sub

Private Sub ImmediateTriggerOptionButton_GotFocus()
    This.FocusedControlName = Me.ImmediateTriggerOptionButton.Name
    DoEvents
End Sub

Private Sub ImmediateTriggerOptionButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
''' <remarks>   The check box change event also causes a click event.
'''             This method issues the click event on actual click.
'''             Note: Make sure to not implement the click event,
Private Sub ImmediateTriggerOptionButton_Change()

    On Error GoTo err_Handler:
    
    Dim p_details As String

    If This.FocusedControlName = Me.ImmediateTriggerOptionButton.Name Then

        On Error GoTo 0

        ' Commands for click
        If ImmediateTriggerOptionButton.Value Then _
            This.ViewModel.ConfigureImmediateTriggerReadingsCommand p_details
    
        Range(m_lastErrorMessageCell).Value = p_details
    
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    
End Sub

Private Sub InputsCheckBox_GotFocus()
    This.FocusedControlName = Me.InputsCheckBox.Name
    DoEvents
End Sub

Private Sub InputsCheckBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub


''' <summary>   Handles the Inputs check box change event. </summary>
Private Sub InputsCheckBox_Change()

    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.InputsCheckBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.FrontInputsRequired = Me.InputsCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.FrontInputsRequired = Me.InputsCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.FrontInputsRequired = Me.InputsCheckBox.Value
    
End Sub

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()

    Dim p_details As String
    
    ' make an immediate measure of one or all channels.
    This.ViewModel.MeasureImmediatelyCommand p_details
    
    Range(m_lastErrorMessageCell).Value = p_details
    
  
End Sub

Private Sub MonitorExtTrigToggleButton_GotFocus()
    This.FocusedControlName = Me.MonitorExtTrigToggleButton.Name
    DoEvents
End Sub

Private Sub MonitorExtTrigToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub MonitorExtTrigToggleButton_Change()

    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.MonitorExtTrigToggleButton.Name Then

        On Error GoTo 0

        If MonitorExtTrigToggleButton.Value Then
            This.ViewModel.StartMonitoringExternalTriggersCommand
        Else
            This.ViewModel.StopMonitoringExternalTriggersCommand
        End If


    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        If MonitorExtTrigToggleButton.Value Then
            This.ViewModel.StartMonitoringExternalTriggersCommand
        Else
            This.ViewModel.StopMonitoringExternalTriggersCommand
        End If

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    If MonitorExtTrigToggleButton.Value Then
        This.ViewModel.StartMonitoringExternalTriggersCommand
    Else
        This.ViewModel.StopMonitoringExternalTriggersCommand
    End If
    
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

    This.ViewModel.QueryInputsCommand

End Sub

Private Sub SingleMultipleCheckBox_GotFocus()
    This.FocusedControlName = Me.SingleMultipleCheckBox.Name
    DoEvents
End Sub

Private Sub SingleMultipleCheckBox_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

Private Sub SingleMultipleCheckBox_Change()

    On Error GoTo err_Handler:

    If This.FocusedControlName = Me.SingleMultipleCheckBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
    
End Sub

''' <summary>   Populates the Function combo box. </summary>
Public Sub PopulateFunctionComboBox()

    Const p_procedureName = "PopulateFunctionComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_fourWireResistanceIndex As Integer
    Me.FunctionComboBox.Clear
    Dim p_item As Variant
    Dim p_senseFunction As String
    For Each p_item In This.ViewModel.SenseFunctions
        p_senseFunction = VBA.CStr(p_item)
        Me.FunctionComboBox.AddItem p_senseFunction
        If "FRES" = p_senseFunction Then
            p_fourWireResistanceIndex = Me.FunctionComboBox.Count
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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Events
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the measurement display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Private Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)

    Dim p_rowNo As String: p_rowNo = VBA.CStr(m_measurementTopRow + a_channelNo - 1)
    Dim p_cell As String
    
    p_cell = m_measurementNoCol & p_rowNo
    Range(p_cell).Value = VBA.CStr(a_channelNo)
    
    p_cell = m_measurementValueCol & p_rowNo
    Range(p_cell).Value = Format(a_value, This.ViewModel.MeasuredValueFormat)
    
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_viewModel_ChannelMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        SetReading a_eventArgs.ChannelNumber, -1#
    Else
        SetReading a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Private Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "m_K2700ViewModel_PropertyChanged"

    Select Case a_propertyName

        Case "Host"
            Range(m_inputHostCell).Value = This.ViewModel.Host
        Case "Port"
            Range(m_inputPortCell).Value = This.ViewModel.Port
        Case "SocketAddress"
            Range(m_socketAddressCell).Value = This.ViewModel.SocketAddress
        Case "SocketId"
            Range(m_socketIdCell).Value = This.ViewModel.SocketId
        Case "SessionTimeout"
            Range(m_sessionTimeoutCell).Value = This.ViewModel.SessionTimeout
        Case "PauseRequested"
        Case "TimerInterval"
            Range(m_timerIntervalCell).Value = This.ViewModel.TimerInterval
        Case "Connected"
            Range(m_connectedCell).Value = This.ViewModel.Connected
    
        Case "ElapsedTime"
            ' Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTime
        Case "ElapsedTimeFormat"
            Range(m_elapsedTimeFormatCell).Value = This.ViewModel.ElapsedTimeFormat
        Case "ElapsedTimeCaption"
            Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTimeCaption
        
        Case "MessageToSend"
        Case "SentMessage"
            Range(m_sentMessageCell).Value = This.ViewModel.SentMessage
        Case "ReceivedMessage"
            Range(m_receivedMessageCell).Value = This.ViewModel.ReceivedMessage
        Case "ReceivedMessageLength"
            Range(m_receivedMessageLengthCell).Value = This.ViewModel.ReceivedMessageLength
        Case "ReadStatusAfterWriteEnabled"
            Me.ReadStatusAfterWriteCheckBox.Value = This.ViewModel.ReadStatusAfterWriteEnabled
        Case "LastErrorMessage"
            Range(m_lastErrorMessageCell).Value = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            Range(m_lastMessageCell).Value = This.ViewModel.LastMessage
        Case "SerialPollByte"
            Range(m_serialPollCell).Value = This.ViewModel.SerialPollByte
        Case "StatusByte"
            Range(m_statusByteCell).Value = This.ViewModel.StatusByte
        Case "StandardByte"
            Range(m_standardByteCell).Value = This.ViewModel.StandardByte
        Case "ServiceRequested"
            Range(m_srqCell).Value = This.ViewModel.ServiceRequested
        Case "GpibAddress"
            Range(m_gpibAddressCell).Value = This.ViewModel.GpibAddress
        Case "AutoAssertTalk"
            Range(m_autoAssertTalkCell).Value = This.ViewModel.AutoAssertTalk
        Case "GpibLanReadTimeout"
            Range(m_autoAssertTalkCell).Value = This.ViewModel.GpibLanReadTimeout
            
        Case "TopCard"
            Range(m_card1Cell).Value = This.ViewModel.TopCard
        Case "BottomCard"
            Range(m_card2Cell).Value = This.ViewModel.BottomCard
        Case "TopCardFunctionScanList"
            Range(m_card1ScanListCell).Value = This.ViewModel.TopCardFunctionScanList
        Case "BottomCardFunctionScanList"
            Range(m_card2ScanListCell).Value = This.ViewModel.BottomCardFunctionScanList
        Case "FrontInputsRequired"
            Me.InputsCheckBox.Value = This.ViewModel.FrontInputsRequired
        Case "FrontInputsValue"
        Case "FrontInputsCaption"
            Range(m_inputsCell).Value = This.ViewModel.FrontInputsCaption
        Case "SelectedChannelNumber"
            Range(m_channelNumberCell).Value = This.ViewModel.SelectedChannelNumber
        Case "ChannelCount"
            Range(m_channelCountCell).Value = This.ViewModel.ChannelCount
        Case "CurrentChannelNumber"
        Case "MeasuredChannelNumber"
        Case "SenseFunctionName"
        Case "SenseFunctionCaption"
            Range(m_functionNameCell).Value = This.ViewModel.SenseFunctionCaption
        Case "SenseFunctionUnit"
            Range(m_functionUnitCell).Value = This.ViewModel.SenseFunctionUnit
        Case "MeasuredReading"
            Range(m_functionValueCell).Value = This.ViewModel.MeasuredReading
        Case "MeasuredValueFormat"
            Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        
        Case "MeasuredValue"
        Case "ExternalTrigMonitoringEnabled"
            Me.MonitorExtTrigToggleButton.Value = This.ViewModel.ExternalTrigMonitoringEnabled
        Case "MonitoringState"
            ' Range(m_readingFormatCell).Value = This.ViewModel.MonitoringState
        Case "MeasuredValueFormat"
            Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        Case "AutoIncrementChannelNoEnabled"
            Me.AutoIncrementCheckBox.Value = This.ViewModel.AutoIncrementChannelNoEnabled
        Case "SingleReadEnabled"
            Me.SingleMultipleCheckBox.Value = This.ViewModel.SingleReadEnabled
        Case "SenseFunctions"
            Me.PopulateFunctionComboBox
        Case "ScpiCommands"
            Me.PopulateCommandsComboBox
            
        ' !@# do we need this?
        Case "ImmediateTriggerOptionSelected"
            ' Me.ImmediateTriggerOptionButton.Value = This.ViewModel.ImmediateTriggerOptionSelected
        ' !@# do we need this?
        Case "ExternalTriggerOptionSelected"
            ' Me.ExternalTriggerOptionButton.Value = This.ViewModel.ExternalTriggerOptionSelected
        Case "ReadingOffset"
            Range(m_readingOffsetCell).Value = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Range(m_externalTriggerStatusCell).Value = This.ViewModel.MonitoringState
        
        Case "OpenConnectionExecutable"
        Case "CloseConnectionExecutable"
        Case "ToggleConnectionExecutable"
            ConnectToggleButton.Enabled = This.ViewModel.ToggleConnectionExecutable
        Case "ResetKnownStateExecutable"
            RSTButton.Enabled = This.ViewModel.ResetKnownStateExecutable
        Case "ClearExecutionStateExecutable"
            CLSButton.Enabled = This.ViewModel.ClearExecutionStateExecutable
        Case "QueryExecutable"
            QueryButton.Enabled = This.ViewModel.QueryExecutable
        Case "ReadExecutable"
            ReadButton.Enabled = This.ViewModel.ReadExecutable
        Case "ReadStatusByteExecutable"
            ReadStatusByteButton.Enabled = This.ViewModel.ReadStatusByteExecutable
        Case "ReadStandardEventStatusExecutable"
            ReadStandardEventStatusButton.Enabled = This.ViewModel.ReadStandardEventStatusExecutable
        Case "ReadStatusAfterWriteExecutable"
            ReadStatusAfterWriteCheckBox.Enabled = This.ViewModel.ReadStandardEventStatusExecutable
        Case "WriteExecutable"
            WriteButton.Enabled = This.ViewModel.WriteExecutable
        Case "QueryInputsExecutable"
            QueryInputsButton.Enabled = This.ViewModel.QueryInputsExecutable
        Case "ClearReadingsExecutable"
            ClearReadingsButton.Enabled = This.ViewModel.ClearReadingsExecutable
        Case "MeasureExecutable"
            MeasureButton.Enabled = This.ViewModel.MeasureExecutable
        Case "StopMonitoringExecutable"
            MonitorExtTrigToggleButton.Enabled = This.ViewModel.StopMonitoringExecutable Or _
                This.ViewModel.StartMonitoringExecutable
        Case "StartMonitoringExecutable"
            MonitorExtTrigToggleButton.Enabled = This.ViewModel.StopMonitoringExecutable Or _
                This.ViewModel.StartMonitoringExecutable
        Case "ImmediateTriggerOptionExecutable"
            ImmediateTriggerOptionButton.Enabled = This.ViewModel.ImmediateTriggerOptionExecutable
        Case "ExternalTriggerOptionExecutable"
            ExternalTriggerOptionButton.Enabled = This.ViewModel.ExternalTriggerOptionExecutable

        Case "SelectiveDeviceClearExecutable"
            SDCButton.Enabled = This.ViewModel.SelectiveDeviceClearExecutable
        Case "GoToLocalExecutable"
            GoToLocalButton.Enabled = This.ViewModel.GoToLocalExecutable
        Case "LocalLockoutExecutable"
            LocalLockoutButton.Enabled = This.ViewModel.LocalLockoutExecutable
        Case "AutoAssertTalkSetExecutable"
            ReadAfterWriteSetButton.Enabled = This.ViewModel.AutoAssertTalkSetExecutable
        Case "AutoAssertTalkGetExecutable"
            ReadAfterWriteGetButton.Enabled = This.ViewModel.AutoAssertTalkGetExecutable
        Case "ReadStatusByteExecutable"
            SerialPollButton.Enabled = This.ViewModel.ReadStatusByteExecutable
        Case "ServiceRequestedExecutable"
            SRQButton.Enabled = This.ViewModel.ServiceRequestedExecutable
        Case "GpibAddressGetExecutable"
            GpibAddressGetButton.Enabled = This.ViewModel.GpibAddressGetExecutable
        Case "GpibAddressSetExecutable"
            GpibAddressSetButton.Enabled = This.ViewModel.GpibAddressSetExecutable
        Case "GpibLanReadTimeoutGetExecutable"
            ReadTimeoutGetButton.Enabled = This.ViewModel.GpibLanReadTimeoutGetExecutable
        Case "GpibLanReadTimeoutSetExecutable"
            ReadTimeoutSetButton.Enabled = This.ViewModel.GpibLanReadTimeoutSetExecutable
        
        
        Case "MessageToSendExecutable"
            CommandsComboBox.Enabled = This.ViewModel.MessageToSendExecutable
        Case "FunctionNameExecutable"
            FunctionComboBox.Enabled = This.ViewModel.FunctionNameExecutable

    End Select
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_K2700ViewModel_PropertyChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    HandleViewModelPropertyChanged a_eventArgs.PropertyName

    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Worksheet events
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Sub Worksheet_Deactivate()

    Me.Dispose
    
End Sub

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set m_viewModel = Nothing
    
    If Not This.ViewModel Is Nothing Then This.ViewModel.Dispose
    Set This.ViewModel = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub OnWorksheetChange(ByVal a_range As Range)

    ' If Not Intersect(a_range, Me.Range("H5")) Is Nothing Then
    ' this uses fewer resources than Intersect,
    ' which will be helpful if your worksheet changes a lot.
    Select Case a_range.Address
        
        Case m_sessionTimeoutCell
            This.ViewModel.SessionTimeout = Range(m_sessionTimeoutCell).Value
        Case m_inputHostCell
            This.ViewModel.SocketAddress = Range(m_inputHostCell).Value & ":" & _
                Range(m_inputPortCell).Value
        Case m_inputPortCell
            This.ViewModel.SocketAddress = Range(m_inputHostCell).Value & ":" & _
                Range(m_inputPortCell).Value
        Case m_socketAddressCell
            This.ViewModel.SocketAddress = Range(m_socketAddressCell).Value
        Case m_socketIdCell
            ' This.ViewModel.SocketId = Range(m_socketIdCell).Value
        Case m_sessionTimeoutCell
            This.ViewModel.SessionTimeout = Range(m_sessionTimeoutCell).Value
            
        Case m_connectedCell
            ' connected is read only
            ' This.ViewModel.Connected = Range(m_connectedCell).Value
        Case m_sentMessageCell
            This.ViewModel.SentMessage = Range(m_sentMessageCell).Value
        Case m_receivedMessageLengthCell
            This.ViewModel.ReceivedMessageLength = Range(m_receivedMessageLengthCell).Value
        Case m_receivedMessageCell
            This.ViewModel.ReceivedMessage = Range(m_receivedMessageCell).Value
        Case m_elapsedTimeCell
            This.ViewModel.ElapsedTimeCaption = Range(m_elapsedTimeCell).Value
        Case m_elapsedTimeFormatCell
            This.ViewModel.ElapsedTimeFormat = Range(m_elapsedTimeFormatCell).Value
        Case m_lastMessageCell
            This.ViewModel.LastMessage = Range(m_lastMessageCell).Value
        Case m_lastErrorMessageCell
            This.ViewModel.LastErrorMessage = Range(m_lastErrorMessageCell).Value
        Case m_serialPollCell
            This.ViewModel.SerialPollByte = Range(m_serialPollCell).Value
        Case m_statusByteCell
            This.ViewModel.StatusByte = Range(m_statusByteCell).Value
        Case m_standardByteCell
            This.ViewModel.StandardByte = Range(m_standardByteCell).Value

        Case m_srqCell
            This.ViewModel.ServiceRequested = Range(m_srqCell).Value
        Case m_gpibAddressCell
            This.ViewModel.GpibAddress = Range(m_gpibAddressCell).Value
        Case m_readTimeoutCell
            This.ViewModel.GpibLanReadTimeout = Range(m_readTimeoutCell).Value
        Case m_autoAssertTalkCell
            This.ViewModel.AutoAssertTalk = VBA.CBool(Range(m_autoAssertTalkCell).Value)
        Case m_timerIntervalCell
            This.ViewModel.TimerInterval = Range(m_timerIntervalCell).Value

        Case m_readingFormatCell
            This.ViewModel.MeasuredValueFormat = Range(m_readingFormatCell).Value
        Case m_card1Cell
            This.ViewModel.TopCard = Range(m_card1Cell).Value
        Case m_card2Cell
            This.ViewModel.TopCard = Range(m_card2Cell).Value
        Case m_card1ScanListCell
            This.ViewModel.TopCardFunctionScanList = Range(m_card1ScanListCell).Value
        Case m_card2ScanListCell
            This.ViewModel.BottomCardFunctionScanList = Range(m_card2ScanListCell).Value
        Case m_inputsCell
            This.ViewModel.FrontInputsCaption = Range(m_inputsCell).Value
        Case m_channelNumberCell
            This.ViewModel.SelectedChannelNumber = Range(m_channelNumberCell).Value
        Case m_channelCountCell
            This.ViewModel.ChannelCount = Range(m_channelCountCell).Value
        Case m_functionNameCell
            This.ViewModel.SenseFunctionCaption = Range(m_functionNameCell).Value
        Case m_functionValueCell
            This.ViewModel.MeasuredReading = Range(m_functionValueCell).Value
        Case m_functionUnitCell
            This.ViewModel.SenseFunctionUnit = Range(m_functionUnitCell).Value
        Case m_externalTriggerStatusCell
            This.ViewModel.MonitoringState = Range(m_externalTriggerStatusCell).Value
        Case m_readingOffsetCell
            This.ViewModel.ReadingOffset = Range(m_readingOffsetCell).Value

    End Select

End Sub


''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub Worksheet_Change(ByVal a_range As Range)

    OnWorksheetChange a_range

End Sub

 
