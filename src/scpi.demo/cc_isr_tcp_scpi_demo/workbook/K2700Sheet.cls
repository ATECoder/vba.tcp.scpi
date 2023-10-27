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
Private Const m_socketHostCell As String = "B2"
Private Const m_socketPortCell As String = "B3"
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
Private Const m_monitoringStateCell As String = "C31"
Private Const m_measuredChannelNumberCell As String = "D28"

Private Const m_readingOffsetCell As String = "L4"
Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 14 ' 2
Private Const m_channelRow As Integer = 2 ' 7
Private Const m_channelCol As Integer = 13 ' 1
Private Const m_byRow As Boolean = True ' False
Private Const m_displayIncrement As Integer = 1 '  2

Private Const m_maximumChannelNumber As Integer = 48

Private Const m_gpibLanControllerPort As Long = 1234

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_
    Initialized As Boolean
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    FocusedControlName As String
    DesignMode As Boolean
    MeasuredValue As Double
    Connected As Boolean
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
    
        Set m_viewModel = cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        
        ' initialize the views before initializing the view model
        ' the view model gets initialized upon connecting.
        K2700Observer.Initialize m_viewModel

        Me.Synchronize
        

    End If
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the socket address. </summary>
''' <value>   [String]. </value
Public Property Get SocketAddress() As String
    SocketAddress = VBA.CStr(K2700Sheet.Range(m_socketAddressCell).Value)
End Property

Public Property Let SocketAddress(ByVal a_value As String)
    K2700Sheet.Range(m_socketAddressCell).Value = a_value
End Property

''' <summary>   Gets or sets the session timeout. </summary>
''' <value>   [Long]. </value>
Public Property Get SessionTimeout() As Long
    SessionTimeout = VBA.CLng(K2700Sheet.Range(m_sessionTimeoutCell).Value)
End Property

Public Property Let SessionTimeout(ByVal a_value As Long)
    K2700Sheet.Range(m_sessionTimeoutCell).Value = a_value
End Property

''' <summary>   Gets or sets the connected state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property

Public Property Let Connected(ByVal a_value As Boolean)
    If This.Connected <> a_value Then
        This.Connected = a_value
        K2700Sheet.Range(m_connectedCell).Value = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the execution (Can Execute) status of the Open Connection command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = Me.OpenConnectionButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Open Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    Me.OpenConnectionButton.Enabled = a_value
    ' allow power on reset only if not connected.
    Me.PowerOnResetButton.Enabled = a_value
End Property

''' <summary>   Handles the open Connection command event. </summary>
Private Sub OnOpenConnection()

    ' initialize the views if necessary
    InitializeViews
    
    K2700Observer.OnOpenConnectionButtonClick

End Sub

''' <summary>   Handles open connection button click event. </summary>
Private Sub OpenConnectionButton_Click()
    OnOpenConnection
End Sub

''' <summary>   Gets the execution (Can Execute) status of the Close Connection command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    CloseConnectionExecutable = Me.CloseConnectionButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Close Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    Me.CloseConnectionButton.Enabled = a_value
End Property

''' <summary>   Handles the open Connection command event. </summary>
Private Sub OnCloseConnection()

    If K2700Observer.Initialized Then
    
        K2700Observer.OnCloseConnectionButtonClick
    
    End If

End Sub

''' <summary>   Handles close connection button click event. </summary>
Private Sub CloseConnectionButton_Click()
    OnCloseConnection
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
    K2700Observer.OnClearExecutionState
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
        K2700Observer.MessageToSend = Me.CommandsComboBox.Text
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        K2700Observer.MessageToSend = Me.CommandsComboBox.Text

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    K2700Observer.MessageToSend = Me.CommandsComboBox.Text

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Public Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.CommandsComboBox.Clear
    Dim p_item As Variant
    Dim p_scpiCommand As String
    For Each p_item In K2700Observer.ScpiCommands
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
        K2700Observer.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        K2700Observer.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    K2700Observer.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
    
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStatusByteButton_Click()
    K2700Observer.OnQueryStatusByte
End Sub

''' <summary>   Handles the Read Standard Event Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()
    K2700Observer.OnReadStandardEventStatus
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    K2700Observer.OnResetKnownState
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
    TimerInterval = VBA.CInt(K2700Sheet.Range(m_gpibAddressCell).Value)
End Property

Public Property Let PrimaryGpibAddress(ByVal a_value As Integer)
    K2700Sheet.Range(m_gpibAddressCell).Value = a_value
End Property

''' <summary>   Gets the Gpib Lan Controller Port number. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = m_gpibLanControllerPort
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()
    K2700Observer.OnGoToLocal
End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GpibAddressGetButton_Click()
    K2700Observer.OnGpibAddressGet
End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Private Sub GpibAddressSetButton_Click()
    K2700Observer.OnGpibAddressSet
End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()
    K2700Observer.OnLocalLockout
End Sub

''' <summary>   Handles the click event of the power on reset.  </summary>
Private Sub PowerOnResetButton_Click()
    OnPowerOnReset
End Sub

''' <summary>   Handles the power on reset command event. </summary>
Private Sub OnPowerOnReset()

    ' initialize the views if necessary
    InitializeViews
    
    K2700Observer.OnPowerOneResetButtonClick

End Sub

''' <summary>   Handles the read after write set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()
    K2700Observer.OnAutoAssertTalkSet
End Sub

''' <summary>   Handles the read after write get button click event. </summary>
Private Sub ReadAfterWriteGetButton_Click()
    K2700Observer.OnAutoAssertTalkGet
End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()
    K2700Observer.OnReadTimeoutGet
End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()
    K2700Observer.OnReadTimeoutSet
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    K2700Observer.OnSDC
End Sub

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()
    K2700Observer.OnReadStatusByte
End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()
    K2700Observer.OnSRQ
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the maxmum channge count. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumChannelNumber() As Integer
    MaximumChannelNumber = m_maximumChannelNumber
End Property

''' <summary> Get or sets the timer interval. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = VBA.CInt(K2700Sheet.Range(m_timerIntervalCell).Value)
End Property

Public Property Let TimerInterval(ByVal a_value As Integer)
    K2700Sheet.Range(m_timerIntervalCell).Value = a_value
End Property

''' <summary>   Gets the selected channel number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedChannelNumber() As Integer
    SelectedChannelNumber = VBA.CInt(K2700Sheet.Range(m_channelNumberCell).Value)
End Property

''' <summary>   Sets the selected channel number. </summary>
''' <param name="a_value">   [Integer] the selected channel number. </param>
Public Property Let SelectedChannelNumber(ByVal a_value As Integer)
    K2700Sheet.Range(m_channelNumberCell).Value = a_value
End Property

''' <summary>   Gets the Measured channel number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredChannelNumber() As Integer
    MeasuredChannelNumber = VBA.CInt(K2700Sheet.Range(m_measuredChannelNumberCell).Value)
End Property

''' <summary>   Sets the Measured channel number. </summary>
''' <param name="a_value">   [Integer] the Measured channel number. </param>
Public Property Let MeasuredChannelNumber(ByVal a_value As Integer)
    K2700Sheet.Range(m_measuredChannelNumberCell).Value = a_value
End Property

''' <summary>   Gets the Measured reading. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = K2700Sheet.Range(m_functionValueCell).Value
End Property

''' <summary>   Sets the Measured reading. </summary>
''' <param name="a_value">   [String] the Measured reading. </param>
Public Property Let MeasuredReading(ByVal a_value As String)
    K2700Sheet.Range(m_functionValueCell).Value = a_value
End Property

''' <summary>   Gets the reading offset. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadingOffset() As Double
    ReadingOffset = K2700Sheet.Range(m_readingOffsetCell).Value
End Property

''' <summary>   Sets the reading offset. </summary>
''' <param name="a_value">   [Double] the reading offset. </param>
Public Property Let ReadingOffset(ByVal a_value As Double)
    K2700Sheet.Range(m_readingOffsetCell).Value = a_value
End Property

''' <summary>   Gets the Measured Value. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredValue() As Double
    MeasuredValue = This.MeasuredValue
End Property

''' <summary>   Sets the Measured Value. </summary>
''' <param name="a_value">   [Double] the Measured Value. </param>
Public Property Let MeasuredValue(ByVal a_value As Double)
    This.MeasuredValue = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Auto Increment Channel Number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementChannelNoEnabled() As Boolean
    AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
End Property

''' <summary>   Sets the Auto Increment Channel Number Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment Channel Number Enabled value. </value>
Public Property Let AutoIncrementChannelNoEnabled(ByVal a_value As Boolean)
    If Me.AutoIncrementCheckBox.Value <> a_value Then
        Me.AutoIncrementCheckBox.Value = a_value
    End If
End Property

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
        K2700Observer.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        K2700Observer.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    K2700Observer.AutoIncrementChannelNoEnabled = Me.AutoIncrementCheckBox.Value
    
End Sub

''' <summary>   Gets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <value>   [Boolean] True if the Clear Readings command can be executed. </value>
Public Property Get ClearReadingsExecutable() As Boolean
    ClearReadingsExecutable = Me.ClearReadingsButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    Me.ClearReadingsButton.Enabled = a_value
End Property

''' <summary>   Handles the click event of the Clear Radings button. </summary>
Private Sub ClearReadingsButton_Click()
    K2700Observer.OnClearReadings
End Sub

''' <summary>   Gets the External Trig Monitoring Enabled condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopd on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = Me.ExternalTriggerOptionButton.Value
End Property

''' <summary>   Sets the External Trig Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trig Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If Me.ExternalTriggerOptionButton.Value <> a_value Then
        Me.ExternalTriggerOptionButton.Value = a_value
    End If
End Property


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
        If Me.ExternalTriggerOptionButton.Value Then
            K2700Observer.ExternalSenseFunctionName = Me.FunctionComboBox.Text
            K2700Observer.OnExternalTriggerOption
        End If
        
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
        If ImmediateTriggerOptionButton.Value Then
            K2700Observer.ImmediateSenseFunctionName = Me.FunctionComboBox.Text
            K2700Observer.OnImmediateTriggerOption
        End If
    
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
        K2700Observer.FrontInputsRequired = Me.InputsCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        K2700Observer.FrontInputsRequired = Me.InputsCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    K2700Observer.FrontInputsRequired = Me.InputsCheckBox.Value
    
End Sub

''' <summary>   Gets the Front Inputs Required. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsRequired() As Boolean
    FrontInputsRequired = Me.InputsCheckBox.Value
End Property

''' <summary>   Sets the Front Inputs Required. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are Required. </value>
Public Property Let FrontInputsRequired(ByVal a_value As Boolean)
    If Me.InputsCheckBox.Value <> a_value Then
        Me.InputsCheckBox.Value = a_value
        K2700Observer.FrontInputsRequired = a_value
    End If
End Property

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()

    Dim p_details As String
    
    ' make an immediate measure of one or all channels.
    This.ViewModel.MeasureImmediatelyCommand p_details
    
    K2700Sheet.Range(m_lastErrorMessageCell).Value = p_details
    
  
End Sub

''' <summary>   Handles the start monitoring button click event. </summary>
Private Sub StartMonitoringButton_Click()
    K2700Observer.OnStartMonitorExtTrigButtonClick
End Sub

''' <summary>   Handles the stop monitoring button click event. </summary>
Private Sub StopMonitoringButton_Click()
    K2700Observer.OnStopMonitorExtTrigButtonClick
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()

    This.ViewModel.QueryInputsCommand

End Sub

''' <summary>   Gets the Single Read Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get SingleReadEnabled() As Boolean
    SingleReadEnabled = Me.SingleMultipleCheckBox.Value
End Property

''' <summary>   Sets the Single Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Single Read Enabled value. </value>
Public Property Let SingleReadEnabled(ByVal a_value As Boolean)
    If Me.SingleMultipleCheckBox.Value <> a_value Then
        Me.SingleMultipleCheckBox.Value = a_value
    End If
End Property

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
        K2700Observer.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        K2700Observer.SingleReadEnabled = Me.SingleMultipleCheckBox.Value

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    K2700Observer.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
    
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
    For Each p_item In K2700Observer.SenseFunctions
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
' View Model Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the design mode for unit testing and restoration of state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get DesignMode() As Boolean
    DesignMode = This.DesignMode
End Property

''' <summary>   Sets the design mode for unit testing and restoration of state. </summary>
''' <summary>   Sets the design mode for unit testing. </summary>
''' <param name="a_value">   [Boolean] true if in design mode. </param>
Public Property Let DesignMode(ByVal a_value As Boolean)
    This.DesignMode = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = K2700Sheet.Range(m_lastErrorMessageCell).Value
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    K2700Sheet.Range(m_lastErrorMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = K2700Sheet.Range(m_lastMessageCell).Value
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastMessage(ByVal a_value As String)
    K2700Sheet.Range(m_lastMessageCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Measurments View
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Public Property Get MeasurementRow() As Integer
    MeasurementRow = m_measurementRow
End Property
Public Property Get MeasurementCol() As Integer
    MeasurementCol = m_measurementCol
End Property
Public Property Get ChannelRow() As Integer
    ChannelRow = m_channelRow
End Property
Public Property Get ChannelCol() As Integer
    ChannelCol = m_channelCol
End Property
Public Property Get ByRow() As Boolean
    ByRow = m_byRow
End Property
Public Property Get DisplayIncrement() As Integer
    DisplayIncrement = m_displayIncrement
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' View Model Events
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the measurement display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Private Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)

    Dim p_rowNo As String: p_rowNo = VBA.CStr(Me.MeasurementRow + a_channelNo - 1)
    Dim p_cell As String
    
    p_cell = Me.ChannelCol & p_rowNo
    K2700Sheet.Range(p_cell).Value = VBA.CStr(a_channelNo)
    
    p_cell = Me.MeasurementCol & p_rowNo
    K2700Sheet.Range(p_cell).Value = Format(a_value, This.ViewModel.MeasuredValueFormat)
    
End Sub

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_viewModel_ChannelMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        ' let the observer make these updates.
        ' SetReading a_eventArgs.ChannelNumber, -1#
    Else
        ' let the observer make these updates.
        ' SetReading a_eventArgs.ChannelNumber, CDbl(a_eventArgs.Reading)
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

Public Sub Synchronize()

    This.Connected = This.ViewModel.Connected
    This.DesignMode = This.ViewModel.DesignMode
    
    Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
    Me.ReadingOffset = This.ViewModel.ReadingOffset
    Me.ClearReadingsExecutable = This.ViewModel.ClearExecutionStateExecutable
    
    Me.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
    Me.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
    
    Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.LastErrorMessage = This.ViewModel.LastErrorMessage
    Me.LastMessage = This.ViewModel.LastMessage
    Me.SelectedChannelNumber = This.ViewModel.SelectedChannelNumber
        
    Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
    Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
    Me.ExternalTrigMonitoringEnabled = This.ViewModel.ExternalTrigMonitoringEnabled
        
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Private Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "HandleViewModelPropertyChanged"

    Select Case a_propertyName

        Case "Host"
            K2700Sheet.Range(m_socketHostCell).Value = This.ViewModel.Host
        Case "Port"
            K2700Sheet.Range(m_socketPortCell).Value = This.ViewModel.Port
        Case "SocketAddress"
            K2700Sheet.Range(m_socketAddressCell).Value = This.ViewModel.SocketAddress
        Case "SocketId"
            K2700Sheet.Range(m_socketIdCell).Value = This.ViewModel.SocketId
        Case "SessionTimeout"
            K2700Sheet.Range(m_sessionTimeoutCell).Value = This.ViewModel.SessionTimeout
        Case "PauseRequested"
        Case "TimerInterval"
            K2700Sheet.Range(m_timerIntervalCell).Value = This.ViewModel.TimerInterval
        Case "Connected"
            Me.Connected = This.ViewModel.Connected
        Case "ElapsedTime"
            ' K2700Sheet.Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTime
        Case "ElapsedTimeFormat"
            K2700Sheet.Range(m_elapsedTimeFormatCell).Value = This.ViewModel.ElapsedTimeFormat
        Case "ElapsedTimeCaption"
            K2700Sheet.Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTimeCaption
        
        Case "MessageToSend"
        Case "SentMessage"
            K2700Sheet.Range(m_sentMessageCell).Value = This.ViewModel.SentMessage
        Case "ReceivedMessage"
            K2700Sheet.Range(m_receivedMessageCell).Value = This.ViewModel.ReceivedMessage
        Case "ReceivedMessageLength"
            K2700Sheet.Range(m_receivedMessageLengthCell).Value = This.ViewModel.ReceivedMessageLength
        Case "ReadStatusAfterWriteEnabled"
            Me.ReadStatusAfterWriteCheckBox.Value = This.ViewModel.ReadStatusAfterWriteEnabled
        Case "LastErrorMessage"
            K2700Sheet.Range(m_lastErrorMessageCell).Value = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            K2700Sheet.Range(m_lastMessageCell).Value = This.ViewModel.LastMessage
        Case "SerialPollByte"
            K2700Sheet.Range(m_serialPollCell).Value = This.ViewModel.SerialPollByte
        Case "StatusByte"
            K2700Sheet.Range(m_statusByteCell).Value = This.ViewModel.StatusByte
        Case "StandardByte"
            K2700Sheet.Range(m_standardByteCell).Value = This.ViewModel.StandardByte
        Case "ServiceRequested"
            K2700Sheet.Range(m_srqCell).Value = This.ViewModel.ServiceRequested
        Case "GpibAddress"
            K2700Sheet.Range(m_gpibAddressCell).Value = This.ViewModel.GpibAddress
        Case "AutoAssertTalk"
            K2700Sheet.Range(m_autoAssertTalkCell).Value = This.ViewModel.AutoAssertTalk
        Case "GpibLanReadTimeout"
            K2700Sheet.Range(m_autoAssertTalkCell).Value = This.ViewModel.GpibLanReadTimeout
            
        Case "TopCard"
            K2700Sheet.Range(m_card1Cell).Value = This.ViewModel.TopCard
        Case "BottomCard"
            K2700Sheet.Range(m_card2Cell).Value = This.ViewModel.BottomCard
        Case "TopCardFunctionScanList"
            K2700Sheet.Range(m_card1ScanListCell).Value = This.ViewModel.TopCardFunctionScanList
        Case "BottomCardFunctionScanList"
            K2700Sheet.Range(m_card2ScanListCell).Value = This.ViewModel.BottomCardFunctionScanList
        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "FrontInputsValue"
        Case "FrontInputsCaption"
            K2700Sheet.Range(m_inputsCell).Value = This.ViewModel.FrontInputsCaption
        Case "SelectedChannelNumber"
            K2700Sheet.Range(m_channelNumberCell).Value = This.ViewModel.SelectedChannelNumber
        Case "ChannelCount"
            K2700Sheet.Range(m_channelCountCell).Value = This.ViewModel.ChannelCount
        Case "CurrentChannelNumber"
        Case "MeasuredChannelNumber"
            Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
        Case "SenseFunctionName"
        Case "SenseFunctionCaption"
            K2700Sheet.Range(m_functionNameCell).Value = This.ViewModel.SenseFunctionCaption
        Case "SenseFunctionUnit"
            K2700Sheet.Range(m_functionUnitCell).Value = This.ViewModel.SenseFunctionUnit
        Case "MeasuredReading"
            K2700Sheet.Range(m_functionValueCell).Value = This.ViewModel.MeasuredReading
        Case "MeasuredValueFormat"
            K2700Sheet.Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "ExternalTrigMonitoringEnabled"
            Me.ExternalTrigMonitoringEnabled = This.ViewModel.ExternalTrigMonitoringEnabled
        Case "MeasuredValueFormat"
            K2700Sheet.Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        Case "AutoIncrementChannelNoEnabled"
            Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
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
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            K2700Sheet.Range(m_monitoringStateCell).Value = This.ViewModel.MonitoringState
        
        Case "OpenConnectionExecutable"
            Me.OpenConnectionButton.Enabled = This.ViewModel.OpenConnectionExecutable
        Case "CloseConnectionExecutable"
            Me.CloseConnectionButton.Enabled = This.ViewModel.CloseConnectionExecutable
        Case "ToggleConnectionExecutable"
            ' ConnectToggleButton.Enabled = This.ViewModel.ToggleConnectionExecutable
        Case "ResetKnownStateExecutable"
            Me.RSTButton.Enabled = This.ViewModel.ResetKnownStateExecutable
        Case "ClearExecutionStateExecutable"
            Me.CLSButton.Enabled = This.ViewModel.ClearExecutionStateExecutable
        Case "QueryExecutable"
            Me.QueryButton.Enabled = This.ViewModel.QueryExecutable
        Case "ReadExecutable"
            Me.ReadButton.Enabled = This.ViewModel.ReadExecutable
        Case "ReadStatusByteExecutable"
            Me.ReadStatusByteButton.Enabled = This.ViewModel.ReadStatusByteExecutable
        Case "ReadStandardEventStatusExecutable"
            Me.ReadStandardEventStatusButton.Enabled = This.ViewModel.ReadStandardEventStatusExecutable
        Case "ReadStatusAfterWriteExecutable"
            Me.ReadStatusAfterWriteCheckBox.Enabled = This.ViewModel.ReadStandardEventStatusExecutable
        Case "WriteExecutable"
            Me.WriteButton.Enabled = This.ViewModel.WriteExecutable
        Case "QueryInputsExecutable"
            Me.QueryInputsButton.Enabled = This.ViewModel.QueryInputsExecutable
        Case "ClearReadingsExecutable"
            ClearReadingsButton.Enabled = This.ViewModel.ClearReadingsExecutable
        Case "MeasureExecutable"
            MeasureButton.Enabled = This.ViewModel.MeasureExecutable
        Case "StopMonitoringExecutable"
            StopMonitoringButton.Enabled = This.ViewModel.StopMonitoringExecutable
        Case "StartMonitoringExecutable"
            StartMonitoringButton.Enabled = This.ViewModel.StartMonitoringExecutable
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
Private Sub OnWorksheetChange(ByVal a_range As Excel.Range)

    ' If Not Intersect(a_range, Me.K2700Sheet.Range("H5")) Is Nothing Then
    ' this uses fewer resources than Intersect,
    ' which will be helpful if your worksheet changes a lot.
    Select Case a_range.Address
        
        Case m_socketHostCell
            This.ViewModel.SocketAddress = K2700Sheet.Range(m_socketHostCell).Value & ":" & _
                K2700Sheet.Range(m_socketPortCell).Value
        Case m_socketPortCell
            This.ViewModel.SocketAddress = K2700Sheet.Range(m_socketHostCell).Value & ":" & _
                K2700Sheet.Range(m_socketPortCell).Value
        Case m_socketAddressCell
            K2700Observer.SocketAddress = Me.SocketAddress
        Case m_socketIdCell
            K2700Observer.SocketId = K2700Sheet.Range(m_socketIdCell).Value
        Case m_sessionTimeoutCell
            K2700Observer.SessionTimeout = Me.SessionTimeout
            
        Case m_connectedCell
            ' connected is read only
            ' This.ViewModel.Connected = K2700Sheet.Range(m_connectedCell).Value
        Case m_sentMessageCell
            This.ViewModel.SentMessage = K2700Sheet.Range(m_sentMessageCell).Value
        Case m_receivedMessageLengthCell
            This.ViewModel.ReceivedMessageLength = K2700Sheet.Range(m_receivedMessageLengthCell).Value
        Case m_receivedMessageCell
            This.ViewModel.ReceivedMessage = K2700Sheet.Range(m_receivedMessageCell).Value
        Case m_elapsedTimeCell
            This.ViewModel.ElapsedTimeCaption = K2700Sheet.Range(m_elapsedTimeCell).Value
        Case m_elapsedTimeFormatCell
            This.ViewModel.ElapsedTimeFormat = K2700Sheet.Range(m_elapsedTimeFormatCell).Value
        Case m_lastMessageCell
            This.ViewModel.LastMessage = K2700Sheet.Range(m_lastMessageCell).Value
        Case m_lastErrorMessageCell
            This.ViewModel.LastErrorMessage = K2700Sheet.Range(m_lastErrorMessageCell).Value
        Case m_serialPollCell
            This.ViewModel.SerialPollByte = K2700Sheet.Range(m_serialPollCell).Value
        Case m_statusByteCell
            This.ViewModel.StatusByte = K2700Sheet.Range(m_statusByteCell).Value
        Case m_standardByteCell
            This.ViewModel.StandardByte = K2700Sheet.Range(m_standardByteCell).Value

        Case m_srqCell
            This.ViewModel.ServiceRequested = K2700Sheet.Range(m_srqCell).Value
        Case m_gpibAddressCell
            This.ViewModel.GpibAddress = K2700Sheet.Range(m_gpibAddressCell).Value
        Case m_readTimeoutCell
            This.ViewModel.GpibLanReadTimeout = K2700Sheet.Range(m_readTimeoutCell).Value
        Case m_autoAssertTalkCell
            This.ViewModel.AutoAssertTalk = VBA.CBool(K2700Sheet.Range(m_autoAssertTalkCell).Value)
        Case m_timerIntervalCell
            This.ViewModel.TimerInterval = K2700Sheet.Range(m_timerIntervalCell).Value

        Case m_readingFormatCell
            This.ViewModel.MeasuredValueFormat = K2700Sheet.Range(m_readingFormatCell).Value
        Case m_card1Cell
            This.ViewModel.TopCard = K2700Sheet.Range(m_card1Cell).Value
        Case m_card2Cell
            This.ViewModel.TopCard = K2700Sheet.Range(m_card2Cell).Value
        Case m_card1ScanListCell
            This.ViewModel.TopCardFunctionScanList = K2700Sheet.Range(m_card1ScanListCell).Value
        Case m_card2ScanListCell
            This.ViewModel.BottomCardFunctionScanList = K2700Sheet.Range(m_card2ScanListCell).Value
        Case m_inputsCell
            This.ViewModel.FrontInputsCaption = K2700Sheet.Range(m_inputsCell).Value
        Case m_channelNumberCell
            This.ViewModel.SelectedChannelNumber = K2700Sheet.Range(m_channelNumberCell).Value
        Case m_channelCountCell
            This.ViewModel.ChannelCount = K2700Sheet.Range(m_channelCountCell).Value
        Case m_functionNameCell
            This.ViewModel.SenseFunctionCaption = K2700Sheet.Range(m_functionNameCell).Value
        Case m_functionValueCell
            This.ViewModel.MeasuredReading = MeasuredReading
        Case m_functionUnitCell
            This.ViewModel.SenseFunctionUnit = K2700Sheet.Range(m_functionUnitCell).Value
        Case m_monitoringStateCell
            This.ViewModel.MonitoringState = K2700Sheet.Range(m_monitoringStateCell).Value
        Case m_readingOffsetCell
            This.ViewModel.ReadingOffset = K2700Sheet.Range(m_readingOffsetCell).Value

    End Select

End Sub


''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub Worksheet_Change(ByVal a_range As Excel.Range)

    OnWorksheetChange a_range

End Sub

 
