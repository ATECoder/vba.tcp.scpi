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
Private Const m_timerIntervalGetCell As String = "C22"
Private Const m_timerIntervalSetCell As String = "D22"
Private Const m_topCardCell As String = "C23"
Private Const m_bottomCardCell As String = "C24"
Private Const m_topCardScanListCell As String = "C25"
Private Const m_bottomCardScanListCell As String = "C26"
Private Const m_inputsCell As String = "C27"
Private Const m_dutNumberCell As String = "C28"
Private Const m_targetDutNumberCell As String = "D28"
Private Const m_functionNameCell As String = "B29"
Private Const m_functionValueCell As String = "C29"
Private Const m_functionUnitCell As String = "D29"
Private Const m_functionUnicodeCell As String = "E29"
Private Const m_dutCountCell As String = "C30"
Private Const m_maximumDutNumberCell As String = "D30"
Private Const m_monitoringStateCell As String = "C31"
Private Const m_triggerModeCell As String = "C32"
Private Const m_MeasuredDutNumberCell As String = "E28"

Private Const m_readingOffsetCell As String = "L4"

Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 14 ' 2
Private Const m_dutNumberRow As Integer = 2 ' 7
Private Const m_dutNumberCol As Integer = 13 ' 1
Private Const m_byRow As Boolean = True ' False
Private Const m_displayIncrement As Integer = 1 '  2

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_

    ' view model settings
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    ViewModelOwner As Boolean
    
    FocusedControlName As String
    DesignMode As Boolean
    MeasuredValue As Double
    
    ' Initial User inputs
    GpibLanControllerPort As Long
    MaximumDutNumber As Integer
    
    ' IEEE488
    ScpiCommands As VBA.Collection
    SenseFunctions As VBA.Collection
    
    ' command info
    LastCommandOutcome As Boolean
    LastCommandDetails As String
    ElapsedTime As Double
    
    ' data acquisition
    Measuring As Boolean
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    FrontInputsValue As Boolean
    PauseRequested As Boolean
    StopRequested As Boolean
    
    MeasurementRow As Integer
    MeasurementColumn As Integer
    DutNumberRow As Integer
    DutNumberColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String
    DutNumberCaptionPrefix As String
    MeasureMode As cc_isr_Tcp_Scpi.MeasureMode
    
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Construction and cleanup
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize(ByVal a_viewModel As K2700ViewModel) As K2700Sheet
    
    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    This.ViewModelOwner = False
    
    Set This.ViewModel = a_viewModel
    
    ' set the initial values that are provided by the user interface
    Me.SocketAddress = Me.Host & ":" & VBA.CStr(Me.Port)
    Me.GpibLanControllerPort = 1234
    This.DutNumberCaptionPrefix = "dut"
    Me.MaximumDutNumber = 48
    Me.FrontInputsRequired = True
    Me.SingleReadEnabled = True
    
    Me.Synchronize
    
    ' set reference to the view model event handler.
    
    Set m_viewModel = a_viewModel
    
    Set Initialize = Me

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Initializes this object if not initialized. </summary>
Private Sub Initialize_()

    Me.Initialize cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
    This.ViewModelOwner = True
    
End Sub

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set m_viewModel = Nothing
    
    If This.ViewModelOwner And Not This.ViewModel Is Nothing Then
        If This.ViewModel.Connected Then
            This.ViewModel.Connectable.CloseConnection
        End If
        This.ViewModel.Dispose
    End If
    
    Set This.ViewModel = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Winsock Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the socket address. </summary>
''' <value>   [String]. </value
Public Property Get SocketAddress() As String
    If 0 = Len(Me.Range(m_socketAddressCell).Value) Then Me.Range(m_socketAddressCell).Value = "192.168.0.252:1234"
    SocketAddress = VBA.CStr(Me.Range(m_socketAddressCell).Value)
End Property

Public Property Let SocketAddress(ByVal a_value As String)
    If Me.SocketAddress <> a_value Then
        Me.Range(m_socketAddressCell).Value = a_value
    End If
End Property

''' <summary>   Gets or sets the session timeout. </summary>
''' <value>   [Long]. </value>
Public Property Get SessionTimeout() As Long
    If 0 = Len(Me.Range(m_sessionTimeoutCell).Value) Then Me.Range(m_sessionTimeoutCell).Value = 3000
    SessionTimeout = VBA.CLng(Me.Range(m_sessionTimeoutCell).Value)
End Property

Public Property Let SessionTimeout(ByVal a_value As Long)
    If Me.SessionTimeout <> a_value Then
        Me.Range(m_sessionTimeoutCell).Value = a_value
    End If
End Property

''' <summary>   Gets or sets the connected state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    If 0 = Len(Me.Range(m_connectedCell).Value) Then Me.Range(m_connectedCell).Value = False
    Connected = VBA.CBool(Me.Range(m_connectedCell).Value)
End Property

Public Property Let Connected(ByVal a_value As Boolean)
    If Me.Connected <> a_value Then
        Me.Range(m_connectedCell).Value = a_value
    End If
End Property

''' <summary>   Gets the INET Address dotted IP Address. </summary>
''' <value>   [String]. </value>
Public Property Get Host() As String
    Host = VBA.CStr(Me.Range(m_socketHostCell).Value)
End Property

''' <summary>   Sets the Host. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let Host(ByVal a_value As String)
    If Me.Host <> a_value Then
        Me.Range(m_socketHostCell).Value = a_value
    End If
End Property

''' <summary>   Gets the port number. </summary>
''' <value>   [Long]. </value>
Public Property Get Port() As Long
    If 0 = Len(Me.Range(m_socketPortCell).Value) Then Me.Range(m_socketPortCell).Value = 0
    Port = VBA.CLng(Me.Range(m_socketPortCell).Value)
End Property

''' <summary>   Sets the Port. </summary>
''' <param name="a_value">   [Long] the Port. </value>
Public Property Let Port(ByVal a_value As Long)
    If Me.Port <> a_value Then
        Me.Range(m_socketPortCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Socket Id. </summary>
''' <value>   [Long]. </value>
Public Property Get SocketId() As Long
    If 0 = Len(Me.Range(m_socketIdCell).Value) Then Me.Range(m_socketIdCell).Value = 0
    SocketId = VBA.CLng(Me.Range(m_socketIdCell).Value)
End Property

''' <summary>   Sets the Socket Id. </summary>
''' <param name="a_value">   [Long] the Socket Id. </value>
Public Property Let SocketId(ByVal a_value As Long)
    If Me.SocketId <> a_value Then
        Me.Range(m_socketIdCell).Value = a_value
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

Private Sub m_viewModel_MeasurementStateChanged(ByVal a_eventArgs As cc_isr_Tcp_Scpi.StateChangedEventArgs)
    
    Dim p_caption As String
    Select Case This.MeasurementMode
        Case cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous
            p_caption = "Continuous"
        Case cc_isr_Tcp_Scpi.MeasurementModeOption.External
            p_caption = "External"
        Case cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
            p_caption = "Immediate"
        Case cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring
            p_caption = "Monitoring"
        Case cc_isr_Tcp_Scpi.MeasurementModeOption.None
            p_caption = "None"
    End Select
    TriggerMode = p_caption
End Sub

''' <summary>   Handles open connection button click event. </summary>
Private Sub OpenConnectionButton_Click()
    ' initialize if necessary
    Initialize_
   
    This.ViewModel.OpenConnectionCommand Me.SocketAddress, Me.SessionTimeout
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

''' <summary>   Handles close connection button click event. </summary>
Private Sub CloseConnectionButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.CloseConnectionCommand
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Message to Sent. </summary>
''' <value>   [String]. </value>
Public Property Get MessageToSend() As String
    MessageToSend = Me.CommandsComboBox.Text
End Property

''' <summary>   Sets the Message to Sent. </summary>
''' <param name="a_value">   [String] the Message to Sent. </value>
Public Property Let MessageToSend(ByVal a_value As String)
    If Me.MessageToSend <> a_value Then
        Me.CommandsComboBox.Text = a_value
    End If
End Property

''' <summary>   Gets the Sent Message. </summary>
''' <value>   [String]. </value>
Public Property Get SentMessage() As String
    SentMessage = Me.Range(m_sentMessageCell).Value
End Property

''' <summary>   Sets the Sent Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SentMessage(ByVal a_value As String)
    If Me.SentMessage <> a_value Then
        Me.Range(m_sentMessageCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Received Message. </summary>
''' <value>   [String]. </value>
Public Property Get ReceivedMessage() As String
    ReceivedMessage = Me.Range(m_receivedMessageCell).Value
End Property

''' <summary>   Sets the Received Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let ReceivedMessage(ByVal a_value As String)
    If Me.ReceivedMessage <> a_value Then
        Me.Range(m_receivedMessageCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Received Message Length. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReceivedMessageLength() As Long
    If 0 = Len(Me.Range(m_receivedMessageLengthCell).Value) Then Me.Range(m_receivedMessageLengthCell).Value = 0
    ReceivedMessageLength = VBA.CLng(Me.Range(m_receivedMessageLengthCell).Value)
End Property

''' <summary>   Sets the Received Message Length. </summary>
''' <param name="a_value">   [Integer] the Message Length value. </value>
Public Property Let ReceivedMessageLength(ByVal a_value As Long)
    If Me.ReceivedMessageLength <> a_value Then
        Me.Range(m_receivedMessageLengthCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Status Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StatusByte() As Integer
    If 0 = Len(Me.Range(m_statusByteCell).Value) Then Me.Range(m_statusByteCell).Value = 0
    StatusByte = VBA.CInt(Me.Range(m_statusByteCell).Value)
End Property

''' <summary>   Sets the Status Byte. </summary>
''' <param name="a_value">   [Integer] the Status Byte value. </value>
Public Property Let StatusByte(ByVal a_value As Integer)
    If Me.StatusByte <> a_value Then
        Me.Range(m_statusByteCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    If 0 = Len(Me.Range(m_standardByteCell).Value) Then Me.Range(m_standardByteCell).Value = 0
    StandardByte = VBA.CInt(Me.Range(m_standardByteCell).Value)
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Public Property Let StandardByte(ByVal a_value As Integer)
    If Me.StandardByte <> a_value Then
        Me.Range(m_standardByteCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Auto Status Read Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadStatusAfterWriteEnabled() As Boolean
    ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
End Property

''' <summary>   Sets the Auto Status Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Status Read Enabled value. </value>
Public Property Let ReadStatusAfterWriteEnabled(ByVal a_value As Boolean)
    If Me.ReadStatusAfterWriteEnabled <> a_value Then
        Me.ReadStatusAfterWriteCheckBox.Value = a_value
    End If
End Property

''' <summary>   Gets the Scpi Commands. </summary>
''' <value>   [Collection]. </value>
Public Property Get ScpiCommands() As VBA.Collection
    Set ScpiCommands = This.ScpiCommands
End Property

''' <summary>   Sets the Scpi Commands. </summary>
''' <param name="a_value">   [Collection] the Scpi Commands value. </value>
Public Property Let ScpiCommands(ByVal a_value As VBA.Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.ScpiCommands, a_value) Then
        Set This.ScpiCommands = a_value
        Me.PopulateCommandsComboBox
    End If
End Property

''' <summary>   Gets the Sense Functions. </summary>
''' <value>   [Collection]. </value>
Public Property Get SenseFunctions() As Collection
    Set SenseFunctions = This.SenseFunctions
End Property

''' <summary>   Sets the Sense Functions. </summary>
''' <param name="a_value">   [Collection] the Sense Functions value. </value>
Public Property Let SenseFunctions(ByVal a_value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.SenseFunctions, a_value) Then
        Set This.SenseFunctions = a_value
        Me.PopulateSenseFunctionComboBox
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.LastCommandOutcome = This.ViewModel.ClearExecutionStateCommand(This.LastCommandDetails)
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
        ' If Not This.ViewModel Is Nothing Then _
        '     This.ViewModel.MessageToSend = Me.MessageToSend
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        ' If Not This.ViewModel Is Nothing Then _
        '     This.ViewModel.MessageToSend = Me.MessageToSend

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    ' If Not This.ViewModel Is Nothing Then _
    '     This.ViewModel.MessageToSend = Me.MessageToSend

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Public Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.ScpiCommands Is Nothing Then Exit Sub
    
    Me.CommandsComboBox.Clear
    Dim p_item As Variant
    Dim p_scpiCommand As String
    For Each p_item In This.ScpiCommands
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
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.QueryCommand Me.MessageToSend
End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()
    If Not This.ViewModel Is Nothing Then _
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
    If This.ViewModel Is Nothing Then Exit Sub
    If This.FocusedControlName = Me.ReadStatusAfterWriteCheckBox.Name Then
        On Error GoTo 0
        ' Commands for click
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
    Else
        On Error GoTo 0
        ' Commands for change from within the sheet
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    ' If Not This.ViewModel Is Nothing Then _
    '   This.ViewModel.ReadStatusAfterWriteEnabled = Me.ReadStatusAfterWriteCheckBox.Value
    
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Public Sub ReadStatusByteButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.QueryStatusByteCommand
End Sub

''' <summary>   Handles the Read Standard Event Status Byte button click event. </summary>
Private Sub ReadStandardEventStatusButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.ReadStandardEventStatusCommand
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.LastCommandOutcome = This.ViewModel.ResetKnownStateCommand(This.LastCommandDetails)
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.WriteCommand Me.MessageToSend, Me.ReadStatusAfterWriteEnabled
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE4888 Executables
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the condition indicating that the Reset Known State command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Reset Known State command can be executed. </value>
Public Property Get ResetKnownStateExecutable() As Boolean
    ResetKnownStateExecutable = Me.RSTButton.Enabled
End Property

Public Property Let ResetKnownStateExecutable(ByVal a_value As Boolean)
    Me.RSTButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Clear Execution State Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Clear Execution State Command can be executed. </value>
Public Property Get ClearExecutionStateExecutable() As Boolean
    ClearExecutionStateExecutable = Me.CLSButton.Enabled
End Property

Public Property Let ClearExecutionStateExecutable(ByVal a_value As Boolean)
    Me.CLSButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Query command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query command can be executed. </value>
Public Property Get QueryExecutable() As Boolean
    QueryExecutable = Me.QueryButton.Enabled
End Property

Public Property Let QueryExecutable(ByVal a_value As Boolean)
    Me.QueryButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Read command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read command can be executed. </value>
Public Property Get ReadExecutable() As Boolean
    ReadExecutable = Me.ReadButton.Enabled
End Property

Public Property Let ReadExecutable(ByVal a_value As Boolean)
    Me.ReadButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Write command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Write command can be executed. </value>
Public Property Get WriteExecutable() As Boolean
    WriteExecutable = Me.WriteButton.Enabled
End Property

Public Property Let WriteExecutable(ByVal a_value As Boolean)
    Me.WriteButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Query Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status Byte command can be executed. </value>
Public Property Get QueryStatusByteExecutable() As Boolean
    QueryStatusByteExecutable = Me.ReadStatusByteButton.Enabled
End Property

Public Property Let QueryStatusByteExecutable(ByVal a_value As Boolean)
    Me.ReadStatusByteButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Read Standard Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Standard Byte command can be executed. </value>
Public Property Get ReadStandardByteExecutable() As Boolean
    ReadStandardByteExecutable = Me.ReadStandardEventStatusButton.Enabled
End Property

Public Property Let ReadStandardByteExecutable(ByVal a_value As Boolean)
    Me.ReadStandardEventStatusButton.Enabled = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Auto Assert Talk status. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    If 0 = Len(Me.Range(m_autoAssertTalkCell).Value) Then Me.Range(m_autoAssertTalkCell).Value = False
    AutoAssertTalk = VBA.CBool(Me.Range(m_autoAssertTalkCell).Value)
End Property

''' <summary>   Sets the Auto Assert Talk status. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let AutoAssertTalk(ByVal a_value As Boolean)
    If Me.AutoAssertTalk <> a_value Then
        Me.Range(m_autoAssertTalkCell).Value = a_value
    End If
End Property

''' <summary>   Gets or sets the primary Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Integer
    If 0 = Len(Me.Range(m_gpibAddressCell).Value) Then Me.Range(m_gpibAddressCell).Value = 16
    PrimaryGpibAddress = VBA.CInt(Me.Range(m_gpibAddressCell).Value)
End Property

Public Property Let PrimaryGpibAddress(ByVal a_value As Integer)
    If Me.PrimaryGpibAddress <> a_value Then
        Me.Range(m_gpibAddressCell).Value = a_value
    End If
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    If This.GpibLanControllerPort <> a_value Then
        This.GpibLanControllerPort = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibLanControllerPort = a_value
    End If
End Property

''' <summary>   Gets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanReadTimeout() As Integer
    If 0 = Len(Me.Range(m_readTimeoutCell).Value) Then Me.Range(m_readTimeoutCell).Value = 3000
    GpibLanReadTimeout = VBA.CInt(Me.Range(m_readTimeoutCell).Value)
End Property

''' <summary>   Sets the GPIB-Lan Device timeout in milliseconds for reading from the
'''             instrument to the Gpib-Lan device using the <c>++read</c> command. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let GpibLanReadTimeout(ByVal a_value As Integer)
    If Me.GpibLanReadTimeout <> a_value Then
        Me.Range(m_readTimeoutCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Serial Poll Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get SerialPollByte() As Integer
    If 0 = Len(Me.Range(m_serialPollCell).Value) Then Me.Range(m_serialPollCell).Value = 0
    SerialPollByte = VBA.CInt(Me.Range(m_serialPollCell).Value)
End Property

''' <summary>   Sets the Serial Poll Byte. </summary>
''' <param name="a_value">   [Integer] the Serial Poll Byte value. </value>
Public Property Let SerialPollByte(ByVal a_value As Integer)
    If Me.SerialPollByte <> a_value Then
        Me.Range(m_serialPollCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    If 0 = Len(Me.Range(m_srqCell).Value) Then Me.Range(m_srqCell).Value = 0
    ServiceRequested = VBA.CInt(Me.Range(m_srqCell).Value)
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_value As Integer)
    If Me.ServiceRequested <> a_value Then
        Me.Range(m_srqCell).Value = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.GoToLocalCommand
End Sub

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GpibAddressGetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.GpibAddressGetCommand
End Sub

''' <summary>   Handles the GPIB Address set button click event. </summary>
Private Sub GpibAddressSetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.GpibAddressSetCommand Me.PrimaryGpibAddress
End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.LocalLockoutCommand
End Sub

''' <summary>   Handles the click event of the power on reset.  </summary>
Private Sub PowerOnResetButton_Click()
    ' initialize if necessary
    Initialize_
    This.ViewModel.PowerOnResetCommand
End Sub

''' <summary>   Handles the read after write (Auto Assert TALK) set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.AutoAssertTalkSetCommand Me.AutoAssertTalk
End Sub

''' <summary>   Handles the read after write (Auto Assert TALK) get button click event. </summary>
Private Sub ReadAfterWriteGetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.AutoAssertTalkGetCommand
End Sub

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.GpibLanReadTimeoutGetCommand
End Sub

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.GpibLanReadTimeoutSetCommand Me.GpibLanReadTimeout
End Sub

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.SelectiveDeviceClearCommand
End Sub

''' <summary>   Handles the serial poll button click event. </summary>
Private Sub SerialPollButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.ReadStatusByteCommand
End Sub

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.ServiceRequestedQueryCommand
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' GPIB-Lan Executables
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets or sets the condition indicating that the Auto Asset TALK get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK get command can be executed. </value>
Public Property Get AutoAssertTalkGetExecutable() As Boolean
    AutoAssertTalkGetExecutable = Me.ReadAfterWriteGetButton.Enabled
End Property

Public Property Let AutoAssertTalkGetExecutable(ByVal a_value As Boolean)
    Me.ReadAfterWriteGetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Auto Asset TALK Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Auto Asset TALK get command can be executed. </value>
Public Property Get AutoAssertTalkSetExecutable() As Boolean
    AutoAssertTalkSetExecutable = Me.ReadAfterWriteSetButton.Enabled
End Property

Public Property Let AutoAssertTalkSetExecutable(ByVal a_value As Boolean)
    Me.ReadAfterWriteSetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Gpib Address Get Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Get Command can be executed. </value>
Public Property Get GpibAddressGetExecutable() As Boolean
    GpibAddressGetExecutable = Me.GpibAddressGetButton.Enabled
End Property

Public Property Let GpibAddressGetExecutable(ByVal a_value As Boolean)
     Me.GpibAddressGetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Gpib Address Set Command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Gpib Address Set Command can be executed. </value>
Public Property Get GpibAddressSetExecutable() As Boolean
    GpibAddressSetExecutable = Me.GpibAddressSetButton.Enabled
End Property

Public Property Let GpibAddressSetExecutable(ByVal a_value As Boolean)
     Me.GpibAddressSetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Read Timeout Get command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Get command can be executed. </value>
Public Property Get GpibLanReadTimeoutGetExecutable() As Boolean
    GpibLanReadTimeoutGetExecutable = Me.ReadTimeoutGetButton.Enabled
End Property

Public Property Let GpibLanReadTimeoutGetExecutable(ByVal a_value As Boolean)
     Me.ReadTimeoutGetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Read Timeout Set command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Timeout Set command can be executed. </value>
Public Property Get GpibLanReadTimeoutSetExecutable() As Boolean
    GpibLanReadTimeoutSetExecutable = Me.ReadTimeoutSetButton.Enabled
End Property

Public Property Let GpibLanReadTimeoutSetExecutable(ByVal a_value As Boolean)
     Me.ReadTimeoutSetButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Service Requested command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Service Requested command can be executed. </value>
Public Property Get ServiceRequestedExecutable() As Boolean
    ServiceRequestedExecutable = Me.SRQButton.Enabled
End Property

Public Property Let ServiceRequestedExecutable(ByVal a_value As Boolean)
    Me.SRQButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Read Status Byte command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Read Status Byte command can be executed. </value>
Public Property Get ReadStatusByteExecutable() As Boolean
    ReadStatusByteExecutable = Me.SerialPollButton.Enabled
End Property

Public Property Let ReadStatusByteExecutable(ByVal a_value As Boolean)
    Me.SerialPollButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Go To Local command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Go To Local command can be executed. </value>
Public Property Get GoToLocalExecutable() As Boolean
    GoToLocalExecutable = Me.GoToLocalButton.Enabled
End Property

Public Property Let GoToLocalExecutable(ByVal a_value As Boolean)
    Me.GoToLocalButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Local Lockout command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Local Lockout command can be executed. </value>
Public Property Get LocalLockoutExecutable() As Boolean
    LocalLockoutExecutable = Me.LocalLockoutButton.Enabled
End Property

Public Property Let LocalLockoutExecutable(ByVal a_value As Boolean)
    Me.LocalLockoutButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the serial poll command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Serial Poll command can be executed. </value>
Public Property Get SerialPollExecutable() As Boolean
    SerialPollExecutable = Me.SerialPollButton.Enabled
End Property

Public Property Let SerialPollExecutable(ByVal a_value As Boolean)
    Me.SerialPollButton.Enabled = a_value
End Property

''' <summary>   Gets or sets  the condition indicating that the Selective Device Clear command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Selective Device Clear command can be executed. </value>
Public Property Get SelectiveDeviceClearExecutable() As Boolean
    SelectiveDeviceClearExecutable = Me.SDCButton.Enabled
End Property

Public Property Let SelectiveDeviceClearExecutable(ByVal a_value As Boolean)
    Me.SDCButton.Enabled = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary> Get or sets the timer interval to use for setting the actual timer interval. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerIntervalGet() As Integer
    If 0 = Len(Me.Range(m_timerIntervalGetCell).Value) Then Me.Range(m_timerIntervalGetCell).Value = 500
    TimerIntervalGet = VBA.CInt(Me.Range(m_timerIntervalGetCell).Value)
End Property

Public Property Let TimerIntervalGet(ByVal a_value As Integer)
    If Me.TimerIntervalGet <> a_value Then
        Me.Range(m_timerIntervalGetCell).Value = a_value
    End If
End Property

''' <summary> Get or sets the timer interval that was set. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerIntervalSet() As Integer
    If 0 = Len(Me.Range(m_timerIntervalSetCell).Value) Then Me.Range(m_timerIntervalSetCell).Value = 500
    TimerIntervalSet = VBA.CInt(Me.Range(m_timerIntervalSetCell).Value)
End Property

Public Property Let TimerIntervalSet(ByVal a_value As Integer)
    If Me.TimerIntervalSet <> a_value Then
        Me.Range(m_timerIntervalSetCell).Value = a_value
    End If
End Property

''' <summary>   Gets the DUT count. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutCount() As Integer
    If 0 = Len(Me.Range(m_dutCountCell).Value) Then Me.Range(m_dutCountCell).Value = 0
    DutCount = VBA.CInt(Me.Range(m_dutCountCell).Value)
End Property

''' <summary>   Sets the DUT count. </summary>
''' <param name="a_value">   [Integer] the DUT count value. </value>
Public Property Let DutCount(ByVal a_value As Integer)
    If Me.DutCount <> a_value Then
        Me.Range(m_dutCountCell).Value = a_value
    End If
End Property

''' <summary>   Gets the selected DUT number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedDutNumber() As Integer
    If 0 = Len(Me.Range(m_dutNumberCell).Value) Then Me.Range(m_dutNumberCell).Value = 1
    SelectedDutNumber = VBA.CInt(Me.Range(m_dutNumberCell).Value)
End Property

''' <summary>   Sets the selected DUT number. </summary>
''' <param name="a_value">   [Integer] the selected DUT number. </param>
Public Property Let SelectedDutNumber(ByVal a_value As Integer)
    Me.Range(m_dutNumberCell).Value = a_value
End Property

''' <summary>   Gets the target DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get TargetDutNumber() As Integer
    If 0 = Len(Me.Range(m_targetDutNumberCell).Value) Then Me.Range(m_targetDutNumberCell).Value = 0
    TargetDutNumber = VBA.CInt(Me.Range(m_targetDutNumberCell).Value)
End Property

''' <summary>   Sets the target DUT number. </summary>
''' <param name="a_value">   [Integer] the target DUT number. </param>
Public Property Let TargetDutNumber(ByVal a_value As Integer)
    If Me.TargetDutNumber <> a_value Then
        Me.Range(m_targetDutNumberCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Top Card name. </summary>
''' <value>   [String]. </value>
Public Property Get TopCard() As String
    TopCard = Me.Range(m_topCardCell).Value
End Property

''' <summary>   Sets the Top Card name. </summary>
''' <param name="a_value">   [String] the Top Card name. </value>
Public Property Let TopCard(ByVal a_value As String)
    If Me.TopCard <> a_value Then
        Me.Range(m_topCardCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Top Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get TopCardFunctionScanList() As String
    TopCardFunctionScanList = Me.Range(m_topCardCell).Value
End Property

''' <summary>   Sets the Top Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Top Card function scan list. </value>
Public Property Let TopCardFunctionScanList(ByVal a_value As String)
    If Me.TopCardFunctionScanList <> a_value Then
         Me.Range(m_topCardCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Bottom Card name. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCard() As String
    BottomCard = Me.Range(m_bottomCardCell).Value
End Property

''' <summary>   Sets the Bottom Card name. </summary>
''' <param name="a_value">   [String] the Bottom Card name. </value>
Public Property Let BottomCard(ByVal a_value As String)
    If Me.BottomCard <> a_value Then
        Me.Range(m_bottomCardCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Bottom Card Function Scan List. </summary>
''' <value>   [String]. </value>
Public Property Get BottomCardFunctionScanList() As String
    BottomCardFunctionScanList = Me.Range(m_bottomCardCell).Value
End Property

''' <summary>   Sets the Bottom Card Function Scan List. </summary>
''' <param name="a_value">   [String] the Bottom Card function scan list. </value>
Public Property Let BottomCardFunctionScanList(ByVal a_value As String)
    If Me.BottomCardFunctionScanList <> a_value Then
         Me.Range(m_bottomCardCell).Value = a_value
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
    If This.FrontInputsValue <> a_value Then
        This.FrontInputsValue = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Caption. </summary>
''' <remarks>   True if front inputs; otherwise, Rear inputs. </remarks>
''' <Caption>   [Boolean]. </Caption>
Public Property Get FrontInputsCaption() As String
    FrontInputsCaption = Me.Range(m_inputsCell).Value
End Property

''' <summary>   Sets the Front Inputs Caption. </summary>
''' <param name="a_value">   [String] the front inputs caption. </Caption>
Public Property Let FrontInputsCaption(ByVal a_value As String)
    If Me.FrontInputsCaption <> a_value Then
        Me.Range(m_inputsCell).Value = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Data Acquisition properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the maximum DUT number; equals to the muximum number of DUTs that can be tested. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutNumber() As Integer
    If 0 = Len(Me.Range(m_maximumDutNumberCell).Value) Then Me.Range(m_maximumDutNumberCell).Value = This.MaximumDutNumber
    MaximumDutNumber = VBA.CInt(Me.Range(m_maximumDutNumberCell).Value)
End Property

''' <summary>   Sets the Naximum DUT Number. </summary>
''' <param name="a_value">   [Integer] the maximum DUT number. </value>
Public Property Let MaximumDutNumber(ByVal a_value As Integer)
    If Me.MaximumDutNumber <> a_value Then
        Me.Range(m_maximumDutNumberCell).Value = a_value
    End If
End Property

''' <summary>   Gets the measuring state. True when measurements are collected. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Measuring() As Boolean
    Measuring = This.Measuring
End Property

''' <summary>   Sets the measuring state. True when measurements are collected. </summary>
''' <param name="a_value">   [Boolean] the measuring state; true when measuring is active. </value>
Public Property Let Measuring(ByVal a_value As Boolean)
    If This.Measuring <> a_value Then
        This.Measuring = a_value
    End If
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get MeasurementMode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    MeasurementMode = This.MeasurementMode
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Let MeasurementMode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    If This.MeasurementMode <> a_value Then
        This.MeasurementMode = a_value
    End If
End Property

''' <summary>   Gets the trigger mode. </summary>
''' <value>   [String]. </value>
Public Property Get TriggerMode() As String
    TriggerMode = Me.Range(m_triggerModeCell).Value
End Property

''' <summary>   Gets the trigger mode. </summary>
''' <param name="a_value">   [String] the trigger mode. </value>
Public Property Let TriggerMode(ByVal a_value As String)
    If Me.TriggerMode <> a_value Then
        Me.Range(m_triggerModeCell).Value = a_value
    End If
End Property


''' <summary>   Gets the Monitoring State. </summary>
''' <value>   [String]. </value>
Public Property Get MonitoringState() As String
    MonitoringState = Me.Range(m_monitoringStateCell).Value
End Property

''' <summary>   Sets the Monitoring State. </summary>
''' <param name="a_value">   [String] the Monitoring State value. </value>
Public Property Let MonitoringState(ByVal a_value As String)
    If Me.MonitoringState <> a_value Then
        Me.Range(m_monitoringStateCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Pause Requested condition. </summary>
''' <remarks>   When true, a pause request is sent to the timer so that the timer
'''             event is not handled and the timer is paused on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get PauseRequested() As Boolean
    PauseRequested = This.PauseRequested
End Property

''' <summary>   Sets the Pause Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Pause Requested value. </value>
Friend Property Let PauseRequested(ByVal a_value As Boolean)
    If This.PauseRequested <> a_value Then
        This.PauseRequested = a_value
    End If
End Property

''' <summary>   Gets the Stop Requested condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopd on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get StopRequested() As Boolean
    StopRequested = This.StopRequested
End Property

''' <summary>   Sets the Stop Requested condition. </summary>
''' <param name="a_value">   [Boolean] the Stop Requested value. </value>
Friend Property Let StopRequested(ByVal a_value As Boolean)
    If This.StopRequested <> a_value Then _
        This.StopRequested = a_value
End Property

''' <summary>   Gets the Sense Function Name. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = Me.SenseFunctionComboBox.Text
End Property

''' <summary>   Sets the Sense Function Name. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let SenseFunctionName(ByVal a_value As String)
    If Me.SenseFunctionName <> a_value Then _
        Me.SenseFunctionComboBox.Text = a_value
End Property

''' <summary>   Gets the  Sense Function Caption. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionCaption() As String
    SenseFunctionCaption = Me.Range(m_functionNameCell).Value
End Property

''' <summary>   Sets the  Sense Function Caption. </summary>
''' <param name="a_value">   [String] the  Sense Function Caption. </value>
Public Property Let SenseFunctionCaption(ByVal a_value As String)
    If Me.SenseFunctionCaption <> a_value Then
        Me.Range(m_functionNameCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Sense Function Unit. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunctionUnit() As String
    SenseFunctionUnit = Me.Range(m_functionUnitCell).Value
End Property

''' <summary>   Gets the Sense Function Unit Unicode value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SenseFunctionUnicode() As Integer
    SenseFunctionUnicode = Me.Range(m_functionUnicodeCell).Value
End Property

''' <summary>   Sets the Sense Function Unit Unicode value. </summary>
''' <param name="a_value">   [Integer] the Sense Function Unit value. </value>
Public Property Let SenseFunctionUnicode(ByVal a_value As Integer)
    If SenseFunctionUnicode <> a_value & a_value > 0 Then
        Me.Range(m_functionUnicodeCell).Value = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SenseFunctionUnit = StrConv(Me.SenseFunctionUnit, vbFromUnicode)
    End If
End Property

''' <summary>   Gets the measured DUT number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredDutNumber() As Integer
    If 0 = Len(Me.Range(m_MeasuredDutNumberCell).Value) Then Me.Range(m_MeasuredDutNumberCell).Value = -1
    MeasuredDutNumber = VBA.CInt(Me.Range(m_MeasuredDutNumberCell).Value)
End Property

''' <summary>   Sets the measured DUT number. </summary>
''' <param name="a_value">   [Integer] the measured DUT number. </param>
Public Property Let MeasuredDutNumber(ByVal a_value As Integer)
    If Me.MeasuredDutNumber <> a_value Then
        Me.Range(m_MeasuredDutNumberCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Measured reading. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = Me.Range(m_functionValueCell).Value
End Property

''' <summary>   Sets the Measured reading. </summary>
''' <param name="a_value">   [String] the Measured reading. </param>
Public Property Let MeasuredReading(ByVal a_value As String)
    If Me.MeasuredReading <> a_value Then
        Me.Range(m_functionValueCell).Value = a_value
    End If
End Property

''' <summary>   Gets the reading offset. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadingOffset() As Double
    If 0 = Len(Me.Range(m_readingOffsetCell).Value) Then Me.Range(m_readingOffsetCell).Value = 0#
    ReadingOffset = VBA.CDbl(Me.Range(m_readingOffsetCell).Value)
End Property

''' <summary>   Sets the reading offset. </summary>
''' <param name="a_value">   [Double] the reading offset. </param>
Public Property Let ReadingOffset(ByVal a_value As Double)
    If Me.ReadingOffset <> a_value Then
        Me.Range(m_readingOffsetCell).Value = a_value
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
    If This.MeasuredValue <> a_value Then
        This.MeasuredValue = a_value
    End If
End Property

''' <summary>   Gets the Measured Value Format. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredValueFormat() As String
    MeasuredValueFormat = Me.Range(m_readingFormatCell).Value
End Property

''' <summary>   Sets the Measured Value Format. </summary>
''' <param name="a_value">   [String] the Measured Value Format value. </value>
Public Property Let MeasuredValueFormat(ByVal a_value As String)
    If Me.MeasuredValueFormat <> a_value Then
        Me.Range(m_readingFormatCell).Value = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the auto increment DUT number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementDutNumberEnabled() As Boolean
    AutoIncrementDutNumberEnabled = Me.AutoIncrementCheckBox.Value
End Property

''' <summary>   Sets the auto increment DUT number Enabled. </summary>
''' <param name="a_value">   [Boolean] the auto increment DUT number Enabled value. </value>
Public Property Let AutoIncrementDutNumberEnabled(ByVal a_value As Boolean)
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
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.AutoIncrementDutNumberEnabled = Me.AutoIncrementCheckBox.Value
    Else
        On Error GoTo 0
        ' Commands for change from within the sheet
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.AutoIncrementDutNumberEnabled = Me.AutoIncrementCheckBox.Value
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    ' If Not This.ViewModel Is Nothing Then _
    '   This.ViewModel.AutoIncrementDutNumberEnabled = Me.AutoIncrementCheckBox.Value
    
End Sub

''' <summary>   Handles the click event of the Clear Radings button. </summary>
Private Sub ClearReadingsButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.ClearReadingsCommand Me.ReadingOffset, 1, Me.MaximumDutNumber
End Sub

''' <summary>   Handles the external trigger option button affirmative action event. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function OnExternalTriggerOption(ByRef a_details As String) As Boolean

    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.AutoIncrement = Me.AutoIncrementDutNumberEnabled
    p_mode.BeepEnabled = False
    p_mode.FrontInputs = Me.FrontInputsRequired
    p_mode.DutNumber = Me.SelectedDutNumber
    p_mode.MaximumDutCount = Me.MaximumDutNumber
    p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
    p_mode.ReadingOffset = Me.ReadingOffset
    p_mode.SenseFunction = Me.SenseFunctionComboBox.Text
    p_mode.SingleRead = Me.SingleReadEnabled
    p_mode.TimerInterval = Me.TimerIntervalGet
    Set This.MeasureMode = p_mode
    OnExternalTriggerOption = This.ViewModel.ConfigureMeasureCommand(p_mode, a_details)

End Function

''' <summary>   Handles the external trigger button click event. </summary>
Private Sub ExternalTriggerButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.LastCommandOutcome = OnExternalTriggerOption(This.LastCommandDetails)
End Sub

''' <summary>   Handles the immediate trigger option button affirmative action event. </summary>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Boolean]. </returns>
Public Function OnImmediateTriggerOption(ByRef a_details As String) As Boolean

    Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
    Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
    p_mode.AutoIncrement = Me.AutoIncrementDutNumberEnabled
    p_mode.BeepEnabled = False
    p_mode.FrontInputs = Me.FrontInputsRequired
    p_mode.DutNumber = Me.SelectedDutNumber
    p_mode.MaximumDutCount = Me.MaximumDutNumber
    p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
    p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
    p_mode.ReadingOffset = Me.ReadingOffset
    p_mode.SenseFunction = Me.SenseFunctionComboBox.Text
    p_mode.SingleRead = Me.SingleReadEnabled
    p_mode.TimerInterval = Me.TimerIntervalGet
    Set This.MeasureMode = p_mode
    OnImmediateTriggerOption = This.ViewModel.ConfigureMeasureCommand(p_mode, a_details)

End Function


''' <summary>   Configure immediate trigger mode and enables the read button. </summary>
Private Sub ImmediateTriggerButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.LastCommandOutcome = OnImmediateTriggerOption(This.LastCommandDetails)
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
    End If
End Property

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()
    ' make an immediate measure of one or all channels.
    If Not This.ViewModel Is Nothing Then _
        This.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(This.LastCommandDetails)
End Sub

''' <summary>   Handles the start monitoring button click event. </summary>
Private Sub StartMonitoringButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.StartMonitoringExternalTriggersCommand
End Sub

''' <summary>   Handles the stop monitoring button click event. </summary>
Private Sub StopMonitoringButton_Click()
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.StopMonitoringExternalTriggersCommand
End Sub

''' <summary> Queries the Inputs status. </summary>
Private Sub QueryInputsButton_Click()
    If Not This.ViewModel Is Nothing Then _
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
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
    Else
        On Error GoTo 0
        ' Commands for change from within the sheet
        ' If Not This.ViewModel Is Nothing Then _
        '   This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    ' If Not This.ViewModel Is Nothing Then _
    '   This.ViewModel.SingleReadEnabled = Me.SingleMultipleCheckBox.Value
    
End Sub

''' <summary>   Populates the Function combo box. </summary>
Public Sub PopulateSenseFunctionComboBox()

    Const p_procedureName = "PopulateSenseFunctionComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.SenseFunctions Is Nothing Then Exit Sub

    Dim p_fourWireResistanceIndex As Integer
    Me.SenseFunctionComboBox.Clear
    Dim p_item As Variant
    Dim p_senseFunction As String
    For Each p_item In This.SenseFunctions
        p_senseFunction = VBA.CStr(p_item)
        Me.SenseFunctionComboBox.AddItem p_senseFunction
        If "FRES" = p_senseFunction Then
            p_fourWireResistanceIndex = Me.SenseFunctionComboBox.ListCount
        End If
    Next

    Me.SenseFunctionComboBox.ListIndex = p_fourWireResistanceIndex
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' K2700 executables
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

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

''' <summary>   Gets or sets the condition indicating that the Query Inputs command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Query Inputs command can be executed. </value>
Public Property Get QueryInputsExecutable() As Boolean
    QueryInputsExecutable = Me.QueryInputsButton.Enabled
End Property

Public Property Let QueryInputsExecutable(ByVal a_value As Boolean)
     Me.QueryInputsButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Measure command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Measure command can be executed. </value>
Public Property Get MeasureExecutable() As Boolean
    MeasureExecutable = Me.MeasureButton.Enabled
End Property

Public Property Let MeasureExecutable(ByVal a_value As Boolean)
     Me.MeasureButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Immediate Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Immediate Trigger Option command  can be executed. </value>
Public Property Get ImmediateTriggerOptionExecutable() As Boolean
    ImmediateTriggerOptionExecutable = Me.ImmediateTriggerButton.Enabled
End Property

Public Property Let ImmediateTriggerOptionExecutable(ByVal a_value As Boolean)
     Me.ImmediateTriggerButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the External Trigger Option command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the External Trigger Option command  can be executed. </value>
Public Property Get ExternalTriggerOptionExecutable() As Boolean
    ExternalTriggerOptionExecutable = Me.ExternalTriggerButton.Enabled
End Property

Public Property Let ExternalTriggerOptionExecutable(ByVal a_value As Boolean)
     Me.ExternalTriggerButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the function name command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the function name command can be executed. </value>
Public Property Get FunctionNameExecutable() As Boolean
    FunctionNameExecutable = Me.SenseFunctionComboBox.Enabled
End Property

Public Property Let FunctionNameExecutable(ByVal a_value As Boolean)
     Me.SenseFunctionComboBox.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Start Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Start Monitoring External Trigger command can be executed. </value>
Public Property Get StartMonitoringExecutable() As Boolean
    StartMonitoringExecutable = Me.StartMonitoringButton.Enabled
End Property

Public Property Let StartMonitoringExecutable(ByVal a_value As Boolean)
     Me.StartMonitoringButton.Enabled = a_value
End Property

''' <summary>   Gets or sets the condition indicating that the Stop Monitoring External Trigger command
'''             can be executed. </summary>
''' <value>   [Boolean] True if the Stop Monitoring External Trigger command can be executed. </value>
Public Property Get StopMonitoringExecutable() As Boolean
    StopMonitoringExecutable = Me.StopMonitoringButton.Enabled
End Property

Public Property Let StopMonitoringExecutable(ByVal a_value As Boolean)
     Me.StopMonitoringButton.Enabled = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' View Model Settings
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the design mode for unit testing and restoration of state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get DesignMode() As Boolean
    DesignMode = This.DesignMode
End Property

''' <summary>   Sets the design mode for unit testing and restoration of state. </summary>
''' <param name="a_value">   [Boolean] true if in design mode. </param>
Public Property Let DesignMode(ByVal a_value As Boolean)
    This.DesignMode = a_value
End Property

''' <summary>   Gets the Elapsed Time in milliseconds. </summary>
''' <value>   [Double]. </value>
Public Property Get ElapsedTime() As Double
    ElapsedTime = This.ElapsedTime
End Property

''' <summary>   Sets the Elapsed Time in milliseconds. </summary>
''' <param name="a_value">   [Double] the Elapsed Time in milliseconds. </value>
Public Property Let ElapsedTime(ByVal a_value As Double)
    If Me.ElapsedTime <> a_value Then
        This.ElapsedTime = a_value
    End If
End Property

''' <summary>   Gets the Elapsed Time Format. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeFormat() As String
    ElapsedTimeFormat = Me.Range(m_elapsedTimeFormatCell).Value
End Property

''' <summary>   Sets the Elapsed Time Format. </summary>
''' <param name="a_value">   [String] the elapsed time format. </value>
Public Property Let ElapsedTimeFormat(ByVal a_value As String)
    If Me.ElapsedTimeFormat <> a_value Then
        Me.Range(m_elapsedTimeFormatCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Elapsed Time Caption. </summary>
''' <value>   [String]. </value>
Public Property Get ElapsedTimeCaption() As String
    ElapsedTimeCaption = Me.Range(m_elapsedTimeCell).Value
End Property

''' <summary>   Sets the Elapsed Time Caption. </summary>
''' <param name="a_value">   [String] the elapsed time Caption. </value>
Public Property Let ElapsedTimeCaption(ByVal a_value As String)
    If Me.ElapsedTimeCaption <> a_value Then
        Me.Range(m_elapsedTimeCell).Value = a_value
    End If
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = Me.Range(m_lastErrorMessageCell).Value
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    Me.Range(m_lastErrorMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = Me.Range(m_lastMessageCell).Value
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastMessage(ByVal a_value As String)
    Me.Range(m_lastMessageCell).Value = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Measurments View
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the DUT number caption prefix. </summary>
''' <value>   [String]. </value>
Public Property Get DutNumberCaptionPrefix() As String
    DutNumberCaptionPrefix = This.DutNumberCaptionPrefix
End Property

''' <summary>   Sets the DUT number caption prefix. </summary>
''' <param name="a_value">   [String] the DUT number caption prefix. </value>
Public Property Let DutNumberCaptionPrefix(ByVal a_value As String)
    This.DutNumberCaptionPrefix = a_value
End Property

''' <summary>   Gets the Measurement Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementRow() As Integer
    MeasurementRow = This.MeasurementRow
End Property

''' <summary>   Sets the Measurement Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Row. </value>
Public Property Let MeasurementRow(ByVal a_value As Integer)
    If This.MeasurementRow <> a_value Then
        This.MeasurementRow = a_value
    End If
End Property

''' <summary>   Gets the Measurement Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementColumn() As Integer
    MeasurementColumn = This.MeasurementColumn
End Property

''' <summary>   Sets the Measurement Column. </summary>
''' <param name="a_value">   [Integer] the Measurement Column. </value>
Public Property Let MeasurementColumn(ByVal a_value As Integer)
    If This.MeasurementColumn <> a_value Then
        This.MeasurementColumn = a_value
    End If
End Property

''' <summary>   Gets the DUT number row. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumberRow() As Integer
    DutNumberRow = This.DutNumberRow
End Property

''' <summary>   Sets the DUT number row. </summary>
''' <param name="a_value">   [Integer] the DUT number row. </value>
Public Property Let DutNumberRow(ByVal a_value As Integer)
    If This.DutNumberRow <> a_value Then
        This.DutNumberRow = a_value
    End If
End Property

''' <summary>   Gets the DUT number Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumberColumn() As Integer
    DutNumberColumn = This.DutNumberColumn
End Property

''' <summary>   Sets the DUT number Column. </summary>
''' <param name="a_value">   [Integer] the DUT number Column. </value>
Public Property Let DutNumberColumn(ByVal a_value As Integer)
    If This.DutNumberColumn <> a_value Then
        This.DutNumberColumn = a_value
    End If
End Property

''' <summary>   Gets the Measurement Cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementCellValue() As String
    MeasurementCellValue = Me.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If Me.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        Me.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
    End If
End Property
        
''' <summary>   Gets the channel cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCellValue() As String
    ChannelCellValue = Me.Cells(Me.DutNumberRow, Me.DutNumberColumn).Value
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If Me.ChannelCellValue <> a_value Then
        This.ChannelCellValue = a_value
        Me.Cells(Me.DutNumberRow, Me.DutNumberColumn).Value = a_value
    End If
End Property

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_dutNumber">   [single] the DUT number. </param>
''' <param name="a_value">       [double] the value. </param>
Public Sub SetReading(ByVal a_dutNumber As Integer, ByVal a_value As Double)
    SetReading_ a_dutNumber, a_value, m_byRow, m_displayIncrement
End Sub

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_dutNumber">   [single] the DUT number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_byRow">       [Boolean] true if the next reading goes to the next row. </param>
''' <param name="a_increment">   [Integer] the row or column increment. </param>
Private Sub SetReading_(ByVal a_dutNumber As Integer, ByVal a_value As Double, _
    ByVal a_byRow As Boolean, ByVal a_increment As Integer)

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_dutNumber - 1), _
        m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, _
        m_measurementCol + a_increment * (a_dutNumber - 1))
   
    Me.DutNumberRow = IIf(a_byRow, m_dutNumberRow + a_increment * (a_dutNumber - 1), _
        m_dutNumberRow)
    Me.DutNumberColumn = IIf(a_byRow, m_dutNumberCol, _
        m_dutNumberCol + a_increment * (a_dutNumber - 1))
   
    Me.ChannelCellValue = This.DutNumberCaptionPrefix & VBA.CStr(a_dutNumber)
    Me.MeasurementCellValue = VBA.CStr(a_value)
    
End Sub

Public Function GetDutNumberCaptionRow(ByVal a_dutNumber As Integer) As Integer
    GetDutNumberCaptionRow = IIf(m_byRow, m_dutNumberRow + m_displayIncrement * (a_dutNumber - 1), _
        m_dutNumberRow)
End Function

Public Function GetDutNumberCaptionColum(ByVal a_dutNumber As Integer) As Integer
    GetDutNumberCaptionColum = IIf(m_byRow, m_dutNumberCol, _
        m_dutNumberCol + m_displayIncrement * (a_dutNumber - 1))
End Function

Public Function GetMeasurementRow(ByVal a_dutNumber As Integer) As Integer
    GetMeasurementRow = IIf(m_byRow, m_dutNumberRow + m_displayIncrement * (a_dutNumber - 1), _
        m_dutNumberRow)
End Function

Public Function GetMeasurementColum(ByVal a_dutNumber As Integer) As Integer
    GetMeasurementColum = IIf(m_byRow, m_measurementCol, _
        m_measurementCol + m_displayIncrement * (a_dutNumber - 1))
End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' View Model Events
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.DutMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.DutReadingEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_DutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

    Const p_procedureName = "m_viewModel_DutMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' these two are populated by the property changed event.
    ' Me.MeasuredReading = a_eventArgs.Reading
    ' Me.MeasuredDutNumber = a_eventArgs.channelnumber
    
    ' this is done in the view model
    ' Dim p_value As Double
    ' Dim p_details As String
    ' If VBA.vbNullString = a_eventArgs.Reading Then
    '     p_value = cc_isr_Ieee488.Syntax.NotANumber
    ' Else
    '   If Not cc_isr_Core.StringExtensions.TryParseDouble(a_eventArgs.Reading, p_value, p_details) Then
    '         p_value = cc_isr_Ieee488.Syntax.NotANumber
    '     End If
    ' End If
    
    ' this is set by the property change event.
    ' Me.MeasuredValue = p_value
    
    Me.SetReading Me.MeasuredDutNumber, Me.MeasuredValue
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Synchronizes the sheet values with the view model. </summary>
Public Sub Synchronize()

    Const p_procedureName = "Synchronize"

    This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort
    
    If This.ViewModel.FrontInputsHasValue Then
        Me.FrontInputsValue = This.ViewModel.FrontInputsValue
        Me.FrontInputsCaption = This.ViewModel.FrontInputsCaption
    End If
    
    If This.ViewModel.GpibAddress > 0 Then
        Me.PrimaryGpibAddress = This.ViewModel.GpibAddress
    End If
    
    If This.ViewModel.HasValidSocketAddress Then
        Me.SocketAddress = This.ViewModel.SocketAddress
        Me.Host = This.ViewModel.Host
        Me.Port = This.ViewModel.Port
    End If
        
    If This.ViewModel.ReadingOffset > 0 Then _
        Me.ReadingOffset = This.ViewModel.ReadingOffset
    
    Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
    Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
    Me.BottomCard = This.ViewModel.BottomCard
    Me.BottomCardFunctionScanList = This.ViewModel.BottomCardFunctionScanList
    Me.DutCount = This.ViewModel.DutCount
    Me.Connected = This.ViewModel.Connected
    Me.ElapsedTime = This.ViewModel.ElapsedTime
    Me.ElapsedTimeFormat = This.ViewModel.ElapsedTimeFormat
    Me.ElapsedTimeCaption = This.ViewModel.ElapsedTimeCaption
    Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
    Me.LastErrorMessage = This.ViewModel.LastErrorMessage
    Me.LastMessage = This.ViewModel.LastMessage
    Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.MonitoringState = This.ViewModel.MonitoringState
    Me.PauseRequested = This.ViewModel.PauseRequested
    Me.ReceivedMessage = This.ViewModel.ReceivedMessage
    Me.ReceivedMessageLength = This.ViewModel.ReceivedMessageLength
    Me.ScpiCommands = This.ViewModel.ScpiCommands
    Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
    Me.SenseFunctionName = This.ViewModel.SenseFunctionName
    Me.SenseFunctionCaption = This.ViewModel.SenseFunctionCaption
    Me.SenseFunctions = This.ViewModel.SenseFunctions
    Me.SenseFunctionUnicode = This.ViewModel.SenseFunctionUnicode
    Me.SentMessage = This.ViewModel.SentMessage
    Me.SerialPollByte = This.ViewModel.SerialPollByte
    Me.ServiceRequested = This.ViewModel.ServiceRequested
    Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
    Me.SocketId = This.ViewModel.SocketId
    Me.StandardByte = This.ViewModel.StandardByte
    Me.StatusByte = This.ViewModel.StatusByte
    Me.StopRequested = This.ViewModel.StopRequested
    Me.TargetDutNumber = This.ViewModel.TargetDutNumber
    Me.TopCard = This.ViewModel.TopCard
    Me.TopCardFunctionScanList = This.ViewModel.TopCardFunctionScanList
    
    Me.AutoAssertTalkSetExecutable = This.ViewModel.AutoAssertTalkSetExecutable
    Me.AutoAssertTalkGetExecutable = This.ViewModel.AutoAssertTalkGetExecutable
    Me.ClearExecutionStateExecutable = This.ViewModel.ClearExecutionStateExecutable
    Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
    Me.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
    Me.ExternalTriggerOptionExecutable = This.ViewModel.ExternalTriggerOptionExecutable
    Me.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable
    Me.GoToLocalExecutable = This.ViewModel.GoToLocalExecutable
    Me.GpibAddressGetExecutable = This.ViewModel.GpibAddressGetExecutable
    Me.GpibAddressSetExecutable = This.ViewModel.GpibAddressSetExecutable
    Me.GpibLanReadTimeoutGetExecutable = This.ViewModel.GpibLanReadTimeoutGetExecutable
    Me.GpibLanReadTimeoutSetExecutable = This.ViewModel.GpibLanReadTimeoutSetExecutable
    Me.ImmediateTriggerOptionExecutable = This.ViewModel.ImmediateTriggerOptionExecutable
    Me.LocalLockoutExecutable = This.ViewModel.LocalLockoutExecutable
    Me.MeasureExecutable = This.ViewModel.MeasureExecutable
    Me.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
    Me.QueryExecutable = This.ViewModel.QueryExecutable
    Me.QueryInputsExecutable = This.ViewModel.QueryInputsExecutable
    Me.QueryStatusByteExecutable = This.ViewModel.QueryStatusByteExecutable
    Me.ReadExecutable = This.ViewModel.ReadExecutable
    Me.ReadStandardByteExecutable = This.ViewModel.ReadStandardEventStatusExecutable
    Me.ReadStatusByteExecutable = This.ViewModel.ReadStatusByteExecutable
    Me.ResetKnownStateExecutable = This.ViewModel.ResetKnownStateExecutable
    Me.SelectiveDeviceClearExecutable = This.ViewModel.SelectiveDeviceClearExecutable
    Me.SerialPollExecutable = This.ViewModel.ReadStatusByteExecutable
    Me.ServiceRequestedExecutable = This.ViewModel.ServiceRequestedExecutable
    Me.StartMonitoringExecutable = This.ViewModel.StartMonitoringExecutable
    Me.StopMonitoringExecutable = This.ViewModel.StopMonitoringExecutable
    Me.WriteExecutable = This.ViewModel.WriteExecutable
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Friend Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "HandleViewModelPropertyChanged"

    Select Case a_propertyName

        Case "Host"
            Me.Host = This.ViewModel.Host
        Case "Port"
            Me.Port = This.ViewModel.Port
        Case "SocketAddress"
            Me.SocketAddress = This.ViewModel.SocketAddress
        Case "SocketId"
            Me.SocketId = This.ViewModel.SocketId
        Case "SessionTimeout"
            Me.SessionTimeout = This.ViewModel.SessionTimeout
        Case "PauseRequested"
            Me.PauseRequested = This.ViewModel.PauseRequested
        Case "StopRequested"
            Me.StopRequested = This.ViewModel.StopRequested
        Case "TimerInterval"
            Me.TimerIntervalSet = This.ViewModel.TimerInterval
        Case "Connected"
            Me.Connected = This.ViewModel.Connected
        Case "SentMessage"
            Me.SentMessage = This.ViewModel.SentMessage
        Case "ReceivedMessage"
            Me.ReceivedMessage = This.ViewModel.ReceivedMessage
        Case "ReceivedMessageLength"
            Me.ReceivedMessageLength = This.ViewModel.ReceivedMessageLength
        Case "ElapsedTime"
            Me.ElapsedTime = This.ViewModel.ElapsedTime
        Case "ElapsedTimeFormat"
            Me.ElapsedTimeFormat = This.ViewModel.ElapsedTimeFormat
        Case "ElapsedTimeCaption"
            Me.ElapsedTimeCaption = This.ViewModel.ElapsedTimeCaption
        Case "LastErrorMessage"
            Me.LastErrorMessage = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            Me.LastMessage = This.ViewModel.LastMessage
        Case "SerialPollByte"
            Me.SerialPollByte = This.ViewModel.SerialPollByte
        Case "StatusByte"
            Me.StatusByte = This.ViewModel.StatusByte
        Case "StandardByte"
            Me.StandardByte = This.ViewModel.StandardByte
        Case "ServiceRequested"
            Me.ServiceRequested = This.ViewModel.ServiceRequested
        Case "GpibAddress"
            Me.PrimaryGpibAddress = This.ViewModel.GpibAddress
        Case "AutoAssertTalk"
            Me.AutoAssertTalk = This.ViewModel.AutoAssertTalk
        Case "GpibLanReadTimeout"
            Me.GpibLanReadTimeout = This.ViewModel.GpibLanReadTimeout
            
        Case "TopCard"
            Me.TopCard = This.ViewModel.TopCard
        Case "BottomCard"
            Me.BottomCard = This.ViewModel.BottomCard
        Case "TopCardFunctionScanList"
            Me.TopCardFunctionScanList = This.ViewModel.TopCardFunctionScanList
        Case "BottomCardFunctionScanList"
            Me.BottomCardFunctionScanList = This.ViewModel.BottomCardFunctionScanList
        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "FrontInputsValue"
            Me.FrontInputsValue = This.ViewModel.FrontInputsValue
        Case "FrontInputsCaption"
            Me.FrontInputsCaption = This.ViewModel.FrontInputsCaption
        Case "SelectedDutNumber"
            Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
        Case "DutCount"
            Me.DutCount = This.ViewModel.DutCount
        Case "TargetDutNumber"
            Me.TargetDutNumber = This.ViewModel.TargetDutNumber
        Case "MeasuredDutNumber"
            Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
        Case "SenseFunctionName"
            Me.SenseFunctionName = This.ViewModel.SenseFunctionName
        Case "SenseFunctionCaption"
            Me.SenseFunctionCaption = This.ViewModel.SenseFunctionCaption
        Case "SenseFunctionUnicode"
            Me.SenseFunctionUnicode = This.ViewModel.SenseFunctionUnicode
        Case "MeasuredReading"
            Me.MeasuredReading = This.ViewModel.MeasuredReading
        Case "MeasuredValueFormat"
            Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
        
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "MeasuredValueFormat"
            Me.MeasuredValueFormat = This.ViewModel.MeasuredValueFormat
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        Case "AutoIncrementDutNumberEnabled"
            Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
        Case "SenseFunctions"
            Me.SenseFunctions = This.ViewModel.SenseFunctions
        Case "ScpiCommands"
            Me.ScpiCommands = This.ViewModel.ScpiCommands
        Case "Measuring"
            Me.Measuring = This.ViewModel.Measuring
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Me.MonitoringState = This.ViewModel.MonitoringState
        
        Case "AutoAssertTalkGetExecutable"
            Me.AutoAssertTalkGetExecutable = This.ViewModel.AutoAssertTalkGetExecutable
        Case "AutoAssertTalkSetExecutable"
            Me.AutoAssertTalkSetExecutable = This.ViewModel.AutoAssertTalkSetExecutable
        Case "ClearExecutionStateExecutable"
            Me.ClearExecutionStateExecutable = This.ViewModel.ClearExecutionStateExecutable
        Case "ClearReadingsExecutable"
            Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
        Case "CloseConnectionExecutable"
            Me.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
        Case "ExternalTriggerOptionExecutable"
            Me.ExternalTriggerOptionExecutable = This.ViewModel.ExternalTriggerOptionExecutable
        Case "FunctionNameExecutable"
            Me.FunctionNameExecutable = This.ViewModel.FunctionNameExecutable
        Case "GpibAddressGetExecutable"
            Me.GpibAddressGetExecutable = This.ViewModel.GpibAddressGetExecutable
        Case "GpibAddressSetExecutable"
            Me.GpibAddressSetExecutable = This.ViewModel.GpibAddressSetExecutable
        Case "GpibLanReadTimeoutGetExecutable"
            Me.GpibLanReadTimeoutGetExecutable = This.ViewModel.GpibLanReadTimeoutGetExecutable
        Case "GpibLanReadTimeoutSetExecutable"
            Me.GpibLanReadTimeoutSetExecutable = This.ViewModel.GpibLanReadTimeoutSetExecutable
        Case "GoToLocalExecutable"
            Me.GoToLocalExecutable = This.ViewModel.GoToLocalExecutable
        Case "ImmediateTriggerOptionExecutable"
            Me.ImmediateTriggerOptionExecutable = This.ViewModel.ImmediateTriggerOptionExecutable
        Case "LocalLockoutExecutable"
            Me.LocalLockoutExecutable = This.ViewModel.LocalLockoutExecutable
        Case "MeasureExecutable"
            Me.MeasureExecutable = This.ViewModel.MeasureExecutable
        Case "OpenConnectionExecutable"
            Me.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
        Case "QueryExecutable"
            Me.QueryExecutable = This.ViewModel.QueryExecutable
        Case "QueryInputsExecutable"
            Me.QueryInputsExecutable = This.ViewModel.QueryInputsExecutable
        Case "QueryStatusByteExecutable"
            Me.QueryStatusByteExecutable = This.ViewModel.QueryStatusByteExecutable
        Case "ReadExecutable"
            Me.ReadExecutable = This.ViewModel.ReadExecutable
        Case "ReadStandardEventStatusExecutable"
            Me.ReadStandardByteExecutable = This.ViewModel.ReadStandardEventStatusExecutable
        Case "ReadStatusByteExecutable"
            Me.ReadStatusByteExecutable = This.ViewModel.ReadStatusByteExecutable
            Me.SerialPollExecutable = This.ViewModel.ReadStatusByteExecutable
        Case "ResetKnownStateExecutable"
            Me.ResetKnownStateExecutable = This.ViewModel.ResetKnownStateExecutable
        Case "SelectiveDeviceClearExecutable"
            Me.SelectiveDeviceClearExecutable = This.ViewModel.SelectiveDeviceClearExecutable
        Case "ServiceRequestedExecutable"
            Me.ServiceRequestedExecutable = This.ViewModel.ServiceRequestedExecutable
        Case "StartMonitoringExecutable"
            Me.StartMonitoringExecutable = This.ViewModel.StartMonitoringExecutable
        Case "StopMonitoringExecutable"
            Me.StopMonitoringExecutable = This.ViewModel.StopMonitoringExecutable
        Case "WriteExecutable"
            Me.WriteExecutable = This.ViewModel.WriteExecutable

    End Select
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_viewModel_PropertyChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not a_eventArgs Is Nothing Then
    
        Dim p_propertyName As String
        p_propertyName = a_eventArgs.PropertyName
        HandleViewModelPropertyChanged p_propertyName
    
    End If
   
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' append a description
    Err.Description = Err.Description & ". Property name: '" & p_propertyName & "'."
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Worksheet events
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub OnWorksheetChange(ByVal a_range As Excel.Range)

    ' If Not Intersect(a_range, Me.Me.Range("H5")) Is Nothing Then
    ' this uses fewer resources than Intersect,
    ' which will be helpful if your worksheet changes a lot.
    Select Case a_range.Address
        
        Case m_readingFormatCell
            This.ViewModel.MeasuredValueFormat = Me.MeasuredValueFormat

    End Select

End Sub


''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub Worksheet_Change(ByVal a_range As Excel.Range)

    If Not This.ViewModel Is Nothing Then _
        OnWorksheetChange a_range

End Sub

