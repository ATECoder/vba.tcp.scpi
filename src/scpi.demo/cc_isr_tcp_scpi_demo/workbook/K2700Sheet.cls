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
Private Const m_topCardCell As String = "C23"
Private Const m_bottomCardCell As String = "C24"
Private Const m_topCardScanListCell As String = "C25"
Private Const m_bottomCardScanListCell As String = "C26"
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
    MaximumChannelNumber As Integer
    
    ' IEEE488
    ScpiCommands As VBA.Collection
    SenseFunctions As VBA.Collection
    
    ' command info
    LastCommandOutcome As Boolean
    LastCommandDetails As String
    
    ' data acquisition
    Measuring As Boolean
    MeasurmentMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    
    MeasurementRow As Integer
    MeasurementColumn As Integer
    ChannelRow As Integer
    ChannelColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String
    ChannelNumberCaptionPrefix As String
    
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
    Me.GpibLanControllerPort = 1234
    Me.ChannelNumberCaptionPrefix = "ch"
    Me.MaximumChannelNumber = 48
    Me.FrontInputsRequired = True
    Me.SingleReadEnabled = True
    Me.ImmediateTriggerOptionButton.Value = True
    Me.ExternalTriggerOptionButton.Value = False
    
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Function

''' <summary>   Initializes this object if not initialized. </summary>
Private Sub Initialize_()

    If m_viewModel Is Nothing Then
    
        Me.Initialize cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        This.ViewModelOwner = True
        
    End If
    
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
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
    If 0 = Len(Me.Range(m_socketAddressCell).Value) Then Me.Range(m_socketAddressCell).Value = "192.168.0.252"
    SocketAddress = VBA.CStr(Me.Range(m_socketAddressCell).Value)
End Property

Public Property Let SocketAddress(ByVal a_value As String)
    If Me.SocketAddress <> a_value Then
        Me.Range(m_socketAddressCell).Value = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SocketAddress = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SessionTimeout = a_value
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

''' <summary>   Handles open connection button click event. </summary>
Private Sub OpenConnectionButton_Click()
    ' initialize if necessary
    Initialize_
    
    This.ViewModel.OpenConnectionCommand
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
    If This.MessageToSend <> a_value Then
        Me.CommandsComboBox.Text = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MessageToSend = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SentMessage = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReceivedMessage = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReceivedMessageLength = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.StatusByte = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.StandardByte = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadStatusAfterWriteEnabled = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ScpiCommands = a_value
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
        Me.PopulateFunctionComboBox
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SenseFunctions = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' IEEE488 Control
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
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

    If This.ViewModel Is Nothing Then Exit Sub

    If This.FocusedControlName = Me.CommandsComboBox.Name Then

        On Error GoTo 0

        ' Commands for click
        This.ViewModel.MessageToSend = Me.MessageToSend
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        This.ViewModel.MessageToSend = Me.MessageToSend

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    This.ViewModel.MessageToSend = Me.MessageToSend

End Sub

''' <summary>   Populates the Commands combo box. </summary>
Public Sub PopulateCommandsComboBox()

    Const p_procedureName = "PopulateCommandsComboBox"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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

    If This.ViewModel Is Nothing Then Exit Sub
    
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
    This.ViewModel.ReadStatusByteCommand
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    This.LastCommandOutcome = This.ViewModel.ResetKnownStateCommand(This.LastCommandDetails)
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()
    This.ViewModel.WriteCommand
End Sub

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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.AutoAssertTalk = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibAddress = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibLanReadTimeout = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SerialPollByte = a_value
    End If
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_value As Integer)
    If This.ServiceRequested <> a_value Then
        This.ServiceRequested = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ServiceRequested = a_value
    End If
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
    This.ViewModel.LocalLockoutCommand
End Sub

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()
    This.ViewModel.OnLocalLockout
End Sub

''' <summary>   Handles the click event of the power on reset.  </summary>
Private Sub PowerOnResetButton_Click()
    ' initialize if necessary
    Initialize_
    This.ViewModel.OnPowerOneResetButtonClick
End Sub

''' <summary>   Handles the read after write (Auto Assert TALK) set button click event. </summary>
Private Sub ReadAfterWriteSetButton_Click()
    This.ViewModel.AutoAssertTalkSetCommand
End Sub

''' <summary>   Handles the read after write (Auto Assert TALK) get button click event. </summary>
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
    This.ViewModel.ReadStatusByteCommand
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
    If 0 = Len(Me.Range(m_timerIntervalCell).Value) Then Me.Range(m_timerIntervalCell).Value = 500
    TimerInterval = VBA.CInt(Me.Range(m_timerIntervalCell).Value)
End Property

Public Property Let TimerInterval(ByVal a_value As Integer)
    If Me.TimerInterval <> a_value Then
        Me.Range(m_timerIntervalCell).Value = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TimerInterval = a_value
    End If
End Property

''' <summary>   Gets the selected channel number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedChannelNumber() As Integer
    If 0 = Len(Me.Range(m_channelNumberCell).Value) Then Me.Range(m_channelNumberCell).Value = 1
    SelectedChannelNumber = VBA.CInt(Me.Range(m_channelNumberCell).Value)
End Property

''' <summary>   Sets the selected channel number. </summary>
''' <param name="a_value">   [Integer] the selected channel number. </param>
Public Property Let SelectedChannelNumber(ByVal a_value As Integer)
    Me.Range(m_channelNumberCell).Value = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TopCard = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TopCardFunctionScanList = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.BottomCard = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.BottomCardFunctionScanList = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Value. </summary>
''' <remarks>   True if front inputs; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsValue() As Boolean
    If 0 = Len(Me.Range(m_inputsCell).Value) Then Me.Range(m_inputsCell).Value = Me.FrontInputsRequired
    FrontInputsValue = VBA.CBool(Me.Range(m_inputsCell).Value)
End Property

''' <summary>   Sets the Front Inputs Value. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs. </value>
Public Property Let FrontInputsValue(ByVal a_value As Boolean)
    If Me.FrontInputsValue <> a_value Then
        Me.Range(m_inputsCell).Value = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsValue = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Data Acquisition properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the maximum Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumChannelNumber() As Integer
    MaximumChannelNumber = This.MaximumChannelNumber
End Property

''' <summary>   Sets the maximum Channel Number. </summary>
''' <param name="a_value">   [Integer] the maximum Channel Number value. </value>
Public Property Let MaximumChannelNumber(ByVal a_value As Integer)
    If This.MaximumChannelNumber <> a_value Then
        This.MaximumChannelNumber = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.MaximumChannelNumber = a_value
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

''' <summary>   Gets the Measured channel number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredChannelNumber() As Integer
    If 0 = Len(Me.Range(m_measuredChannelNumberCell).Value) Then Me.Range(m_measuredChannelNumberCell).Value = -1
    MeasuredChannelNumber = VBA.CInt(Me.Range(m_measuredChannelNumberCell).Value)
End Property

''' <summary>   Sets the Measured channel number. </summary>
''' <param name="a_value">   [Integer] the Measured channel number. </param>
Public Property Let MeasuredChannelNumber(ByVal a_value As Integer)
    If Me.MeasuredChannelNumber <> a_value Then
        Me.Range(m_measuredChannelNumberCell).Value = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadingOffset = a_value
    End If
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

    If This.ViewModel Is Nothing Then Exit Sub

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
    This.ViewModel.ClearReadingsCommand
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

    If This.ViewModel Is Nothing Then Exit Sub

    If This.FocusedControlName = Me.ExternalTriggerOptionButton.Name Then

        On Error GoTo 0

        ' Commands for click
        If Me.ExternalTriggerOptionButton.Value Then
            This.ViewModel.ExternalSenseFunctionName = Me.FunctionComboBox.Text
            This.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(This.LastCommandDetails)
        End If
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        If Me.ExternalTriggerOptionButton.Value Then
            This.ViewModel.ExternalSenseFunctionName = Me.FunctionComboBox.Text
            This.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(This.LastCommandDetails)
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
    If Me.ExternalTriggerOptionButton.Value Then
        This.ViewModel.ExternalSenseFunctionName = Me.FunctionComboBox.Text
        This.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(This.LastCommandDetails)
    End If
    
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
    
    If This.ViewModel Is Nothing Then Exit Sub

    If This.FocusedControlName = Me.ImmediateTriggerOptionButton.Name Then

        On Error GoTo 0

        ' Commands for click
        If ImmediateTriggerOptionButton.Value Then
            This.ViewModel.ImmediateSenseFunctionName = Me.FunctionComboBox.Text
            This.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(This.LastCommandDetails)
        End If
    
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        If ImmediateTriggerOptionButton.Value Then
            This.ViewModel.ImmediateSenseFunctionName = Me.FunctionComboBox.Text
            This.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(This.LastCommandDetails)
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
    If ImmediateTriggerOptionButton.Value Then
        This.ViewModel.ImmediateSenseFunctionName = Me.FunctionComboBox.Text
        This.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(This.LastCommandDetails)
    End If
    
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

    If This.ViewModel Is Nothing Then Exit Sub

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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsRequired = a_value
    End If
End Property

''' <summary>   Handles the measure button click event. Takes a single or multiple
'''             readings from the front or rear panel. </summary>
Private Sub MeasureButton_Click()
    ' make an immediate measure of one or all channels.
    This.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(This.LastCommandDetails)
End Sub

''' <summary>   Handles the start monitoring button click event. </summary>
Private Sub StartMonitoringButton_Click()
    This.ViewModel.StartMonitoringExternalTriggersCommand
End Sub

''' <summary>   Handles the stop monitoring button click event. </summary>
Private Sub StopMonitoringButton_Click()
    This.ViewModel.StopMonitoringExternalTriggersCommand
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

    If This.ViewModel Is Nothing Then Exit Sub

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
    For Each p_item In This.SenseFunctions
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

''' <summary>   Gets the Channel Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelRow() As Integer
    ChannelRow = This.ChannelRow
End Property

''' <summary>   Sets the Channel Row. </summary>
''' <param name="a_value">   [Integer] the Channel Row. </value>
Public Property Let ChannelRow(ByVal a_value As Integer)
    If This.ChannelRow <> a_value Then
        This.ChannelRow = a_value
    End If
End Property

''' <summary>   Gets the Channel Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelColumn() As Integer
    ChannelColumn = This.ChannelColumn
End Property

''' <summary>   Sets the Channel Column. </summary>
''' <param name="a_value">   [Integer] the Channel Column. </value>
Public Property Let ChannelColumn(ByVal a_value As Integer)
    If This.ChannelColumn <> a_value Then
        This.ChannelColumn = a_value
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
    ChannelCellValue = Me.Cells(Me.ChannelRow, Me.ChannelColumn).Value
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If Me.ChannelCellValue <> a_value Then
        This.ChannelCellValue = a_value
        Me.Cells(Me.ChannelRow, Me.ChannelColumn).Value = a_value
    End If
End Property

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
Public Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double)
    SetReading_ a_channelNo, a_value, m_byRow, m_displayIncrement
End Sub

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_byRow">       [Boolean] true if the next reading goes to the next row. </param>
''' <param name="a_increment">   [Integer] the row or column increment. </param>
Private Sub SetReading_(ByVal a_channelNo As Integer, ByVal a_value As Double, _
    ByVal a_byRow As Boolean, ByVal a_increment As Integer)

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_channelNo - 1), _
        m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, _
        m_measurementCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelRow = IIf(a_byRow, m_channelRow + a_increment * (a_channelNo - 1), _
        m_channelRow)
    Me.ChannelColumn = IIf(a_byRow, m_channelCol, _
        m_channelCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelCellValue = This.ChannelNumberCaptionPrefix & VBA.CStr(a_channelNo)
    Me.MeasurementCellValue = VBA.CStr(a_value)
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' View Model Events
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the K2700 <see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Private Sub m_viewModel_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_viewModel_ChannelMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        SetReading a_eventArgs.ChannelNumber, -1#
    Else
        ' let the observer make these updates.
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
            Me.Range(m_socketHostCell).Value = This.ViewModel.Host
        Case "Port"
            Me.Range(m_socketPortCell).Value = This.ViewModel.Port
        Case "SocketAddress"
            Me.Range(m_socketAddressCell).Value = This.ViewModel.SocketAddress
        Case "SocketId"
            Me.Range(m_socketIdCell).Value = This.ViewModel.SocketId
        Case "SessionTimeout"
            Me.Range(m_sessionTimeoutCell).Value = This.ViewModel.SessionTimeout
        Case "PauseRequested"
        Case "TimerInterval"
            Me.Range(m_timerIntervalCell).Value = This.ViewModel.TimerInterval
        Case "Connected"
            Me.Connected = This.ViewModel.Connected
        Case "ElapsedTime"
            ' Me.Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTime
        Case "ElapsedTimeFormat"
            Me.Range(m_elapsedTimeFormatCell).Value = This.ViewModel.ElapsedTimeFormat
        Case "ElapsedTimeCaption"
            Me.Range(m_elapsedTimeCell).Value = This.ViewModel.ElapsedTimeCaption
        
        Case "MessageToSend"
        Case "SentMessage"
            Me.Range(m_sentMessageCell).Value = This.ViewModel.SentMessage
        Case "ReceivedMessage"
            Me.Range(m_receivedMessageCell).Value = This.ViewModel.ReceivedMessage
        Case "ReceivedMessageLength"
            Me.Range(m_receivedMessageLengthCell).Value = This.ViewModel.ReceivedMessageLength
        Case "ReadStatusAfterWriteEnabled"
            Me.ReadStatusAfterWriteCheckBox.Value = This.ViewModel.ReadStatusAfterWriteEnabled
        Case "LastErrorMessage"
            Me.Range(m_lastErrorMessageCell).Value = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            Me.Range(m_lastMessageCell).Value = This.ViewModel.LastMessage
        Case "SerialPollByte"
            Me.Range(m_serialPollCell).Value = This.ViewModel.SerialPollByte
        Case "StatusByte"
            Me.Range(m_statusByteCell).Value = This.ViewModel.StatusByte
        Case "StandardByte"
            Me.Range(m_standardByteCell).Value = This.ViewModel.StandardByte
        Case "ServiceRequested"
            Me.Range(m_srqCell).Value = This.ViewModel.ServiceRequested
        Case "GpibAddress"
            Me.Range(m_gpibAddressCell).Value = This.ViewModel.GpibAddress
        Case "AutoAssertTalk"
            Me.Range(m_autoAssertTalkCell).Value = This.ViewModel.AutoAssertTalk
        Case "GpibLanReadTimeout"
            Me.Range(m_autoAssertTalkCell).Value = This.ViewModel.GpibLanReadTimeout
            
        Case "TopCard"
            Me.Range(m_topCardCell).Value = This.ViewModel.TopCard
        Case "BottomCard"
            Me.Range(m_bottomCardCell).Value = This.ViewModel.BottomCard
        Case "TopCardFunctionScanList"
            Me.Range(m_topCardScanListCell).Value = This.ViewModel.TopCardFunctionScanList
        Case "BottomCardFunctionScanList"
            Me.Range(m_bottomCardScanListCell).Value = This.ViewModel.BottomCardFunctionScanList
        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "FrontInputsValue"
        Case "FrontInputsCaption"
            Me.Range(m_inputsCell).Value = This.ViewModel.FrontInputsCaption
        Case "SelectedChannelNumber"
            Me.Range(m_channelNumberCell).Value = This.ViewModel.SelectedChannelNumber
        Case "ChannelCount"
            Me.Range(m_channelCountCell).Value = This.ViewModel.ChannelCount
        Case "CurrentChannelNumber"
        Case "MeasuredChannelNumber"
            Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
        Case "SenseFunctionName"
        Case "SenseFunctionCaption"
            Me.Range(m_functionNameCell).Value = This.ViewModel.SenseFunctionCaption
        Case "SenseFunctionUnit"
            Me.Range(m_functionUnitCell).Value = This.ViewModel.SenseFunctionUnit
        Case "MeasuredReading"
            Me.Range(m_functionValueCell).Value = This.ViewModel.MeasuredReading
        Case "MeasuredValueFormat"
            Me.Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "ExternalTrigMonitoringEnabled"
            Me.ExternalTrigMonitoringEnabled = This.ViewModel.ExternalTrigMonitoringEnabled
        Case "MeasuredValueFormat"
            Me.Range(m_readingFormatCell).Value = This.ViewModel.MeasuredValueFormat
        Case "AutoIncrementChannelNoEnabled"
            Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
        Case "SenseFunctions"
            Me.SenseFunctions = This.ViewModel.SenseFunctions
        Case "ScpiCommands"
            Me.ScpiCommands = This.ViewModel.ScpiCommands
        ' !@# do we need this?
        Case "ImmediateTriggerOptionSelected"
            ' Me.ImmediateTriggerOptionButton.Value = This.ViewModel.ImmediateTriggerOptionSelected
        ' !@# do we need this?
        Case "ExternalTriggerOptionSelected"
            ' Me.ExternalTriggerOptionButton.Value = This.ViewModel.ExternalTriggerOptionSelected
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "MonitoringState"
            Me.Range(m_monitoringStateCell).Value = This.ViewModel.MonitoringState
        
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

''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub OnWorksheetChange(ByVal a_range As Excel.Range)

    ' If Not Intersect(a_range, Me.Me.Range("H5")) Is Nothing Then
    ' this uses fewer resources than Intersect,
    ' which will be helpful if your worksheet changes a lot.
    Select Case a_range.Address
        
        Case m_socketHostCell
            This.ViewModel.SocketAddress = Me.Range(m_socketHostCell).Value & ":" & _
                Me.Range(m_socketPortCell).Value
        Case m_socketPortCell
            This.ViewModel.SocketAddress = Me.Range(m_socketHostCell).Value & ":" & _
                Me.Range(m_socketPortCell).Value
        Case m_socketAddressCell
            This.ViewModel.SocketAddress = Me.SocketAddress
        Case m_socketIdCell
            This.ViewModel.SocketId = Me.Range(m_socketIdCell).Value
        Case m_sessionTimeoutCell
            This.ViewModel.SessionTimeout = Me.SessionTimeout
            
        Case m_connectedCell
            ' connected is read only
            ' This.ViewModel.Connected = Me.Range(m_connectedCell).Value
        Case m_sentMessageCell
            This.ViewModel.SentMessage = Me.Range(m_sentMessageCell).Value
        Case m_receivedMessageLengthCell
            This.ViewModel.ReceivedMessageLength = Me.Range(m_receivedMessageLengthCell).Value
        Case m_receivedMessageCell
            This.ViewModel.ReceivedMessage = Me.Range(m_receivedMessageCell).Value
        Case m_elapsedTimeCell
            This.ViewModel.ElapsedTimeCaption = Me.Range(m_elapsedTimeCell).Value
        Case m_elapsedTimeFormatCell
            This.ViewModel.ElapsedTimeFormat = Me.Range(m_elapsedTimeFormatCell).Value
        Case m_lastMessageCell
            This.ViewModel.LastMessage = Me.Range(m_lastMessageCell).Value
        Case m_lastErrorMessageCell
            This.ViewModel.LastErrorMessage = Me.Range(m_lastErrorMessageCell).Value
        Case m_serialPollCell
            This.ViewModel.SerialPollByte = Me.Range(m_serialPollCell).Value
        Case m_statusByteCell
            This.ViewModel.StatusByte = Me.Range(m_statusByteCell).Value
        Case m_standardByteCell
            This.ViewModel.StandardByte = Me.Range(m_standardByteCell).Value

        Case m_srqCell
            This.ViewModel.ServiceRequested = Me.Range(m_srqCell).Value
        Case m_gpibAddressCell
            This.ViewModel.GpibAddress = Me.Range(m_gpibAddressCell).Value
        Case m_readTimeoutCell
            This.ViewModel.GpibLanReadTimeout = Me.Range(m_readTimeoutCell).Value
        Case m_autoAssertTalkCell
            This.ViewModel.AutoAssertTalk = VBA.CBool(Me.Range(m_autoAssertTalkCell).Value)
        Case m_timerIntervalCell
            This.ViewModel.TimerInterval = Me.Range(m_timerIntervalCell).Value

        Case m_readingFormatCell
            This.ViewModel.MeasuredValueFormat = Me.Range(m_readingFormatCell).Value
        Case m_topCardCell
            This.ViewModel.TopCard = Me.Range(m_topCardCell).Value
        Case m_bottomCardCell
            This.ViewModel.TopCard = Me.Range(m_bottomCardCell).Value
        Case m_topCardScanListCell
            This.ViewModel.TopCardFunctionScanList = Me.Range(m_topCardScanListCell).Value
        Case m_bottomCardScanListCell
            This.ViewModel.BottomCardFunctionScanList = Me.Range(m_bottomCardScanListCell).Value
        Case m_inputsCell
            This.ViewModel.FrontInputsCaption = Me.Range(m_inputsCell).Value
        Case m_channelNumberCell
            This.ViewModel.SelectedChannelNumber = Me.Range(m_channelNumberCell).Value
        Case m_channelCountCell
            This.ViewModel.ChannelCount = Me.Range(m_channelCountCell).Value
        Case m_functionNameCell
            This.ViewModel.SenseFunctionCaption = Me.Range(m_functionNameCell).Value
        Case m_functionValueCell
            This.ViewModel.MeasuredReading = MeasuredReading
        Case m_functionUnitCell
            This.ViewModel.SenseFunctionUnit = Me.Range(m_functionUnitCell).Value
        Case m_monitoringStateCell
            This.ViewModel.MonitoringState = Me.Range(m_monitoringStateCell).Value
        Case m_readingOffsetCell
            This.ViewModel.ReadingOffset = Me.Range(m_readingOffsetCell).Value

    End Select

End Sub


''' <summary>   Handles work sheet change events. </summary>
''' <remarks> The check box change event also causes a click event.
'''           This method issues the click event on actual click.
'''              Note: Make sure to not implement the click event,
Private Sub Worksheet_Change(ByVal a_range As Excel.Range)

    OnWorksheetChange a_range

End Sub

 
