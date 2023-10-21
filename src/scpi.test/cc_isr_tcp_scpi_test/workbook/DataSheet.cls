VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The data sheet. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sessionTimeoutCell As String = "B1"
Private Const m_socketAddressCell As String = "B2"
Private Const m_timerIntervalCell As String = "B3"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B16"
Private Const m_maxResistanceCount As Integer = 48
Private Const m_gpibLanControllerPort As Long = 1234
Private Const m_primaryGpibAddress As Integer = 16

Private Type this_
    FocusedControlName As String
    DesignMode As Boolean
End Type

Private This As this_

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize() As DataSheet
    
    Set Initialize = Me

End Function

''' <summary>   Initializes the views. </summary>
Private Sub InitializeViews()

    If Not DataView.Initialized Then
    
        Dim p_viewModel As cc_isr_Tcp_Scpi.K2700ViewModel
        
        Set p_viewModel = cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        
        ' initialize the views before initializing the view model
        ' the view model gets initialized upon connecting.
        DataView.Initialize p_viewModel
        UserView.Initialize p_viewModel
    
    End If
    
    UserView.AutoSampleFrontInputs = Me.FrontAutoScanCheckBox.Value

End Sub

''' <summary>   Gets the socket address. </summary>
''' <value>   [String]. </value
Public Property Get SocketAddress() As String
    SocketAddress = VBA.CStr(Range(m_socketAddressCell).Value)
End Property

''' <summary>   Gets the session timeout. </summary>
''' <value>   [Long]. </value>
Public Property Get SessionTimeout() As Long
    SessionTimeout = VBA.CLng(Range(m_sessionTimeoutCell).Value)
End Property

''' <summary> Get the timer interval. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = VBA.CInt(Range(m_timerIntervalCell).Value)
End Property

''' <summary>   Gets the Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Long
    PrimaryGpibAddress = m_primaryGpibAddress
End Property

''' <summary>   Gets the Gpib Lan Controller Port number. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = m_gpibLanControllerPort
End Property

''' <summary>   Gets the maximum resistance count. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaxResistanceCount() As Integer
    MaxResistanceCount = m_maxResistanceCount
End Property

''' <summary>   Gets the status of the Beep Allowed condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get BeepAllowed() As Boolean
    BeepAllowed = Me.BeepAllowedCheckBox.Value
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).Value
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).Value = a_value
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = Range(m_lastMessageCell).Value
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the message. </value>
Public Property Let LastMessage(ByVal a_value As String)
    Range(m_lastMessageCell).Value = a_value
End Property

''' <summary>   Handles the change event of the Beep Allowed check box. </summary>
Private Sub BeepAllowedCheckBox_Change()
    DataView.BeepAllowed = Me.BeepAllowedCheckBox.Value
End Sub

''' <summary>   Handles the click event of the Clear Messages button. </summary>
Private Sub ClearMessagesButton_Click()
    DataView.OnClearMessages
End Sub

''' <summary>   Handles the click event of the Clear Radings button. </summary>
Private Sub ClearReadingsButton_Click()
    DataView.OnClearReadings
End Sub

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

''' <summary>   Gets the execution Can Execute) status of the Open Connection command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = Me.OpenConnectionButton.Enabled
End Property

''' <summary>   Sets the execution Can Execute) status of the Open Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    Me.OpenConnectionButton.Enabled = a_value
End Property

''' <summary>   Handles the open Connection command event. </summary>
Private Sub OnOpenConnection()

    ' initialize the views if necessary
    InitializeViews
    
    DataView.OnOpenConnectionButtonClick

End Sub

Private Sub OpenConnectionButton_Click()
    OnOpenConnection
End Sub

''' <summary>   Gets the execution Can Execute) status of the Close Connection command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    CloseConnectionExecutable = Me.CloseConnectionButton.Enabled
End Property

''' <summary>   Sets the execution Can Execute) status of the Close Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    Me.CloseConnectionButton.Enabled = a_value
    Me.PowerOnResetButton.Enabled = a_value
End Property

''' <summary>   Handles the open Connection command event. </summary>
Private Sub OnCloseConnection()

    If DataView.Initialized Then
    
        DataView.OnCloseConnectionButtonClick
    
    End If

End Sub

Private Sub CloseConnectionButton_Click()
    OnCloseConnection
End Sub

''' <summary>   Handles the click event of the power on reset.  </summary>
Private Sub PowerOnResetButton_Click()
    OnPowerOnReset
End Sub

''' <summary>   Handles the power on reset command event. </summary>
Private Sub OnPowerOnReset()

    ' initialize the views if necessary
    InitializeViews
    
    DataView.OnPowerOneResetButtonClick

End Sub

''' <summary>   Handles the change event of the spreadsheet. </summary>
''' <remarks>   Updates the relevant values of the <see cref="DataView"/> </remarks>
''' <param name="a_target">   [<see cref="Excel.Range"/>] the range fo the change event. </value>
Private Sub Worksheet_Change(ByVal a_target As Excel.Range)

    If a_target Is Nothing Then
    ElseIf a_target.Address = m_sessionTimeoutCell Then
        DataView.SessionTimeout = Me.SessionTimeout
    ElseIf a_target.Address = m_socketAddressCell Then
        DataView.SocketAddress = Me.SocketAddress
    ElseIf a_target.Address = m_timerIntervalCell Then
        DataView.TimerInterval = Me.TimerInterval
    End If

End Sub

