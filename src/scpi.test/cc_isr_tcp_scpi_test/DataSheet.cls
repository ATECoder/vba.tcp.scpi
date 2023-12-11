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

Private Const m_rearInputsSenseFunctionName As String = "RES"
Private Const m_frontInputsSenseFunctionName As String = "FRES"

Private Const m_sessionTimeoutCell As String = "B1"
Private Const m_socketAddressCell As String = "B2"
Private Const m_timerIntervalCell As String = "B3"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B16"
Private Const m_maximumDutNumber As Integer = 48
Private Const m_gpibLanControllerPort As Long = 1234
Private Const m_primaryGpibAddress As Integer = 16

Private Type this_
    ViewModel As cc_isr_Tcp_Scpi.K2700ViewModel
    ViewModelOwner As Boolean
    FocusedControlName As String
    RearInputsSenseFunctionName As String
    FrontInputsSenseFunctionName As String
    DutNumberCaptionPrefix As String
    DutTitle As String
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Construction and cleanup
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize(ByVal a_viewModel As K2700ViewModel) As DataSheet
    
    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    This.ViewModelOwner = False
    
    Set This.ViewModel = a_viewModel
    
    ' set the initial values that are provided by the user interface
    This.DutNumberCaptionPrefix = "r"
    This.DutTitle = "resistance"
    
    If DataView.Initialized Then DataView.Dispose
    If UserView.Initialized Then UserView.Dispose
    
    If DataView.Initialized Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & ".DataSheet." & p_procedureName, _
            " Duplicate initiations."
    End If
    
    ' initialize the views before initializing the view model
    ' the view model gets initialized upon connecting.
    DataView.Initialize a_viewModel, Me
    UserView.Initialize a_viewModel, Me
    
    Set Initialize = Me

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

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

End Function

''' <summary>   Initializes this object if not initialized. </summary>
Private Sub Initialize_()

    If Not DataView.Initialized Then
    
        Me.Initialize cc_isr_Tcp_Scpi.Factory.NewK2700ViewModel
        This.ViewModelOwner = True
        
    End If
    
End Sub

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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

''' <summary>   Gets the maximum DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutNumber() As Integer
    MaximumDutNumber = m_maximumDutNumber
End Property

''' <summary>   Gets the status of the Beep Allowed condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get BeepAllowed() As Boolean
    BeepAllowed = Me.BeepAllowedCheckBox.Value
End Property

''' <summary>   Gets the DUT number Caption Prefix. </summary>
''' <value>   [String]. </value>
Public Property Get DutNumberCaptionPrefix() As String
    DutNumberCaptionPrefix = This.DutNumberCaptionPrefix
End Property

''' <summary>   Gets the DUT title. </summary>
''' <value>   [String]. </value>
Public Property Get DutTitle() As String
    DutTitle = This.DutTitle
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

''' <summary>   Gets the execution Can Execute) status of the Open Connection command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    OpenConnectionExecutable = Me.OpenConnectionButton.Enabled
End Property

''' <summary>   Sets the execution Can Execute) status of the Open Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    Me.OpenConnectionButton.Enabled = a_value
    ' allow power on reset only if not connected.
    Me.PowerOnResetButton.Enabled = a_value
End Property

''' <summary>   Handles the open Connection command event. </summary>
Private Sub OnOpenConnection()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' force initialization
    DataView.Dispose

    ' initialize the views if necessary
    Initialize_
    
    DataView.OnClearMessages
    
    DataView.OnOpenConnectionButtonClick

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
      Me.PowerOnResetButton.Enabled = Not This.ViewModel.Connected
      Me.OpenConnectionButton.Enabled = Not This.ViewModel.Connected
      Me.CloseConnectionButton.Enabled = This.ViewModel.Connected
      Me.ClearReadingsButton.Enabled = This.ViewModel.Connected
        
    End If
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

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

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' force initialization
    DataView.Dispose

    ' initialize the views if necessary
    Initialize_
    
    DataView.OnClearMessages
    
    DataView.OnPowerOneResetButtonClick

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Me.PowerOnResetButton.Enabled = True
    Me.OpenConnectionButton.Enabled = True
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

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Sense Functions
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Sense Function Name for rear inputs. </summary>
''' <value>   [String]. </value>
Public Property Get RearInputsSenseFunctionName() As String
    RearInputsSenseFunctionName = m_rearInputsSenseFunctionName
End Property

''' <summary>   Gets the Sense Function Name for front inputs. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsSenseFunctionName() As String
    FrontInputsSenseFunctionName = m_frontInputsSenseFunctionName
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Inputs
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required for immediate trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ImmediateFrontInputsRequired() As Boolean
    ImmediateFrontInputsRequired = Me.FrontAutoScanCheckBox.Value
End Property

''' <summary>   Gets the Front Inputs Required for External trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExternalFrontInputsRequired() As Boolean
    ExternalFrontInputsRequired = True
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' Worksheet event handlers
' . . . . . . . . . . . . . . . . . . . . . . . . .

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

