VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The Data View for the view model. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Const m_measurementRow As Integer = 7 ' 2
Private Const m_measurementCol As Integer = 2 ' 23
Private Const m_dutNumberRow As Integer = 7 ' 2
Private Const m_dutNumberCol As Integer = 1 '  24
Private Const m_byRow As Boolean = False ' True
Private Const m_displayIncrement As Integer = 2

Private Type this_

    ViewModel As K2700ViewModel
    DataSheet As DataSheet

    PrimaryGpibAddress As Integer
    GpibLanControllerPort As Long
    MaximumDutNumber As Integer
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    BeepAllowed As Boolean
    SessionTimeout As Integer
    SocketAddress As String
    TimerInterval As Integer
    LastMessage As String
    LastErrorMessage As String
    LastCommandDetails As String
    LastCommandOutcome As Boolean
        
    MeasurementRow As Integer
    MeasurementColumn As Integer
    DutNumberRow As Integer
    DutNumberColumn As Integer
    DutNumberCellValue As String
    MeasurementCellValue As String

    MeasuredDutNumber As Integer
    MeasuredReading As String
    MeasuredValue As Double
    
    ClearReadingExecutable As Boolean
    CloseConnectionExecutable As Boolean
    OpenConnectionExecutable As Boolean
    DesignMode As Boolean
    
    RearInputsSenseFunctionName As String
    ImmediateFrontInputsRequired As Boolean
    FrontInputsSenseFunctionName As String
    ExternalFrontInputsRequired As Boolean
    
End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
End Sub

''' <summary>   Handles the Class Terminate event. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.DataSheet = Nothing

    Set m_viewModel = Nothing
    
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

''' <summary>   Returns this instance. </summary>
''' <value>   [<see cref="DataView"/>]. </value>
Public Property Get Instance() As DataView
    Set Instance = Me
End Property

''' <summary>   Initialize this object. </summary>
''' <returns>   [<see cref="DataView"/>]. </returns>
Public Function Initialize(ByVal a_viewModel As K2700ViewModel, a_dataSheet As DataSheet) As DataView

    Const p_procedureName = "Initialize"
    
    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.ViewModel = a_viewModel
    
    Set This.DataSheet = a_dataSheet
    
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

''' <summary>    Gets the Initialization condition of this object. </summary>
''' <value>   [Boolean] True if initialized. </value>
Public Property Get Initialized() As Boolean
    Initialized = Not This.ViewModel Is Nothing
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' User interface control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the power on reset button click event. </summary>
Public Sub OnPowerOneResetButtonClick()

    Const p_procedureName As String = "OnPowerOneResetButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' issue a power on reset.
    This.ViewModel.PowerOnResetCommand This.DataSheet.SocketAddress, This.DataSheet.SessionTimeout
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Open Connection button Click event. </summary>
Public Sub OnOpenConnectionButtonClick()

    Const p_procedureName As String = "OnOpenConnectionButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if already connected
    If This.ViewModel.Connected Then Exit Sub
        
    This.ViewModel.OpenConnectionCommand This.DataSheet.SocketAddress, This.DataSheet.SessionTimeout
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Open Connection command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get OpenConnectionExecutable() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.OpenConnectionExecutable = This.DataSheet.OpenConnectionExecutable
    End If
    OpenConnectionExecutable = This.OpenConnectionExecutable
End Property

''' <summary>   Sets the execution (Can Execute) status of the Open Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let OpenConnectionExecutable(ByVal a_value As Boolean)
    If Not This.DataSheet Is Nothing Then _
        This.DataSheet.OpenConnectionExecutable = a_value
    This.OpenConnectionExecutable = a_value
End Property

''' <summary>   Handles the Close Connection button Click event. </summary>
Public Sub OnCloseConnectionButtonClick()

    Const p_procedureName As String = "OnCloseConnectionButtonClick"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if already disconnected
    If Not This.ViewModel.Connected Then Exit Sub
        
    This.ViewModel.CloseConnectionCommand
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Gets the execution (Can Execute) status of the Close Connection command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get CloseConnectionExecutable() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.CloseConnectionExecutable = This.DataSheet.CloseConnectionExecutable
    End If
    CloseConnectionExecutable = This.CloseConnectionExecutable
End Property

''' <summary>   Sets the execution (Can Execute) status of the Close Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let CloseConnectionExecutable(ByVal a_value As Boolean)
    If Not This.DataSheet Is Nothing Then _
        This.DataSheet.CloseConnectionExecutable = a_value
    This.CloseConnectionExecutable = a_value
End Property

''' <summary>   Gets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Get ClearReadingsExecutable() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.ClearReadingExecutable = This.DataSheet.ClearReadingsButton.Enabled
    End If
    ClearReadingsExecutable = This.ClearReadingExecutable
End Property

''' <summary>   Sets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If Not This.DataSheet.ClearReadingsButton Is Nothing Then
        This.DataSheet.ClearReadingsButton.Enabled = a_value
    End If
    This.ClearReadingExecutable = a_value
End Property

''' <summary>   Handles the Clear Messages Button Click Event. </summary>
Public Sub OnClearMessages()
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.ClearMessages
    End If
End Sub

''' <summary>   Handles the Clear Messages Button Click Event. </summary>
Public Sub OnClearReadings()
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.OnError VBA.vbNullString
        This.ViewModel.ClearReadingsCommand UserView.ReadingOffset, 1, Me.MaximumDutNumber
    End If
End Sub

''' <summary>   Gets the primary GPIB address of the instrument. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Integer
    If Not This.DataSheet Is Nothing Then
        This.PrimaryGpibAddress = This.DataSheet.PrimaryGpibAddress
    End If
    PrimaryGpibAddress = This.PrimaryGpibAddress
End Property

''' <summary>   Sets the primary GPIB address of the instrument. </summary>
''' <param name="a_value">   [Integer] the primary GPIB address. </param>
Public Property Let PrimaryGpibAddress(ByVal a_value As Integer)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.PrimaryGpibAddress = a_value
    End If
    This.PrimaryGpibAddress = a_value
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    If Not This.DataSheet Is Nothing Then
        This.GpibLanControllerPort = This.DataSheet.GpibLanControllerPort
    End If
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    If Me.GpibLanControllerPort <> a_value Then
        This.GpibLanControllerPort = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibLanControllerPort = a_value
    End If
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.GpibLanControllerPort = a_value
    End If
End Property

''' <summary>   Gets the maximum DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutNumber() As Integer
    If Not This.DataSheet Is Nothing Then _
        This.MaximumDutNumber = This.DataSheet.MaximumDutNumber
    MaximumDutNumber = This.MaximumDutNumber
End Property

''' <summary>   Sets the maximum DUT number. </summary>
''' <param name="a_value">   [Integer] the maximum DUT number value. </value>
Public Property Let MaximumDutNumber(ByVal a_value As Integer)
    This.MaximumDutNumber = a_value
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <remarks>   True if instrument beep sound on each reading is enabled. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get BeepAllowed() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.BeepAllowed = This.DataSheet.BeepAllowed
    End If
    BeepAllowed = This.BeepAllowed
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <param name="a_value">   [Boolean] instrument beep sound on each reading is enabled. </value>
Public Property Let BeepAllowed(ByVal a_value As Boolean)
    If Me.BeepAllowed <> a_value Then
        If Not This.DataSheet Is Nothing Then
            This.DataSheet.BeepAllowedCheckBox.Value = a_value
        End If
        This.BeepAllowed = a_value
    End If
End Property

''' <summary>   Gets the socket and session timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    If Not This.DataSheet Is Nothing Then
        This.SessionTimeout = This.DataSheet.SessionTimeout
    End If
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket and session timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_value As Integer)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.SessionTimeout = a_value
    End If
    This.SessionTimeout = a_value
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    If Not This.DataSheet Is Nothing Then
        This.SocketAddress = This.DataSheet.SocketAddress
    End If
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_value As String)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.SocketAddress = a_value
    End If
    This.SocketAddress = a_value
End Property

''' <summary>   Gets the Timer Interval in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    If Not This.DataSheet Is Nothing Then
        This.TimerInterval = This.DataSheet.TimerInterval
    End If
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Timer Interval in milliseconds value. </value>
Public Property Let TimerInterval(ByVal a_value As Integer)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.TimerInterval = a_value
    End If
    This.TimerInterval = a_value
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    If Not This.DataSheet Is Nothing Then
        This.LastErrorMessage = This.DataSheet.LastErrorMessage
    End If
    LastErrorMessage = This.LastErrorMessage
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the last error message value. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    If This.LastErrorMessage <> a_value Then
        
        This.LastErrorMessage = a_value
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.OnError a_value
            
        If Not This.DataSheet Is Nothing Then _
            This.DataSheet.LastErrorMessage = a_value
            
    End If
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    If Not This.DataSheet Is Nothing Then
        This.LastMessage = This.DataSheet.LastMessage
    End If
    LastMessage = This.LastMessage
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the last message value. </value>
Public Property Let LastMessage(ByVal a_value As String)
    If This.LastMessage <> a_value Then
        This.LastMessage = a_value
        
        If Not This.DataSheet Is Nothing Then _
            This.DataSheet.LastMessage = a_value
            
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Sense functions
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Sense Function Name for rear inputs. </summary>
''' <value>   [String]. </value>
Public Property Get RearInputsSenseFunctionName() As String
    If Not This.DataSheet Is Nothing Then _
        This.RearInputsSenseFunctionName = This.DataSheet.RearInputsSenseFunctionName
    RearInputsSenseFunctionName = This.RearInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for rear inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let RearInputsSenseFunctionName(ByVal a_value As String)
    This.RearInputsSenseFunctionName = a_value
End Property

''' <summary>   Gets the Sense Function Name for front inputs. </summary>
''' <value>   [String]. </value>
Public Property Get FrontInputsSenseFunctionName() As String
    If Not This.DataSheet Is Nothing Then _
        This.FrontInputsSenseFunctionName = This.DataSheet.FrontInputsSenseFunctionName
    FrontInputsSenseFunctionName = This.FrontInputsSenseFunctionName
End Property

''' <summary>   Sets the Sense Function Name for front inputs. </summary>
''' <param name="a_value">   [String] the Sense Function Name. </value>
Public Property Let FrontInputsSenseFunctionName(ByVal a_value As String)
    This.FrontInputsSenseFunctionName = a_value
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: Immediate
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required for Immediate trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ImmediateFrontInputsRequired() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.ImmediateFrontInputsRequired = This.DataSheet.ImmediateFrontInputsRequired
    End If
    ImmediateFrontInputsRequired = This.ImmediateFrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required for Immediate trigger  mode. </summary>
''' <param name="a_value">   [Boolean] the Front Inputs Required. </value>
Public Property Let ImmediateFrontInputsRequired(ByVal a_value As Boolean)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.ImmediateFrontInputsRequired = a_value
    End If
    This.ImmediateFrontInputsRequired = a_value
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Measurement Mode: External Trigger
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Front Inputs Required for External trigger mode. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExternalFrontInputsRequired() As Boolean
    If Not This.DataSheet Is Nothing Then
        This.ExternalFrontInputsRequired = This.DataSheet.ExternalFrontInputsRequired
    End If
    ExternalFrontInputsRequired = This.ExternalFrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required for External trigger  mode. </summary>
''' <param name="a_value">   [Boolean] the Front Inputs Required. </value>
Public Property Let ExternalFrontInputsRequired(ByVal a_value As Boolean)
    If Not This.DataSheet Is Nothing Then
        ' read only: This.DataSheet.ExternalFrontInputsRequired = a_value
    End If
    This.ExternalFrontInputsRequired = a_value
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: DUT Selection
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the DUT number Caption Prefix. </summary>
''' <value>   [String]. </value>
Public Property Get DutNumberCaptionPrefix() As String
    If This.DataSheet Is Nothing Then
        ' this will cause a failure
        DutNumberCaptionPrefix = VBA.vbNullString
    Else
        DutNumberCaptionPrefix = This.DataSheet.DutNumberCaptionPrefix
    End If
End Property

''' <summary>   Gets the DUT title. </summary>
''' <value>   [String]. </value>
Public Property Get DutTitle() As String
    If This.DataSheet Is Nothing Then
        ' this will cause a failure
        DutTitle = VBA.vbNullString
    Else
        DutTitle = This.DataSheet.DutTitle
    End If
End Property

' . . . . . . . . . . . . . . . . . . . . . . . . .
' K2700: Reading
' . . . . . . . . . . . . . . . . . . . . . . . . .

''' <summary>   Gets the Measured DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredDutNumber() As Integer
    MeasuredDutNumber = This.MeasuredDutNumber
End Property

''' <summary>   Sets the Measured DUT number. </summary>
''' <param name="a_value">   [Integer] the Measured DUT number value. </value>
Public Property Let MeasuredDutNumber(ByVal a_value As Integer)
    This.MeasuredDutNumber = a_value
End Property

''' <summary>   Gets the Measured Reading. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = This.MeasuredReading
End Property

''' <summary>   Sets the Measured Reading. </summary>
''' <param name="a_value">   [String] the Measured Reading value. </value>
Public Property Let MeasuredReading(ByVal a_value As String)
    This.MeasuredReading = a_value
End Property

''' <summary>   Gets the Measured Value. </summary>
''' <value>   [Double]. </value>
Public Property Get MeasuredValue() As Double
    MeasuredValue = This.MeasuredValue
End Property

''' <summary>   Sets the Measured Value. </summary>
''' <param name="a_value">   [Double] the Measured Value. </value>
Public Property Let MeasuredValue(ByVal a_value As Double)
    This.MeasuredValue = a_value
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
    This.MeasurementMode = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Measurement Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementRow() As Integer
    MeasurementRow = This.MeasurementRow
End Property

''' <summary>   Sets the Measurement Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Row. </value>
Public Property Let MeasurementRow(ByVal a_value As Integer)
    This.MeasurementRow = a_value
End Property

''' <summary>   Gets the Measurement Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementColumn() As Integer
    MeasurementColumn = This.MeasurementColumn
End Property

''' <summary>   Sets the Measurement Column. </summary>
''' <param name="a_value">   [Integer] the Measurement Column. </value>
Public Property Let MeasurementColumn(ByVal a_value As Integer)
    This.MeasurementColumn = a_value
End Property

''' <summary>   Gets the DUT number row. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumberRow() As Integer
    DutNumberRow = This.DutNumberRow
End Property

''' <summary>   Sets the DUT number row. </summary>
''' <param name="a_value">   [Integer] the DUT number row. </value>
Public Property Let DutNumberRow(ByVal a_value As Integer)
    This.DutNumberRow = a_value
End Property

''' <summary>   Gets the DUT number Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumberColumn() As Integer
    DutNumberColumn = This.DutNumberColumn
End Property

''' <summary>   Sets the DUT number Column. </summary>
''' <param name="a_value">   [Integer] the DUT number Column. </value>
Public Property Let DutNumberColumn(ByVal a_value As Integer)
    This.DutNumberColumn = a_value
End Property

''' <summary>   Gets the Measurement Cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementCellValue() As String
    If This.DataSheet Is Nothing Then
        MeasurementCellValue = This.MeasurementCellValue
    Else
        MeasurementCellValue = This.DataSheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
    End If
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If Me.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        If Not This.DataSheet Is Nothing Then _
            This.DataSheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
    End If
End Property
        
Friend Function GetDutNumberCellRow(ByVal a_dutNumber As Integer)
    GetDutNumberCellRow = IIf(m_byRow, m_dutNumberRow + m_displayIncrement * (a_dutNumber - 1), _
        m_dutNumberRow)
End Function
        
Friend Function GetDutNumberCellColumn(ByVal a_dutNumber As Integer)
    GetDutNumberCellColumn = IIf(m_byRow, m_dutNumberCol, _
        m_dutNumberCol + m_displayIncrement * (a_dutNumber - 1))
End Function
        
Public Function GetDutNumberCellValue(ByVal a_dutNumber As Integer) As String
    If This.DataSheet Is Nothing Then
        GetDutNumberCellValue = VBA.vbNullString
    Else
        GetDutNumberCellValue = This.DataSheet.Cells(Me.GetDutNumberCellRow(a_dutNumber), _
            Me.GetDutNumberCellColumn(a_dutNumber)).Value
    End If
End Function
        
''' <summary>   Gets the DUT number cell value. </summary>
''' <value>   [String]. </value>
Public Property Get DutNumberCellValue() As String
    If This.DataSheet Is Nothing Then
        DutNumberCellValue = This.DutNumberCellValue
    Else
        DutNumberCellValue = This.DataSheet.Cells(Me.DutNumberRow, Me.DutNumberColumn).Value
    End If
End Property

''' <summary>   Sets the DUT number cell value. </summary>
''' <param name="a_value">   [String] the DUT number cell value. </value>
Public Property Let DutNumberCellValue(ByVal a_value As String)
    If Me.DutNumberCellValue <> a_value Then
        
        This.DutNumberCellValue = a_value
        
        If Not This.DataSheet Is Nothing Then _
            This.DataSheet.Cells(Me.DutNumberRow, Me.DutNumberColumn).Value = a_value
        
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

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_dutNumber - 1), m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, m_measurementCol + a_increment * (a_dutNumber - 1))
   
    Me.DutNumberRow = IIf(a_byRow, m_dutNumberRow + a_increment * (a_dutNumber - 1), m_dutNumberRow)
    Me.DutNumberColumn = IIf(a_byRow, m_dutNumberCol, m_dutNumberCol + a_increment * (a_dutNumber - 1))
   
    Me.DutNumberCellValue = This.DataSheet.DutNumberCaptionPrefix & VBA.CStr(a_dutNumber)
    Me.MeasurementCellValue = VBA.CStr(a_value)
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the property change event of the view model. </summary>
''' <param name="a_propertyName">   [String] the property name. </param>
Friend Sub HandleViewModelPropertyChanged(ByVal a_propertyName As String)

    Const p_procedureName = "HandleViewModelPropertyChanged"

    If This.ViewModel Is Nothing Then Exit Sub

    Select Case a_propertyName

        Case "SocketAddress"
            Me.SocketAddress = This.ViewModel.SocketAddress
        Case "SessionTimeout"
            Me.SessionTimeout = This.ViewModel.SessionTimeout
        Case "TimerInterval"
            Me.TimerInterval = This.ViewModel.TimerInterval
        Case "LastErrorMessage"
            Me.LastErrorMessage = This.ViewModel.LastErrorMessage
        Case "LastMessage"
            Me.LastMessage = This.ViewModel.LastMessage
            
        Case "MeasuredDutNumber"
            Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
        Case "MeasuredReading"
            Me.MeasuredReading = This.ViewModel.MeasuredReading
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "RearInputsSenseFunctionName"
            Me.RearInputsSenseFunctionName = This.ViewModel.RearInputsSenseFunctionName
        Case "FrontInputsSenseFunctionName"
            Me.FrontInputsSenseFunctionName = This.ViewModel.FrontInputsSenseFunctionName
        
        Case "ClearReadingsExecutable"
            Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
        Case "CloseConnectionExecutable"
            Me.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
        Case "OpenConnectionExecutable"
            Me.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable

    End Select
    
End Sub

''' <summary>   Synchronizes the view model with the View. </summary>
''' <remarks>   This is required because the view (ViewModel) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a result, the ViewModel retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' set the view model settings that come from the ViewModel user interface
    This.ViewModel.GpibLanControllerPort = Me.GpibLanControllerPort
    ' This.ViewModel.GpibAddress = Me.PrimaryGpibAddress
    
    ' The initial values must be set for property change to work.
    Me.LastErrorMessage = This.ViewModel.LastErrorMessage
    Me.LastMessage = This.ViewModel.LastMessage
    Me.MeasuredDutNumber = This.ViewModel.MeasuredDutNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    
    Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
    Me.OpenConnectionExecutable = This.ViewModel.OpenConnectionExecutable
    Me.CloseConnectionExecutable = This.ViewModel.CloseConnectionExecutable
    
End Sub

''' <summary>   Handles the property change event of the view model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

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


''' <summary>   Handles the <see cref=cc_isr_Tcp_Scpi.K2700ViewModel"/>.<see cref="cc_isr_Tcp_Scpi.DutMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.DutReadingEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_DutMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.DutReadingEventArgs)

    Const p_procedureName = "m_viewModel_DutMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_DutNumber As Integer: p_DutNumber = a_eventArgs.DutNumber
    
    If UserSheet.ManualSingleToggleValue Then
    
        ' in manual single scan mode, it is expected for the
        ' operator to select a new DUT after the operation is
        ' configured. This DUT needs to be set as the current DUT and the event arguments need
        ' to be modified accordingly.
        
        Dim p_dutCount As Integer
        p_dutCount = This.ViewModel.GetDutCount(This.ViewModel.FrontInputsValue, Me.MaximumDutNumber)
        
        Dim p_details As String
        p_DutNumber = UserView.GetSelectedDutNumber(p_dutCount, p_details)
        
        If p_DutNumber <= 0 Then
            This.ViewModel.OnError p_details
            MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
                "Invalid device under test number"
        Else
            Me.MeasuredDutNumber = p_DutNumber
        End If
    Else
        Me.MeasuredDutNumber = p_DutNumber
    End If
    
    a_eventArgs.Initialize p_DutNumber, a_eventArgs.Reading
    
    ' these values are set by property changes.
    ' Me.MeasuredReading = a_eventArgs.Reading
    ' Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.SetReading This.MeasuredDutNumber, This.MeasuredValue
    
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


