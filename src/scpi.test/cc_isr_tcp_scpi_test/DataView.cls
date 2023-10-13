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

Private Const m_sessionTimeoutCell As String = "B1"
Private Const m_socketAddressCell As String = "C2"
Private Const m_timerIntervalCell As String = "C22"
Private Const m_lastMessageCell As String = "B4"
Private Const m_lastErrorMessageCell As String = "B33"

Private Const m_measurementRow As Integer = 7 ' 2
Private Const m_measurementCol As Integer = 2 ' 23
Private Const m_channelRow As Integer = 7 ' 2
Private Const m_channelCol As Integer = 1 '  24
Private Const m_byRow As Boolean = False ' True
Private Const m_displayIncrement As Integer = 2

Private Type this_

    ViewModel As K2700ViewModel
    DataAcquisitionSheet As DataSheet

    PrimaryGpibAddress As Integer
    GpibLanControllerPort As Long
    MaximumChannelNumber As Integer
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    ReadingOffset As Double
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
    ChannelRow As Integer
    ChannelColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String
    ChannelNumberCaptionPrefix As String

    MeasuredChannelNumber As Integer
    MeasuredReading As String
    MeasuredValue As Double
    
    ClearReadingExecutable As Boolean
    ToggleConnectionValue As Boolean
    ToggleConnectionExecutable As Boolean
    
End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()

    This.ChannelNumberCaptionPrefix = "R"

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
    
    Set This.DataAcquisitionSheet = Nothing

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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Initialize this object. </summary>
Public Function Initialize(ByVal a_viewModel As K2700ViewModel) As DataView

    Const p_procedureName = "Initialize"
    
    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.DataAcquisitionSheet = DataSheet.Initialize()
    
    Set This.ViewModel = a_viewModel
    
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Toggle Connection button change event. </summary>
Public Sub OnToggleConnectionButtonChange()

    Const p_procedureName As String = "OnToggleConnectionButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if already connected or disconnected
    If Me.ToggleConnectionValue = This.ViewModel.Connected Then Exit Sub
    
    If Me.ToggleConnectionValue Then
    
        This.ViewModel.OpenConnectionCommand
        
    Else
    
        This.ViewModel.CloseConnectionCommand
        
    End If
    
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Toggle Connection command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get ToggleConnectionExecutable() As Boolean
    If This.DataAcquisitionSheet Is Nothing Then
        ToggleConnectionExecutable = This.ToggleConnectionExecutable
    Else
        ToggleConnectionExecutable = This.DataAcquisitionSheet.ToggleConnectionExecutable
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Toggle Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ToggleConnectionExecutable(ByVal a_value As Boolean)
    If Me.ToggleConnectionExecutable <> a_value Then
        
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.ToggleConnectionExecutable = a_value
        
        This.ToggleConnectionExecutable = a_value
    
    End If
End Property

''' <summary>   Gets the action state of the Toggle Connection command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ToggleConnectionValue() As Boolean
    If This.DataAcquisitionSheet Is Nothing Then
        ToggleConnectionValue = This.ToggleConnectionValue
    Else
        ToggleConnectionValue = This.DataAcquisitionSheet.ToggleConnectionValue
    End If
End Property

''' <summary>   Sets the action state of the Toggle Connection command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ToggleConnectionValue(ByVal a_value As Boolean)
    If Me.ToggleConnectionValue <> a_value Then
        
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.ToggleConnectionValue = a_value
        
        This.ToggleConnectionValue = a_value
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Clear Reading command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If Not This.DataAcquisitionSheet.ClearReadingsButton Is Nothing Then
        This.DataAcquisitionSheet.ClearReadingsButton.Enabled = a_value
    End If
    This.ClearReadingExecutable = a_value
End Property

''' <summary>   Handles the Clear Messages Button Click Event. </summary>
Public Sub OnClearMessages()
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.LastErrorMessage = VBA.vbNullString
        This.ViewModel.LastMessage = VBA.vbNullString
    End If
End Sub

''' <summary>   Handles the Clear Messages Button Click Event. </summary>
Public Sub OnClearReadings()
    If Not This.ViewModel Is Nothing Then
        This.ViewModel.ClearReadingsCommand
    End If
End Sub

''' <summary>   Gets the primary GPIB address of the instrument. </summary>
''' <value>   [Integer]. </value>
Public Property Get PrimaryGpibAddress() As Integer
    PrimaryGpibAddress = This.PrimaryGpibAddress
End Property

''' <summary>   Sets the primary GPIB address of the instrument. </summary>
''' <param name="a_value">   [Integer] the primary GPIB address. </param>
Public Property Let PrimaryGpibAddress(ByVal a_value As Integer)
    If This.PrimaryGpibAddress <> a_value Then
        This.PrimaryGpibAddress = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.GpibAddress = a_value
    End If
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
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

''' <summary>   Gets the Reading Offset. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadingOffset() As Double
    ReadingOffset = This.ReadingOffset
End Property

''' <summary>   Sets the Reading Offset. </summary>
''' <param name="a_value">   [Double] the Reading Offset value. </value>
Public Property Let ReadingOffset(ByVal a_value As Double)
    If This.ReadingOffset <> a_value Then
        This.ReadingOffset = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.ReadingOffset = a_value
    End If
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <remarks>   True if instrument beep sound on each reading is enabled. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get BeepAllowed() As Boolean
    BeepAllowed = This.BeepAllowed
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <param name="a_value">   [Boolean] instrument beep sound on each reading is enabled. </value>
Public Property Let BeepAllowed(ByVal a_value As Boolean)
    If This.BeepAllowed <> a_value Then _
        This.BeepAllowed = a_value
End Property

''' <summary>   Gets the socket and session timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket and session timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_value As Integer)
    If This.SessionTimeout <> a_value Then
        This.SessionTimeout = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SessionTimeout = a_value
    End If
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SocketAddress(ByVal a_value As String)
    If This.SocketAddress <> a_value Then
        This.SocketAddress = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SocketAddress = a_value
    End If
End Property

''' <summary>   Gets the Timer Interval in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Timer Interval in milliseconds value. </value>
Public Property Let TimerInterval(ByVal a_value As Integer)
    If This.TimerInterval <> a_value Then
        This.TimerInterval = a_value
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.TimerInterval = a_value
    End If
End Property

''' <summary>   Gets the Last Error Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastErrorMessage() As String
    LastErrorMessage = This.LastErrorMessage
End Property

''' <summary>   Sets the Last Error Message. </summary>
''' <param name="a_value">   [String] the last error message value. </value>
Public Property Let LastErrorMessage(ByVal a_value As String)
    If This.LastErrorMessage <> a_value Then
        
        This.LastErrorMessage = a_value
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.LastErrorMessage = a_value
            
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.LastErrorMessage = a_value
            
    End If
End Property

''' <summary>   Gets the Last Message. </summary>
''' <value>   [String]. </value>
Public Property Get LastMessage() As String
    LastMessage = This.LastMessage
End Property

''' <summary>   Sets the Last Message. </summary>
''' <param name="a_value">   [String] the last message value. </value>
Public Property Let LastMessage(ByVal a_value As String)
    If This.LastMessage <> a_value Then
        This.LastMessage = a_value
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.LastMessage = a_value
    
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.LastMessage = a_value
            
    End If
End Property

''' <summary>   Gets the Measured Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasuredChannelNumber() As Integer
    MeasuredChannelNumber = This.MeasuredChannelNumber
End Property

''' <summary>   Sets the Measured Channel Number. </summary>
''' <param name="a_value">   [Integer] the Measured Channel Number value. </value>
Public Property Let MeasuredChannelNumber(ByVal a_value As Integer)
    If This.MeasuredChannelNumber <> a_value Then
        This.MeasuredChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the Measured Reading. </summary>
''' <value>   [String]. </value>
Public Property Get MeasuredReading() As String
    MeasuredReading = This.MeasuredReading
End Property

''' <summary>   Sets the Measured Reading. </summary>
''' <param name="a_value">   [String] the Measured Reading value. </value>
Public Property Let MeasuredReading(ByVal a_value As String)
    If This.MeasuredReading <> a_value Then
        This.MeasuredReading = a_value
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
    If This.DataAcquisitionSheet Is Nothing Then
        MeasurementCellValue = This.MeasurementCellValue
    Else
        MeasurementCellValue = This.DataAcquisitionSheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
    End If
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If Me.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
        
        
    End If
End Property
        
''' <summary>   Gets the channel cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCellValue() As String
    If This.DataAcquisitionSheet Is Nothing Then
        ChannelCellValue = This.ChannelCellValue
    Else
        ChannelCellValue = This.DataAcquisitionSheet.Cells(Me.ChannelRow, Me.ChannelColumn).Value
    End If
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If Me.ChannelCellValue <> a_value Then
        
        This.ChannelCellValue = a_value
        
        If Not This.DataAcquisitionSheet Is Nothing Then _
            This.DataAcquisitionSheet.Cells(Me.ChannelRow, Me.ChannelColumn).Value = a_value
        
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

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_channelNo - 1), m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, m_measurementCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelRow = IIf(a_byRow, m_channelRow + a_increment * (a_channelNo - 1), m_channelRow)
    Me.ChannelColumn = IIf(a_byRow, m_channelCol, m_channelCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelCellValue = This.ChannelNumberCaptionPrefix & VBA.CStr(a_channelNo)
    Me.MeasurementCellValue = VBA.CStr(a_value)
    
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the property change event of the viwe model. </summary>
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
            
        Case "MeasuredChannelNumber"
            Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
        Case "MeasuredReading"
            Me.MeasuredReading = This.ViewModel.MeasuredReading
        Case "MeasuredValue"
            Me.MeasuredValue = This.ViewModel.MeasuredValue
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        
        Case "ClearReadingsExecutable"
            Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
        Case "ToggleConnectionExecutable"
            Me.ToggleConnectionExecutable = This.ViewModel.ToggleConnectionExecutable
        Case "Connected"
            Me.ToggleConnectionValue = This.ViewModel.Connected

    End Select
    
End Sub

''' <summary>   Synchronizes the view mode with the View. </summary>
''' <remarks>   This is required becuase the view (ViewModel) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a resule, the ViewModel retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' set the view model settings that come from the ViewModel user interface
    
    ' set user interface settings
    Me.GpibLanControllerPort = This.DataAcquisitionSheet.GpibLanControllerPort
    Me.PrimaryGpibAddress = This.DataAcquisitionSheet.PrimaryGpibAddress
    Me.MaximumChannelNumber = This.DataAcquisitionSheet.MaxResistanceCount
    Me.SocketAddress = This.DataAcquisitionSheet.SocketAddress
    Me.SessionTimeout = This.DataAcquisitionSheet.SessionTimeout
    Me.TimerInterval = This.DataAcquisitionSheet.TimerInterval
    Me.BeepAllowed = This.DataAcquisitionSheet.BeepAllowed
    
    This.ViewModel.SocketAddress = This.SocketAddress
    This.ViewModel.GpibLanControllerPort = This.GpibLanControllerPort
    This.ViewModel.SessionTimeout = Me.SessionTimeout
    This.ViewModel.TimerInterval = Me.TimerInterval
    This.ViewModel.MaximumChannelNumber = Me.MaximumChannelNumber

    Me.LastErrorMessage = This.ViewModel.LastErrorMessage
    Me.LastMessage = This.ViewModel.LastMessage
    Me.MeasuredChannelNumber = This.ViewModel.MeasuredChannelNumber
    Me.MeasuredReading = This.ViewModel.MeasuredReading
    Me.MeasuredValue = This.ViewModel.MeasuredValue
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.ReadingOffset = This.ViewModel.ReadingOffset
    
    Me.ClearReadingsExecutable = This.ViewModel.ClearReadingsExecutable
    Me.ToggleConnectionExecutable = This.ViewModel.ToggleConnectionExecutable
    
End Sub

''' <summary>   Handles the property change event of the viwe model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_K2700ViewModel_PropertyChanged"
    
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub


''' <summary>   Handles the <see cref=cc_isr_Tcp_Scpi.K2700ViewModel"/>.<see cref="cc_isr_Tcp_Scpi.ChannelMeasured"/> event. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.ChannelReadingEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_ChannelMeasured(ByVal a_eventArgs As cc_isr_Tcp_Scpi.ChannelReadingEventArgs)

    Const p_procedureName = "m_viewModel_ChannelMeasured"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Me.MeasuredReading = a_eventArgs.Reading
    Me.MeasuredChannelNumber = a_eventArgs.ChannelNumber
        
    Dim p_value As Double
    Dim p_details As String
    
    If VBA.vbNullString = a_eventArgs.Reading Then
        p_value = -1#
    Else
        If Not cc_isr_Core.StringExtensions.TryParseDouble(a_eventArgs.Reading, p_value, p_details) Then
            p_value = -1#
        End If
    End If
    Me.MeasuredValue = p_value
    
    Me.SetReading This.MeasuredChannelNumber, This.MeasuredValue
    
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

