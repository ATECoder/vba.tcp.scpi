VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAcquisitionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The Data Acquisition View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sessionTimeout As String = "B1"
Private Const m_inputAddress As String = "B2"
Private Const m_inputPort As String = "B3"
Private Const m_socketAddress As String = "C2"
Private Const m_socketId As String = "D2"
Private Const m_connected As String = "E2"

Private Const m_timerInterval As String = "C22"

Private Const m_lastMessage As String = "B4"
Private Const m_lastErrorMessage As String = "B33"

Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 23
Private Const m_channelRow As Integer = 2
Private Const m_channelCol As Integer = 24
Private Const m_byRow As Boolean = True
Private Const m_displayIncrement As Integer = 1

' Controls

Private WithEvents ClearReadingsButton As MSForms.CommandButton
Attribute ClearReadingsButton.VB_VarHelpID = -1
Private WithEvents AutoScanToggleButton As MSForms.ToggleButton
Attribute AutoScanToggleButton.VB_VarHelpID = -1
Private WithEvents AutoSingleToggleButton As MSForms.ToggleButton
Attribute AutoSingleToggleButton.VB_VarHelpID = -1
Private WithEvents ManualScanToggleButton As MSForms.ToggleButton
Attribute ManualScanToggleButton.VB_VarHelpID = -1
Private WithEvents ManualSingleToggleButton As MSForms.ToggleButton
Attribute ManualSingleToggleButton.VB_VarHelpID = -1

Private Type this_

    AutoIncrementChannelNoEnabled As Boolean
    ReadingOffset As Double
    FrontInputsRequired As Boolean
    FrontInputsValue As Boolean
    FrontInputsCaption  As String
    ReadingBeepEnabled As Boolean
    SelectedChannelNumber As Integer
    SingleReadEnabled As Boolean
    
    MeasurementRow As Integer
    MeasurementColumn As Integer
    ChannelRow As Integer
    ChannelColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String

    Observer As K2700Observer

End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Sub Class_Initialize()
    Set This.Observer = K2700Observer
End Sub

Private Sub Class_Terminate()
    Set This.Observer = Nothing
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Auto Scan toggle button. </summary>
''' <remarks>   In auto scan we have the following settings:
''' Auto Increment channel number: True;
''' Front Inputs Required: False;
''' Single Readings: False;
''' Triggering Mode: Immediate.
''' </remarks>
Private Sub AutoScanToggleButton_Change()

    Const p_procedureName As String = "AutoScanToggleButton_Change"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If AutoScanToggleButton.Value Then
    
        Me.AutoIncrementChannelNoEnabled = True
        Me.FrontInputsRequired = False
        Me.SingleReadEnabled = False
        
        ' next set immediate triggering mode.
        This.Observer.OnImmediateTriggerOptionChange AutoScanToggleButton.Value
        
        ' next the measurement
        This.Observer.OnMeasure
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' toggle the button state
    AutoScanToggleButton.Value = False
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    This.Observer.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Public Property Let AutoScanExecutable(ByVal a_value As Boolean)
    If Not AutoScanToggleButton Is Nothing Then
        AutoScanToggleButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Auto Single toggle button. </summary>
''' <remarks>   In auto single we have the following settings:
''' Auto Increment channel number: False;
''' Front Inputs Required: False;
''' Single Readings: True;
''' Triggering Mode: Immediate.
''' </remarks>
Private Sub AutoSingleToggleButton_Change()

    Const p_procedureName As String = "AutoScanToggleButton_Change"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If AutoScanToggleButton.Value Then
    
        Me.AutoIncrementChannelNoEnabled = False
        Me.FrontInputsRequired = False
        Me.SingleReadEnabled = True
        
        ' next set immediate triggering mode.
        This.Observer.OnImmediateTriggerOptionChange AutoScanToggleButton.Value
        
        ' next the measurement
        This.Observer.OnMeasure
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' toggle the button state
    AutoSingleToggleButton.Value = False
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' report the last error.
    This.Observer.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Public Property Let AutoSingleExecutable(ByVal a_value As Boolean)
    If Not AutoSingleToggleButton Is Nothing Then
        AutoSingleToggleButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Manual Scan toggle button. </summary>
''' <remarks>   In manual scan we have the following settings:
''' Auto Increment channel number: True;
''' Front Inputs Required: True;
''' Single Readings: False;
''' Triggering Mode: External.
''' </remarks>
Private Sub ManualScanToggleButton_Change()

    Const p_procedureName As String = "ManualScanToggleButton_Change"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If ManualScanToggleButton.Value Then
    
        Me.AutoIncrementChannelNoEnabled = True
        Me.FrontInputsRequired = True
        Me.SingleReadEnabled = False
        
        ' next set external triggering triggering mode.
        This.Observer.OnExternalTriggerOptionChange ManualScanToggleButton.Value
        
        ' next start external trigger monitoring
        ' this should wait for monitoring to commence or raise an error.
        This.Observer.OnMonitorExtTrigToggleButtonChange True
        
    Else
    
        ' stop external trigger monitoring
        ' this should wait for monitoring to stop
        This.Observer.OnMonitorExtTrigToggleButtonChange False
    
        ' this should stop external triggering
        This.Observer.OnImmediateTriggerOptionChange False
        
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
    This.Observer.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    This.Observer.OnMonitorExtTrigToggleButtonChange False

    ' this should stop external triggering
    This.Observer.OnImmediateTriggerOptionChange False
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Public Property Let ManualScanExecutable(ByVal a_value As Boolean)
    If Not ManualScanToggleButton Is Nothing Then
        ManualScanToggleButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Manual Single toggle button. </summary>
''' <remarks>   In manual Single we have the following settings:
''' Auto Increment channel number: False;
''' Front Inputs Required = True;
''' Single Readings: True;
''' Triggering Mode: External.
''' </remarks>
Private Sub ManualSingleToggleButton_Change()

    Const p_procedureName As String = "ManualSingleToggleButton_Change"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If ManualSingleToggleButton.Value Then
    
        Me.AutoIncrementChannelNoEnabled = False
        Me.FrontInputsRequired = True
        Me.SingleReadEnabled = True
        
        ' next set external triggering triggering mode.
        This.Observer.OnExternalTriggerOptionChange ManualSingleToggleButton.Value
        
        ' next start external trigger monitoring
        ' this should wait for monitoring to commence or raise an error.
        This.Observer.OnMonitorExtTrigToggleButtonChange True
        
    Else
    
        ' stop external trigger monitoring
        ' this should wait for monitoring to stop
        This.Observer.OnMonitorExtTrigToggleButtonChange False
    
        ' this should stop external triggering
        This.Observer.OnImmediateTriggerOptionChange False
        
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
    This.Observer.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    This.Observer.OnMonitorExtTrigToggleButtonChange False

    ' this should stop external triggering
    This.Observer.OnImmediateTriggerOptionChange False
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Public Property Let ManualSingleExecutable(ByVal a_value As Boolean)
    If Not ManualSingleToggleButton Is Nothing Then
        ManualSingleToggleButton.Enabled = a_value
    End If
End Property

Private Sub ClearReadingsButton_Click()
    This.Observer.OnClearReadings
End Sub

Public Property Let ClearReadingsExecutable(ByVal a_value As Boolean)
    If Not ClearReadingsButton Is Nothing Then
        ClearReadingsButton.Enabled = a_value
    End If
End Property

''' <summary>   Gets the Auto Increment Channel Number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementChannelNoEnabled() As Boolean
    AutoIncrementChannelNoEnabled = This.AutoIncrementChannelNoEnabled
End Property

''' <summary>   Sets the Auto Increment Channel Number Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment Channel Number Enabled value. </value>
Public Property Let AutoIncrementChannelNoEnabled(ByVal a_value As Boolean)
    If This.AutoIncrementChannelNoEnabled <> a_value Then
        This.AutoIncrementChannelNoEnabled = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.AutoIncrementChannelNoEnabled = a_value
    End If
End Property

''' <summary>   Gets the External Trig Monitoring Enabled condition. </summary>
''' <remarks>   When true, a Stop request is sent to the timer so that the timer
'''             event is not handled and the timer is Stopd on the next event. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ExternalTrigMonitoringEnabled() As Boolean
    ExternalTrigMonitoringEnabled = ManualScanToggleButton.Value
End Property

''' <summary>   Sets the External Trig Monitoring Enabled condition. </summary>
''' <param name="a_value">   [Boolean] the External Trig Monitoring Enabled value. </value>
Public Property Let ExternalTrigMonitoringEnabled(ByVal a_value As Boolean)
    If Me.ExternalTrigMonitoringEnabled <> a_value Then
        ManualScanToggleButton.Value = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.ExternalTrigMonitoringEnabled = a_value
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
        If Not This.Observer Is Nothing Then _
                This.Observer.ReadingOffset = a_value
    End If
End Property

''' <summary>   Gets the Front Inputs Required. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputsRequired() As Boolean
    FrontInputsRequired = This.FrontInputsRequired
End Property

''' <summary>   Sets the Front Inputs Required. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are Required. </value>
Public Property Let FrontInputsRequired(ByVal a_value As Boolean)
    If This.FrontInputsRequired <> a_value Then
        This.FrontInputsRequired = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.FrontInputsRequired = a_value
    End If
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <remarks>   True if instrument beep sound on each reading is enabled. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get ReadingBeepEnabled() As Boolean
    ReadingBeepEnabled = This.ReadingBeepEnabled
End Property

''' <summary>   Gets the instrument beep sound on each reading is enabled. </summary>
''' <param name="a_value">   [Boolean] instrument beep sound on each reading is enabled. </value>
Public Property Let ReadingBeepEnabled(ByVal a_value As Boolean)
    If This.ReadingBeepEnabled <> a_value Then
        This.ReadingBeepEnabled = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.ReadingBeepEnabled = a_value
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
''' <value>   [String]. </value>
Public Property Get FrontInputsCaption() As String
    FrontInputsCaption = This.FrontInputsCaption
End Property

''' <summary>   Sets the Front Inputs Caption. </summary>
''' <param name="a_value">   [String] the Front Inputs Caption. </value>
Public Property Let FrontInputsCaption(ByVal a_value As String)
    If This.FrontInputsCaption <> a_value Then
        This.FrontInputsCaption = a_value
    End If
End Property


''' <summary>   Gets the Selected Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedChannelNumber() As Integer
    SelectedChannelNumber = This.SelectedChannelNumber
End Property

''' <summary>   Sets the Selected Channel Number. </summary>
''' <param name="a_value">   [Integer] the Selected Channel Number value. </value>
Public Property Let SelectedChannelNumber(ByVal a_value As Integer)
    If This.SelectedChannelNumber <> a_value Then
        This.SelectedChannelNumber = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.SelectedChannelNumber = a_value
    End If
End Property

''' <summary>   Gets the Single Read Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get SingleReadEnabled() As Boolean
    SingleReadEnabled = This.SingleReadEnabled
End Property

''' <summary>   Sets the Single Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Single Read Enabled value. </value>
Public Property Let SingleReadEnabled(ByVal a_value As Boolean)
    If This.SingleReadEnabled <> a_value Then
        This.SingleReadEnabled = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.SingleReadEnabled = a_value
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
    MeasurementCellValue = This.MeasurementCellValue
    ' MeasurementCellValue = Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If This.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        ' Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
    End If
End Property
        
''' <summary>   Gets the channel cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCellValue() As String
    ChannelCellValue = This.ChannelCellValue
    ' ChannelCellValue = Cells(Me.ChannelRow, Me.ChannelColumn).Value
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If This.ChannelCellValue <> a_value Then
        This.ChannelCellValue = a_value
        ' Cells(Me.ChannelRow, Me.ChannelColumn).Value = a_value
    End If
End Property

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_format">      [String] the format string for formatting the value. </param>
Public Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double, a_format As String)
    SetReading_ a_channelNo, a_value, m_byRow, m_displayIncrement, a_format
End Sub

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_byRow">       [Boolean] true if the next reading goes to the next row. </param>
''' <param name="a_increment">   [Integer] the row or column increment. </param>
''' <param name="a_format">      [String] the format string for formatting the value. </param>
Private Sub SetReading_(ByVal a_channelNo As Integer, ByVal a_value As Double, _
    ByVal a_byRow As Boolean, ByVal a_increment As Integer, a_format As String)

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_channelNo - 1), m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, m_measurementCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelRow = IIf(a_byRow, m_channelRow + a_increment * (a_channelNo - 1), m_channelRow)
    Me.ChannelColumn = IIf(a_byRow, m_channelCol, m_channelCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelCellValue = VBA.CStr(a_channelNo)
    Me.MeasurementCellValue = VBA.Format$(a_value, a_format)
    
End Sub

