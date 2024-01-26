VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The User View for the view model. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private WithEvents m_viewModel As K2700ViewModel
Attribute m_viewModel.VB_VarHelpID = -1

Private Type this_

    ViewModel As K2700ViewModel
    UserSheet As UserSheet
    DataSheet As DataSheet
    DesignMode As Boolean
    
    Connected As Boolean
    
    AutoIncrementDutNumberEnabled As Boolean
    FrontInputsRequired As Boolean
    SingleReadEnabled As Boolean
    
    ReadingOffset As Double
    
    MaximumDutNumber As Integer
    SelectedDutNumber As Integer
    SelectedDutNumberCaption As String
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    Measuring As Boolean
    
    LastCommandDetails As String
    LastCommandOutcome As Boolean
    
    AutoScanToggleExecutable As Boolean
    AutoSingleToggleExecutable As Boolean
    ManualScanToggleExecutable As Boolean
    ManualSingleToggleExecutable As Boolean

    AutoScanToggleValue As Boolean
    AutoSingleToggleValue As Boolean
    ManualScanToggleValue As Boolean
    ManualSingleToggleValue As Boolean
    
    AutoSampleFrontInputs As Boolean

End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
    This.AutoSampleFrontInputs = False
End Sub

''' <summary>   Handles the Class Terminal event. </summary>
Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.UserSheet = Nothing
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
''' <value>   [<see cref="UserView"/>]. </value>
Public Property Get Instance() As UserView
    Set Instance = Me
End Property

''' <summary>   Initialize this object. </summary>
''' <returns>   [<see cref="UserView"/>]. </returns>
Public Function Initialize(ByVal a_viewModel As K2700ViewModel, ByVal a_dataSheet As DataSheet) As UserView

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.ViewModel = a_viewModel
    
    Set This.UserSheet = UserSheet.Initialize()
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

''' <summary>   Gets the selected state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property

''' <summary>   Sets the connected state. </summary>
''' <param name="a_value">   [Boolean] the connected state. </param>
Public Property Let Connected(ByVal a_value As Boolean)
    This.Connected = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' User interface control
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the design mode for unit testing and restoration of state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get DesignMode() As Boolean
    If Not This.UserSheet Is Nothing Then _
        This.DesignMode = This.UserSheet.DesignMode
    DesignMode = This.DesignMode
End Property

''' <summary>   Sets the design mode for unit testing and restoration of state. </summary>
''' <param name="a_value">   [Boolean] true if in design mode. </param>
Public Property Let DesignMode(ByVal a_value As Boolean)
    If Me.DesignMode <> a_value Then
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.DesignMode = a_value
        This.DesignMode = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Auto Scan toggle button. </summary>
''' <remarks>   In auto scan we have the following settings:
''' Auto Increment DUT number: True;
''' Front Inputs Required: False;
''' Single Readings: False;
''' Triggering Mode: Immediate.
''' </remarks>
Public Sub OnAutoScanToggleButtonChange()

    Const p_procedureName As String = "OnAutoScanToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.AutoScanToggleValue And This.ViewModel.ImmediateTriggerOptionExecutable And _
        Not This.Measuring Then
        
        This.UserSheet.DisplayStartDateTime
        
        Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
        Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
        p_mode.AutoIncrement = True
        p_mode.BeepEnabled = False
        p_mode.FrontInputs = This.DataSheet.ImmediateFrontInputsRequired
        p_mode.MaximumDutCount = Me.MaximumDutNumber
        p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
        
        Me.SelectedDutNumber = 1
        p_mode.DutNumber = Me.SelectedDutNumber
        
        p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
        p_mode.ReadingOffset = Me.ReadingOffset
        p_mode.SenseFunction = IIf(p_mode.FrontInputs, This.DataSheet.FrontInputsSenseFunctionName, _
            This.DataSheet.RearInputsSenseFunctionName)
        p_mode.SingleRead = False
        p_mode.TimerInterval = This.DataSheet.TimerInterval
        
        Me.LastCommandOutcome = This.ViewModel.StartMeasureCommand(p_mode, Me.LastCommandDetails)
    
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
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get AutoScanToggleExecutable() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.AutoScanToggleExecutable = This.UserSheet.AutoScanToggleExecutable
    End If
    AutoScanToggleExecutable = This.AutoScanToggleExecutable

End Property

''' <summary>   Sets the execution (Can Execute) status of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoScanToggleExecutable(ByVal a_value As Boolean)
    If Me.AutoScanToggleExecutable <> a_value Then
        
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.AutoScanToggleExecutable = a_value
        
        This.AutoScanToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoScanToggleValue() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.AutoScanToggleValue = This.UserSheet.AutoScanToggleValue
    End If
    AutoScanToggleValue = This.AutoScanToggleValue

End Property

''' <summary>   Sets the action state of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoScanToggleValue(ByVal a_value As Boolean)
    
    ' allow value changes only if the control is enabled.
    ' this prevent value changes while the measurement mode is changing.
    If Me.AutoScanToggleExecutable Then
    
        If Me.AutoScanToggleValue <> a_value Then
        
            If Not This.UserSheet Is Nothing Then _
                This.UserSheet.AutoScanToggleValue = a_value
    
            This.AutoScanToggleValue = a_value
        End If
    
    End If

End Property

''' <summary>   Handles the Auto Single toggle button. </summary>
''' <remarks>   In auto single we have the following settings:
''' Auto Increment DUT number: False;
''' Front Inputs Required: False;
''' Single Readings: True;
''' Triggering Mode: Immediate.
''' </remarks>
Public Sub OnAutoSingleToggleButtonChange()

    Const p_procedureName As String = "OnAutoSingleToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.ViewModel Is Nothing Then Exit Sub

    ' exit if the view model is already in auto single mode and the button state is
    ' already the same as the measuring state
    If This.MeasurementMode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate And _
        This.SingleReadEnabled And Me.AutoSingleToggleValue = Me.Measuring Then
        Exit Sub
    End If

    If Me.AutoSingleToggleValue And This.ViewModel.ImmediateTriggerOptionExecutable And _
        Not Me.Measuring Then
    
        Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
        Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
        p_mode.AutoIncrement = False
        p_mode.BeepEnabled = False
        p_mode.FrontInputs = This.DataSheet.ImmediateFrontInputsRequired
        p_mode.MaximumDutCount = Me.MaximumDutNumber
        p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
        
        Dim p_details As String
        p_mode.DutNumber = Me.GetSelectedDutNumber(p_mode.DutCount, p_details)
        If 0 < p_mode.DutNumber Then
        
            p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate
            p_mode.ReadingOffset = Me.ReadingOffset
            p_mode.SenseFunction = IIf(p_mode.FrontInputs, This.DataSheet.FrontInputsSenseFunctionName, _
                This.DataSheet.RearInputsSenseFunctionName)
            p_mode.SingleRead = True
            p_mode.TimerInterval = This.DataSheet.TimerInterval
    
            This.UserSheet.DisplayStartDateTime
            Me.LastCommandOutcome = This.ViewModel.StartMeasureCommand(p_mode, Me.LastCommandDetails)
        
        Else
            Me.AutoSingleToggleValue = False
            MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
                "Invalid device under test number"
        End If
        
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
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get AutoSingleToggleExecutable() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.AutoSingleToggleExecutable = This.UserSheet.AutoSingleToggleExecutable
    End If
    AutoSingleToggleExecutable = This.AutoSingleToggleExecutable

End Property

''' <summary>   Sets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoSingleToggleExecutable(ByVal a_value As Boolean)
    
    If Me.AutoSingleToggleExecutable <> a_value Then
        
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.AutoSingleToggleExecutable = a_value

        This.AutoSingleToggleExecutable = a_value
    End If

End Property

''' <summary>   Gets the action state of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoSingleToggleValue() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.AutoSingleToggleValue = This.UserSheet.AutoSingleToggleValue
    End If
    AutoSingleToggleValue = This.AutoSingleToggleValue

End Property

''' <summary>   Sets the action state of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoSingleToggleValue(ByVal a_value As Boolean)
    
    ' allow value changes only if the control is enabled.
    ' this prevent value changes while the measurement mode is changing.
    If Me.AutoSingleToggleExecutable Then
    
        If Me.AutoSingleToggleValue <> a_value Then
            
            If Not This.UserSheet Is Nothing Then _
                This.UserSheet.AutoSingleToggleValue = a_value
    
            This.AutoSingleToggleValue = a_value
        End If
    
    End If
    
End Property

''' <summary>   Stops external triggering monitoring. </summary>
Public Sub StopManualTriggering()

    Const p_procedureName As String = "StopManualTriggering"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    ' stop external trigger monitoring
    
    ' this should waits for monitoring to stop
    This.ViewModel.StopMonitoringExternalTriggersCommand
   
           
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
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.StopMonitoringExternalTriggersCommand

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the Manual Scan toggle button. </summary>
''' <remarks>   In manual scan we have the following settings:
''' Auto Increment DUT number: True;
''' Front Inputs Required: True;
''' Single Readings: False;
''' Triggering Mode: External.
''' </remarks>
Public Sub OnManualScanToggleButtonChange()

    Const p_procedureName As String = "OnManualScanToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.ManualScanToggleValue And This.ViewModel.ExternalTriggerOptionExecutable And _
        Not This.Measuring Then
    
        This.UserSheet.DisplayStartDateTime
        
        Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
        Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
        p_mode.AutoIncrement = True
        p_mode.BeepEnabled = True
        p_mode.FrontInputs = This.DataSheet.ExternalFrontInputsRequired
        p_mode.DutNumber = 1
        p_mode.MaximumDutCount = Me.MaximumDutNumber
        p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
        p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
        p_mode.ReadingOffset = Me.ReadingOffset
        p_mode.SenseFunction = IIf(p_mode.FrontInputs, This.DataSheet.FrontInputsSenseFunctionName, _
            This.DataSheet.RearInputsSenseFunctionName)
        p_mode.SingleRead = False
        p_mode.TimerInterval = This.DataSheet.TimerInterval

        Me.LastCommandOutcome = This.ViewModel.StartMonitoringCommand(p_mode, Me.LastCommandDetails)

    ElseIf Not Me.ManualScanToggleValue And This.ViewModel.StopMonitoringExecutable And _
        This.Measuring Then
    
        Me.StopManualTriggering
        
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
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.StopMonitoringExternalTriggersCommand

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Gets the execution (Can Execute) status of the Manual Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get ManualScanToggleExecutable() As Boolean
    If Not This.UserSheet Is Nothing Then
        This.ManualScanToggleExecutable = This.UserSheet.ManualScanToggleExecutable
    End If
    ManualScanToggleExecutable = This.ManualScanToggleExecutable
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualScanToggleExecutable(ByVal a_value As Boolean)
    If Me.ManualScanToggleExecutable <> a_value Then
        
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.ManualScanToggleExecutable = a_value

        This.ManualScanToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Manual Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualScanToggleValue() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.ManualScanToggleValue = This.UserSheet.ManualScanToggleValue
    End If
    ManualScanToggleValue = This.ManualScanToggleValue

End Property

''' <summary>   Sets the action state of the Manual Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualScanToggleValue(ByVal a_value As Boolean)
    
    ' allow value changes only if the control is enabled.
    ' this prevent value changes while the measurement mode is changing.
    If Me.ManualScanToggleExecutable Then
    
        If Me.ManualScanToggleValue <> a_value Then
            
            If Not This.UserSheet Is Nothing Then _
                This.UserSheet.ManualScanToggleValue = a_value
    
            This.ManualScanToggleValue = a_value
        End If
    
    End If
    
End Property

''' <summary>   Handles the Manual Single toggle button. </summary>
''' <remarks>   In manual Single we have the following settings:
''' Auto Increment DUT number: False;
''' Front Inputs Required = True;
''' Single Readings: True;
''' Triggering Mode: External.
''' </remarks>
Public Sub OnManualSingleToggleButtonChange()

    Const p_procedureName As String = "OnManualSingleToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.ManualSingleToggleValue And This.ViewModel.ExternalTriggerOptionExecutable And _
        Not This.Measuring Then
    
        Dim p_mode As cc_isr_Tcp_Scpi.MeasureMode
        Set p_mode = cc_isr_Tcp_Scpi.Factory.NewMeasureMode
        p_mode.AutoIncrement = False
        p_mode.BeepEnabled = True
        p_mode.FrontInputs = This.DataSheet.ExternalFrontInputsRequired
        p_mode.MaximumDutCount = Me.MaximumDutNumber
        p_mode.DutCount = This.ViewModel.GetDutCount(p_mode.FrontInputs, p_mode.MaximumDutCount)
        
        Dim p_details As String
        p_mode.DutNumber = Me.GetSelectedDutNumber(p_mode.DutCount, p_details)
        If 0 < p_mode.DutNumber Then
        
            p_mode.Mode = cc_isr_Tcp_Scpi.MeasurementModeOption.External
            p_mode.ReadingOffset = Me.ReadingOffset
            p_mode.SenseFunction = IIf(p_mode.FrontInputs, This.DataSheet.FrontInputsSenseFunctionName, _
                This.DataSheet.RearInputsSenseFunctionName)
            p_mode.SingleRead = True
            p_mode.TimerInterval = This.DataSheet.TimerInterval
    
            This.UserSheet.DisplayStartDateTime
            Me.LastCommandOutcome = This.ViewModel.StartMonitoringCommand(p_mode, Me.LastCommandDetails)
    
        Else
            Me.ManualSingleToggleValue = False
            MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
                "Invalid device under test number"
        End If
    
    ElseIf Not Me.ManualSingleToggleValue And This.ViewModel.StopMonitoringExecutable And _
        This.Measuring Then
    
        Me.StopManualTriggering
            
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
        This.ViewModel.OnError cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' on error, stop monitoring
    On Error Resume Next
    
    ' stop external trigger monitoring
    ' this should wait for monitoring to stop
    If Not This.ViewModel Is Nothing Then _
        This.ViewModel.StopMonitoringExternalTriggersCommand

    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Gets the execution (Can Execute) status of the Manual Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get ManualSingleToggleExecutable() As Boolean
    If Not This.UserSheet Is Nothing Then
        This.ManualSingleToggleExecutable = This.UserSheet.ManualSingleToggleExecutable
    End If
    ManualSingleToggleExecutable = This.ManualSingleToggleExecutable
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualSingleToggleExecutable(ByVal a_value As Boolean)
    If Me.ManualSingleToggleExecutable <> a_value Then
    
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.ManualSingleToggleExecutable = a_value

        This.ManualSingleToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Manual Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualSingleToggleValue() As Boolean
    
    If Not This.UserSheet Is Nothing Then
        This.ManualSingleToggleValue = This.UserSheet.ManualSingleToggleValue
    End If
    ManualSingleToggleValue = This.ManualSingleToggleValue

End Property

''' <summary>   Sets the action state of the Manual Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualSingleToggleValue(ByVal a_value As Boolean)
    
    ' allow value changes only if the control is enabled.
    ' this prevent value changes while the measurement mode is changing.
    If Me.ManualSingleToggleExecutable Then
    
        If Me.ManualSingleToggleValue <> a_value Then
        
            If Not This.UserSheet Is Nothing Then _
                This.UserSheet.ManualSingleToggleValue = a_value
    
            This.ManualSingleToggleValue = a_value
        End If
    
    End If
    
End Property

''' <summary>   Gets the Auto Increment DUT number Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementDutNumberEnabled() As Boolean
    AutoIncrementDutNumberEnabled = This.AutoIncrementDutNumberEnabled
End Property

''' <summary>   Sets the Auto Increment DUT number Enabled. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment DUT number Enabled value. </value>
Public Property Let AutoIncrementDutNumberEnabled(ByVal a_value As Boolean)
    This.AutoIncrementDutNumberEnabled = a_value
End Property

''' <summary>   Gets the Reading Offset. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ReadingOffset() As Double
    If Not This.UserSheet Is Nothing Then _
        This.ReadingOffset = This.UserSheet.InitialResistance
    ReadingOffset = This.ReadingOffset
End Property

''' <summary>   Sets the Reading Offset. </summary>
''' <param name="a_value">   [Double] the Reading Offset value. </value>
Public Property Let ReadingOffset(ByVal a_value As Double)
    This.ReadingOffset = a_value
End Property

''' <summary>   Gets the condition for using front inputs for auto sample. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoSampleFrontInputs() As Boolean
    AutoSampleFrontInputs = This.AutoSampleFrontInputs
End Property

''' <summary>   Sets the condition for using front inputs for auto sample. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are used for auto sampling. </value>
Public Property Let AutoSampleFrontInputs(ByVal a_value As Boolean)
    This.AutoSampleFrontInputs = a_value
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
    This.FrontInputsRequired = a_value
End Property

''' <summary>   Gets the maximum DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutNumber() As String
    If Not This.DataSheet Is Nothing Then _
        This.MaximumDutNumber = This.DataSheet.MaximumDutNumber
    MaximumDutNumber = This.MaximumDutNumber
End Property

''' <summary>   Gets the Selected DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedDutNumber() As Integer
    SelectedDutNumber = This.SelectedDutNumber
End Property

''' <summary>   Sets the Selected DUT number. </summary>
''' <param name="a_value">   [Integer] the Selected DUT number value. </value>
Public Property Let SelectedDutNumber(ByVal a_value As Integer)
    
    If This.SelectedDutNumber <> a_value Then
        
        This.SelectedDutNumber = a_value
        
        ' emulate selecting a DUT number if within range.
        If a_value > 0 And a_value <= Me.MaximumDutNumber Then
            Dim p_caption As String
            p_caption = This.DataSheet.DutNumberCaptionPrefix & VBA.CStr(a_value)
            This.SelectedDutNumberCaption = p_caption
            If Not This.UserSheet Is Nothing Then
                This.UserSheet.ActivateDutNumberCaption a_value
            End If
        End If
        
    End If

End Property

''' <summary>   Gets the Selected DUT number Caption. </summary>
''' <value>   [String]. </value>
Public Property Get SelectedDutNumberCaption() As String
    If Not This.UserSheet Is Nothing Then _
        This.SelectedDutNumberCaption = This.UserSheet.SelectedDutNumberCaption
    SelectedDutNumberCaption = This.SelectedDutNumberCaption
End Property

''' <summary>   Gets a DUT number from the active DUT number caption. </summary>
''' <param name="a_dutCount">   [Integer] The number of DUTs that is to be tested, which
'''                             also set the top (max) DUT Number of this test. </param>
''' <param name="a_details">    [Out, String] Details the failure information if any. </param>
''' <returns>   [Integer] Returns 0 if invalid DUT number caption was selected. </returns>
Public Function GetSelectedDutNumber(ByVal a_dutCount As Integer, _
    ByRef a_details As String) As Integer

    Dim p_number As Integer
    p_number = This.ViewModel.TryParseDutNumberCaption(Me.SelectedDutNumberCaption, _
        Me.DutNumberCaptionPrefix, Me.DutTitle, a_dutCount, a_details)
    This.SelectedDutNumber = p_number
    GetSelectedDutNumber = p_number
    
End Function


''' <summary>   Gets the Single Read Enabled. </summary>
''' <remarks>   Indicates that front inputs are Value for this mode. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get SingleReadEnabled() As Boolean
    SingleReadEnabled = This.SingleReadEnabled
End Property

''' <summary>   Sets the Single Read Enabled. </summary>
''' <param name="a_value">   [Boolean] the Single Read Enabled value. </value>
Public Property Let SingleReadEnabled(ByVal a_value As Boolean)
    This.SingleReadEnabled = a_value
End Property

''' <summary>   Gets the measuring state. True when measurements are collected. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Measuring() As Boolean
    Measuring = This.Measuring
End Property

''' <summary>   Sets the measuring state. True when measurements are collected. </summary>
''' <param name="a_value">   [Boolean] the measuring state; true when measuring is active. </value>
Public Property Let Measuring(ByVal a_value As Boolean)
    This.Measuring = a_value
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get MeasurementMode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    MeasurementMode = This.MeasurementMode
End Property

''' <summary>   Sets the measurement mode. </summary>
''' <remarks>   This property change is used to toggle the buttons if not toggles. </remarks>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Let MeasurementMode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    This.MeasurementMode = a_value
End Property

''' <summary>   Returns true if the instrument has scan cards. </summary>
Public Function HasScanCards() As Boolean

    If This.Connected Then
        HasScanCards = This.ViewModel.K2700.RouteSystem.ChannelCount > 0
    Else
        HasScanCards = False
    End If
    
End Function

''' <summary>   Returns true if auto measurement is allowed. </summary>
Public Function CanAutoMeasure() As Boolean
    CanAutoMeasure = Me.HasScanCards Or This.AutoSampleFrontInputs
End Function

''' <summary>   Handles the change of measurement state. </summary>
''' <remarks>   This enables or disables the buttons and toggles the buttons
''' state if not toggled. </remarks>
Friend Sub OnMeasurementStateChanged()

    If cc_isr_Tcp_Scpi.MeasurementModeOption.None = This.MeasurementMode Or _
        cc_isr_Tcp_Scpi.MeasurementModeOption.Continuous = This.MeasurementMode Then
        If Me.CanAutoMeasure Then
            Me.AutoScanToggleExecutable = This.Connected
            Me.AutoSingleToggleExecutable = This.Connected
        Else
            Me.AutoScanToggleExecutable = False
            Me.AutoSingleToggleExecutable = False
        End If
        Me.AutoScanToggleValue = False
        Me.AutoSingleToggleValue = False
        Me.ManualScanToggleExecutable = This.Connected
        Me.ManualSingleToggleExecutable = This.Connected
        Me.ManualScanToggleValue = False
        Me.ManualSingleToggleValue = False
    ElseIf cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate = This.MeasurementMode Then
        If Me.CanAutoMeasure Then
            Me.AutoScanToggleExecutable = This.Connected And (Not (This.SingleReadEnabled And This.Measuring))
            Me.AutoSingleToggleExecutable = This.Connected And (This.SingleReadEnabled Or Not This.Measuring)
            Me.AutoScanToggleValue = Not This.SingleReadEnabled And This.Measuring
            Me.AutoSingleToggleValue = This.SingleReadEnabled And This.Measuring
        Else
            Me.AutoScanToggleExecutable = False
            Me.AutoSingleToggleExecutable = False
            Me.AutoScanToggleValue = False
            Me.AutoSingleToggleValue = False
        End If
        Me.ManualScanToggleExecutable = This.Connected And Not This.Measuring
        Me.ManualSingleToggleExecutable = This.Connected And Not This.Measuring
        Me.ManualScanToggleValue = False
        Me.ManualSingleToggleValue = False
    ElseIf cc_isr_Tcp_Scpi.MeasurementModeOption.External = This.MeasurementMode Or _
           cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring = This.MeasurementMode Then
        If Me.CanAutoMeasure Then
            Me.AutoScanToggleExecutable = This.Connected And Not This.Measuring
            Me.AutoSingleToggleExecutable = This.Connected And Not This.Measuring
        Else
            Me.AutoScanToggleExecutable = False
            Me.AutoSingleToggleExecutable = False
        End If
        Me.AutoScanToggleValue = False
        Me.AutoSingleToggleValue = False
        Me.ManualScanToggleExecutable = This.Connected And (Not (This.SingleReadEnabled And This.Measuring))
        Me.ManualSingleToggleExecutable = This.Connected And (This.SingleReadEnabled Or Not This.Measuring)
        Me.ManualScanToggleValue = Not This.SingleReadEnabled And This.Measuring
        Me.ManualSingleToggleValue = This.SingleReadEnabled And This.Measuring
    End If

End Sub

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

''' <summary>   Gets the Last Command Details. </summary>
''' <value>   [String]. </value>
Public Property Get LastCommandDetails() As String
    LastCommandDetails = This.LastCommandDetails
End Property

''' <summary>   Sets the Last Command Details. </summary>
''' <param name="a_value">   [String] the last Command Details value. </value>
Public Property Let LastCommandDetails(ByVal a_value As String)
    This.LastCommandDetails = a_value
End Property

''' <summary>   Gets the Last Command Outcome. </summary>
''' <value>   [Boolean]. </value>
Public Property Get LastCommandOutcome() As Boolean
    LastCommandOutcome = This.LastCommandOutcome
End Property

''' <summary>   Sets the Last Command Outcome. </summary>
''' <param name="a_value">   [Boolean] the last Command Outcome value. </value>
Public Property Let LastCommandOutcome(ByVal a_value As Boolean)
    This.LastCommandOutcome = a_value
End Property

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

        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "SelectedDutNumber"
            Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
        Case "AutoIncrementDutNumberEnabled"
            Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "Measuring"
            Me.Measuring = This.ViewModel.Measuring
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
        Case "Connected"
            Me.Connected = This.ViewModel.Connected
    End Select
    
End Sub

''' <summary>   Synchronizes the view model with the ViewModel. </summary>
''' <remarks>   This is required because the view (ViewModel) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a result, the ViewModel retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' The initial values must be set for property change to work.
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.SelectedDutNumber = This.ViewModel.SelectedDutNumber
    Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
    Me.Measuring = This.ViewModel.Measuring
    Me.Connected = This.ViewModel.Connected
    Me.AutoIncrementDutNumberEnabled = This.ViewModel.AutoIncrementDutNumberEnabled
    Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
    
    ' ensure that the measurement status is set.
    Me.OnMeasurementStateChanged
    
End Sub

''' <summary>   Handles the property change event of the view model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.PropertyChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)

    Const p_procedureName = "m_viewModel_PropertyChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_propertyName As String: p_propertyName = VBA.vbNullString
    If Not a_eventArgs Is Nothing Then
    
        p_propertyName = a_eventArgs.PropertyName
        Me.HandleViewModelPropertyChanged p_propertyName
    
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

''' <summary>   Handles the measurement state change event of the view model. </summary>
''' <remarks>   Disables the user interface controls during measurement state transitions. </remarks>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Core.EventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_MeasurementStateChanging(ByVal a_eventArgs As cc_isr_Core.EventArgs)

    Const p_procedureName = "m_viewModel_MeasurementStateChanging"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not a_eventArgs Is Nothing Then
    
        ' disable controls.
        Me.AutoScanToggleExecutable = False
        Me.AutoSingleToggleExecutable = False
        Me.ManualScanToggleExecutable = False
        Me.ManualSingleToggleExecutable = False
        
    End If
    
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


''' <summary>   Handles the measurement state change event of the view model. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Tcp_Scpi.StateChangedEventArgs"/>] event arguments. </param>
Friend Sub m_viewModel_MeasurementStateChanged(ByVal a_eventArgs As cc_isr_Tcp_Scpi.StateChangedEventArgs)

    Const p_procedureName = "m_viewModel_MeasurementStateChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not a_eventArgs Is Nothing Then
    
        Me.OnMeasurementStateChanged
        
    End If
    
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



