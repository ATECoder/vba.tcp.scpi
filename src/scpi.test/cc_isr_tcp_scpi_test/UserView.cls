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
    UserSampleSheet As UserSheet
    
    AutoIncrementChannelNoEnabled As Boolean
    FrontInputsRequired As Boolean
    SingleReadEnabled As Boolean
    
    ReadingOffset As Double
    
    SelectedChannelNumber As Integer
    SelectedChannelNumberCaption As String
    MeasurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption
    
    ChannelNumberCaptionPrefix As String
    ChannelNumberTitle As String
        
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
    This.ChannelNumberCaptionPrefix = "R"
    This.ChannelNumberTitle = "resistance number"
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
    
    Set This.UserSampleSheet = Nothing
    
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
Public Function Initialize(ByVal a_viewModel As K2700ViewModel) As UserView

    Const p_procedureName = "Initialize"

    ' Error are trapped to prevent crashes because this method gets started when
    ' the main workbook is activated.
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Set This.UserSampleSheet = UserSheet.Initialize()
    
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
' Data Acquisition Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the Auto Scan toggle button. </summary>
''' <remarks>   In auto scan we have the following settings:
''' Auto Increment channel number: True;
''' Front Inputs Required: False;
''' Single Readings: False;
''' Triggering Mode: Immediate.
''' </remarks>
Public Sub OnAutoScanToggleButtonChange()

    Const p_procedureName As String = "OnAutoScanToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.AutoScanToggleValue Then
    
        This.UserSampleSheet.DisplayStartDateTime
        Me.AutoIncrementChannelNoEnabled = True
        Me.FrontInputsRequired = This.AutoSampleFrontInputs
        Me.SingleReadEnabled = False
        This.ViewModel.ReadingBeepEnabled = False
        
        ' next set immediate triggering mode.
        
        If This.ViewModel.ImmediateTriggerOptionExecutable Then
            Me.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(Me.LastCommandDetails)
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Immediate trigger command not enabled."
        End If
        
        If Me.LastCommandOutcome Then
        
            ' next takes the measurements
        
            If This.ViewModel.MeasureExecutable Then
                ' make an immediate measure of one or all channels.
                Me.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(Me.LastCommandDetails)
            Else
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " Measure command not enabled."
            End If
        
        End If
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' toggle the button state
    Me.AutoScanToggleValue = False
    
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

''' <summary>   Gets the execution (Can Execute) status of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get AutoScanToggleExecutable() As Boolean
    
    If This.UserSampleSheet Is Nothing Then
        AutoScanToggleExecutable = This.AutoScanToggleExecutable
    Else
        AutoScanToggleExecutable = This.UserSampleSheet.AutoScanToggleExecutable
    End If

End Property

''' <summary>   Sets the execution (Can Execute) status of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoScanToggleExecutable(ByVal a_value As Boolean)
    If Me.AutoScanToggleExecutable <> a_value Then
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.AutoScanToggleExecutable = a_value
        
        This.AutoScanToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoScanToggleValue() As Boolean
    
    If This.UserSampleSheet Is Nothing Then
        AutoScanToggleValue = This.AutoScanToggleValue
    Else
        AutoScanToggleValue = This.UserSampleSheet.AutoScanToggleValue
    End If

End Property

''' <summary>   Sets the action state of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoScanToggleValue(ByVal a_value As Boolean)
    
    If Me.AutoScanToggleValue <> a_value Then
    
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.AutoScanToggleValue = a_value

        This.AutoScanToggleValue = a_value
    End If

End Property

''' <summary>   Handles the Auto Single toggle button. </summary>
''' <remarks>   In auto single we have the following settings:
''' Auto Increment channel number: False;
''' Front Inputs Required: False;
''' Single Readings: True;
''' Triggering Mode: Immediate.
''' </remarks>
Public Sub OnAutoSingleToggleButtonChange()

    Const p_procedureName As String = "OnAutoSingleToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If This.ViewModel Is Nothing Then Exit Sub

    If Me.AutoSingleToggleValue Then
    
        Me.SelectedChannelNumber = This.UserSampleSheet.SelectedChannelNumber
        This.UserSampleSheet.DisplayStartDateTime
    
        Me.AutoIncrementChannelNoEnabled = False
        Me.FrontInputsRequired = This.AutoSampleFrontInputs
        Me.SingleReadEnabled = True
        This.ViewModel.ReadingBeepEnabled = False
        
        ' next set immediate triggering mode.
        
        If This.ViewModel.ImmediateTriggerOptionExecutable Then
            Me.LastCommandOutcome = This.ViewModel.ConfigureImmediateTriggerReadingsCommand(Me.LastCommandDetails)
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Immediate trigger command not enabled."
        End If
        
        If Me.LastCommandOutcome Then
        
            ' next takes the measurements
        
            If This.ViewModel.MeasureExecutable Then
                ' make an immediate measure of one or all channels.
                Me.LastCommandOutcome = This.ViewModel.MeasureImmediatelyCommand(Me.LastCommandDetails)
            Else
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " Measure command not enabled."
            End If
        
        End If

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' toggle the button state
    Me.AutoSingleToggleValue = False
    
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

''' <summary>   Gets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get AutoSingleToggleExecutable() As Boolean
    
    If This.UserSampleSheet Is Nothing Then
        AutoSingleToggleExecutable = This.AutoSingleToggleExecutable
    Else
        AutoSingleToggleExecutable = This.UserSampleSheet.AutoSingleToggleExecutable
    End If

End Property

''' <summary>   Sets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoSingleToggleExecutable(ByVal a_value As Boolean)
    
    If Me.AutoSingleToggleExecutable <> a_value Then
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.AutoSingleToggleExecutable = a_value

        This.AutoSingleToggleExecutable = a_value
    End If

End Property

''' <summary>   Gets the action state of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoSingleToggleValue() As Boolean
    
    If This.UserSampleSheet Is Nothing Then
        AutoSingleToggleValue = This.AutoSingleToggleValue
    Else
        AutoSingleToggleValue = This.UserSampleSheet.AutoSingleToggleValue
    End If

End Property

''' <summary>   Sets the action state of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoSingleToggleValue(ByVal a_value As Boolean)
    
    If Me.AutoSingleToggleValue <> a_value Then
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.AutoSingleToggleValue = a_value

        This.AutoSingleToggleValue = a_value
    End If

End Property

''' <summary>   Handles the Manual Scan toggle button. </summary>
''' <remarks>   In manual scan we have the following settings:
''' Auto Increment channel number: True;
''' Front Inputs Required: True;
''' Single Readings: False;
''' Triggering Mode: External.
''' </remarks>
Public Sub OnManualScanToggleButtonChange()

    Const p_procedureName As String = "OnManualScanToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.ManualScanToggleValue Then
    
        This.UserSampleSheet.DisplayStartDateTime
    
        Me.AutoIncrementChannelNoEnabled = True
        Me.FrontInputsRequired = True
        Me.SingleReadEnabled = False
        This.ViewModel.ReadingBeepEnabled = DataView.BeepAllowed
        
        ' next set external triggering triggering mode.
        
        If This.ViewModel.ExternalTriggerOptionExecutable Then
            Me.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(Me.LastCommandDetails)
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " External trigger enable command not enabled."
        End If
        
        If Me.LastCommandOutcome Then
        
            ' next start external trigger monitoring
            ' this should wait for monitoring to commence or raise an error.
            If This.ViewModel.StartMonitoringExecutable Then
                This.ViewModel.StartMonitoringExternalTriggersCommand
            Else
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " Start monitoring command not enabled."
            End If
        
        End If
        
    Else
    
        ' stop external trigger monitoring
        ' this should wait for monitoring to stop
        
        If This.ViewModel.StopMonitoringExecutable Then
            This.ViewModel.StopMonitoringExternalTriggersCommand
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Stop monitoring command not enabled."
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
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
    If This.UserSampleSheet Is Nothing Then
        ManualScanToggleExecutable = This.ManualScanToggleExecutable
    Else
        ManualScanToggleExecutable = This.UserSampleSheet.ManualScanToggleExecutable
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualScanToggleExecutable(ByVal a_value As Boolean)
    If Me.ManualScanToggleExecutable <> a_value Then
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.ManualScanToggleExecutable = a_value

        This.ManualScanToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Manuam Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualScanToggleValue() As Boolean
    If This.UserSampleSheet Is Nothing Then
        ManualScanToggleValue = This.ManualScanToggleValue
    Else
        ManualScanToggleValue = This.UserSampleSheet.ManualScanToggleValue
    End If
End Property

''' <summary>   Sets the action state of the Manuam Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualScanToggleValue(ByVal a_value As Boolean)
    If Me.ManualScanToggleValue <> a_value Then
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.ManualScanToggleValue = a_value

        This.ManualScanToggleValue = a_value
    End If
End Property

''' <summary>   Handles the Manual Single toggle button. </summary>
''' <remarks>   In manual Single we have the following settings:
''' Auto Increment channel number: False;
''' Front Inputs Required = True;
''' Single Readings: True;
''' Triggering Mode: External.
''' </remarks>
Public Sub OnManualSingleToggleButtonChange()

    Const p_procedureName As String = "OnManualSingleToggleButtonChange"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.ViewModel Is Nothing Then Exit Sub

    If Me.ManualSingleToggleValue Then
    
        Me.SelectedChannelNumber = This.UserSampleSheet.SelectedChannelNumber
        This.UserSampleSheet.DisplayStartDateTime
        
        Me.AutoIncrementChannelNoEnabled = False
        Me.FrontInputsRequired = True
        Me.SingleReadEnabled = True
        This.ViewModel.ReadingBeepEnabled = DataView.BeepAllowed
        
        ' next set external triggering triggering mode.
        
        If This.ViewModel.ExternalTriggerOptionExecutable Then
            Me.LastCommandOutcome = This.ViewModel.ConfigureExternalTriggerReadingCommand(Me.LastCommandDetails)
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " External trigger enable command not enabled."
        End If
        
        If Me.LastCommandOutcome Then
        
            ' next start external trigger monitoring
            ' this should wait for monitoring to commence or raise an error.
            If This.ViewModel.StartMonitoringExecutable Then
                This.ViewModel.StartMonitoringExternalTriggersCommand
            Else
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " Start monitoring command not enabled."
            End If
        
        End If
        
    Else
    
        ' stop external trigger monitoring
        ' this should wait for monitoring to stop
        
        If This.ViewModel.StopMonitoringExecutable Then
            This.ViewModel.StopMonitoringExternalTriggersCommand
        Else
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Stop monitoring command not enabled."
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
        This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
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
    If This.UserSampleSheet Is Nothing Then
        ManualSingleToggleExecutable = This.ManualSingleToggleExecutable
    Else
        ManualSingleToggleExecutable = This.UserSampleSheet.ManualSingleToggleExecutable
    End If
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualSingleToggleExecutable(ByVal a_value As Boolean)
    If Me.ManualSingleToggleExecutable <> a_value Then
    
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.ManualSingleToggleExecutable = a_value

        This.ManualSingleToggleExecutable = a_value
    End If
End Property

''' <summary>   Gets the action state of the Manuam Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualSingleToggleValue() As Boolean
    If This.UserSampleSheet Is Nothing Then
        ManualSingleToggleValue = This.ManualSingleToggleValue
    Else
        ManualSingleToggleValue = This.UserSampleSheet.ManualSingleToggleValue
    End If
End Property

''' <summary>   Sets the action state of the Manuam Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualSingleToggleValue(ByVal a_value As Boolean)
    If Me.ManualSingleToggleValue <> a_value Then
    
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.ManualSingleToggleValue = a_value

        This.ManualSingleToggleValue = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.AutoIncrementChannelNoEnabled = a_value
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

''' <summary>   Gets the condition for using front inputs for auto sampe. </summary>
''' <remarks>   True if front inputs are Required; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AutoSampleFrontInputs() As Boolean
    AutoSampleFrontInputs = This.AutoSampleFrontInputs
End Property

''' <summary>   Sets the condition for using front inputs for auto sampe. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs are used for auto sampling. </value>
Public Property Let AutoSampleFrontInputs(ByVal a_value As Boolean)
    If This.AutoSampleFrontInputs <> a_value Then
        This.AutoSampleFrontInputs = a_value
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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.FrontInputsRequired = a_value
    End If
End Property

''' <summary>   Gets the Selected Channel Number Caption. </summary>
''' <value>   [String]. </value>
Public Property Get SelectedChannelNumberCaption() As String
    SelectedChannelNumberCaption = This.SelectedChannelNumberCaption
End Property

Public Property Let SelectedChannelNumberCaption(ByVal a_value As String)
    If Not This.ViewModel Is Nothing Then
        If This.SelectedChannelNumberCaption <> a_value Then
            This.SelectedChannelNumberCaption = a_value
            Dim p_details As String: p_details = VBA.vbNullString
            Dim p_channelNumber As Integer
            p_channelNumber = Me.TryParseSelectedChannelNumber(a_value, p_details)
            If 0 < p_channelNumber And This.ViewModel.ChannelCount >= p_channelNumber Then
                Me.SelectedChannelNumber = p_channelNumber
            Else
                MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
                    "Invalid " & This.ChannelNumberTitle
            End If
        End If
    End If
End Property

''' <summary>   Returns true if the value is a valid channel number. </summary>
''' <param name="a_value">   [Integer] the channel number. </value>
''' <returns>   [Boolean]. </returns>
Public Function IsValidChannelNumber(ByVal a_value As Integer) As Boolean
    IsValidChannelNumber = (0 < a_value) And (This.ViewModel.ChannelCount >= a_value)
End Function

''' <summary>   Gets a channel number from the channel number caption. </summary>
''' <param name="a_value">   [String] the selected channel number caption, e.g., R2. </value>
''' <returns>   [Integer] Returns 0 if invalid channel number caption was selected. </returns>
Public Function TryParseSelectedChannelNumber(ByVal a_value As String, ByRef a_details As String) As Integer

    Const p_procedureName = "TryParseSelectedChannelNumber"
    
    Dim p_channelNo As Integer: p_channelNo = 0
    
    If VBA.Left$(a_value, 1) <> This.ChannelNumberCaptionPrefix Then
    
        a_details = "The selected value '" & a_value & _
               "' is not a valid " & This.ChannelNumberTitle & _
               ". Please select a valid " & This.ChannelNumberTitle & _
               ", e.g., " & This.ChannelNumberCaptionPrefix & "2, and try again."
  
    Else
    
        p_channelNo = VBA.CInt(Mid$(a_value, 2))
        
        If Not Me.IsValidChannelNumber(p_channelNo) Then
        
            a_details = "The selected value '" & a_value & _
                   "' yields an out of range " & This.ChannelNumberTitle & _
                   "'. Valid " & This.ChannelNumberTitle & "s are between 1 and " & _
                   VBA.CStr(This.ViewModel.ChannelCount) & _
                   ". Please select a valid " & This.ChannelNumberTitle & _
                   ", e.g., " & This.ChannelNumberCaptionPrefix & "2, and try again."
            
            p_channelNo = 0
            
        End If
    
    End If
  
    TryParseSelectedChannelNumber = p_channelNo
  
End Function

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
        
        ' emulate selecting a channel number if within range.
        If Me.IsValidChannelNumber(a_value) Then _
            Me.SelectedChannelNumberCaption = This.ChannelNumberCaptionPrefix & VBA.CStr(a_value)
        
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SelectedChannelNumber = a_value
        
        If Not This.UserSampleSheet Is Nothing Then _
            This.UserSampleSheet.SelectedChannelNumber = a_value

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
        If Not This.ViewModel Is Nothing Then _
            This.ViewModel.SingleReadEnabled = a_value
        Me.OnMeasurementModeChanged Me.MeasurementMode, a_value
    End If
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
    If This.MeasurementMode <> a_value Then
        This.MeasurementMode = a_value
        Me.OnMeasurementModeChanged a_value, Me.SingleReadEnabled
    End If
End Property

''' <summary>   Handles the change of measurment mode or single read enabled. </summary>
''' <remarks>   This enableds or disables the buttons and toggles the buttons
''' state if not toggled. </remarks>
''' <param name=-"a_measurementMode>     [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <param name=-"a_singleReadEnabled>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
Friend Sub OnMeasurementModeChanged(ByVal a_measurementMode As cc_isr_Tcp_Scpi.MeasurementModeOption, _
    ByVal a_singleReadEnabled As Boolean)

    If cc_isr_Tcp_Scpi.MeasurementModeOption.None = a_measurementMode Then
        Me.ManualScanToggleExecutable = True
        Me.ManualSingleToggleExecutable = True
        Me.AutoScanToggleExecutable = True
        Me.AutoSingleToggleExecutable = True
        If AutoScanToggleValue Then _
            Me.AutoScanToggleValue = False
        If AutoSingleToggleValue Then _
            Me.AutoSingleToggleValue = False
        If ManualScanToggleValue Then _
            Me.ManualScanToggleValue = False
        If ManualSingleToggleValue Then _
            Me.ManualSingleToggleValue = False
    ElseIf cc_isr_Tcp_Scpi.MeasurementModeOption.Immediate = a_measurementMode Then
        Me.ManualScanToggleExecutable = False
        Me.ManualSingleToggleExecutable = False
        Me.AutoScanToggleExecutable = Not a_singleReadEnabled
        Me.AutoSingleToggleExecutable = a_singleReadEnabled
        If ManualScanToggleValue Then _
            Me.ManualScanToggleValue = False
        If ManualSingleToggleValue Then _
            Me.ManualSingleToggleValue = False
    ElseIf cc_isr_Tcp_Scpi.MeasurementModeOption.External = a_measurementMode Or _
           cc_isr_Tcp_Scpi.MeasurementModeOption.Monitoring = a_measurementMode Then
        Me.ManualScanToggleExecutable = Not a_singleReadEnabled
        Me.ManualSingleToggleExecutable = a_singleReadEnabled
        Me.AutoScanToggleExecutable = False
        Me.AutoSingleToggleExecutable = False
        If Me.AutoScanToggleValue Then _
            Me.AutoScanToggleValue = False
        If Me.AutoSingleToggleValue Then _
            Me.AutoSingleToggleValue = False
    End If

End Sub

''' <summary>   Gets the Last Command Details. </summary>
''' <value>   [String]. </value>
Public Property Get LastCommandDetails() As String
    LastCommandDetails = This.LastCommandDetails
End Property

''' <summary>   Sets the Last Command Details. </summary>
''' <param name="a_value">   [String] the last Command Details value. </value>
Public Property Let LastCommandDetails(ByVal a_value As String)
    If This.LastCommandDetails <> a_value Then
        This.LastCommandDetails = a_value
    End If
End Property

''' <summary>   Gets the Last Command Outcome. </summary>
''' <value>   [Boolean]. </value>
Public Property Get LastCommandOutcome() As Boolean
    LastCommandOutcome = This.LastCommandOutcome
End Property

''' <summary>   Sets the Last Command Outcome. </summary>
''' <param name="a_value">   [Boolean] the last Command Outcome value. </value>
Public Property Let LastCommandOutcome(ByVal a_value As Boolean)
    If This.LastCommandOutcome <> a_value Then
        This.LastCommandOutcome = a_value
    End If
End Property

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

        Case "FrontInputsRequired"
            Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
        Case "SelectedChannelNumber"
            Me.SelectedChannelNumber = This.ViewModel.SelectedChannelNumber
        Case "AutoIncrementChannelNoEnabled"
            Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
        Case "SingleReadEnabled"
            Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
        Case "MeasurementMode"
            Me.MeasurementMode = This.ViewModel.MeasurementMode
        Case "ReadingOffset"
            Me.ReadingOffset = This.ViewModel.ReadingOffset
    End Select
    
End Sub

''' <summary>   Synchronizes the view mode with the ViewModel. </summary>
''' <remarks>   This is required becuase the view (ViewModel) that is planned for this
''' view model is a Sheet, which is a singleton class whereas the view model is a
''' non-singleton class. As a resule, the ViewModel retains values from previous
''' instances of the View Model. </remarks>
Friend Sub Synchronize()

    Const p_procedureName = "Synchronize"

    ' set the view model settings that come from the ViewModel user interface
    Me.ReadingOffset = This.UserSampleSheet.InitialResistance
    
    This.ViewModel.ReadingOffset = This.ReadingOffset

    Me.AutoIncrementChannelNoEnabled = This.ViewModel.AutoIncrementChannelNoEnabled
    Me.FrontInputsRequired = This.ViewModel.FrontInputsRequired
    Me.MeasurementMode = This.ViewModel.MeasurementMode
    Me.ReadingOffset = This.ViewModel.ReadingOffset
    Me.SelectedChannelNumber = This.ViewModel.SelectedChannelNumber
    Me.SingleReadEnabled = This.ViewModel.SingleReadEnabled
    
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

