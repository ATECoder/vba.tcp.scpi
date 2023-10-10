VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The user sheet. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_timeFormat As String = "h:mm"
Private Const m_dateCell As String = "G10"
Private Const m_timeCell As String = "G11"
Private Const m_probesResistanceCell As String = "C10"
Private Const m_compensationFactorCell As String = "C9"

Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 23
Private Const m_channelRow As Integer = 2
Private Const m_channelCol As Integer = 24
Private Const m_byRow As Boolean = True
Private Const m_displayIncrement As Integer = 1

Private Type this_
    FocusedControlName As String
End Type

Private This As this_

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize() As UserSheet
    
    Set Initialize = Me

End Function

''' <summary>   Set the sheet date and time display. </summary>
Public Sub DisplayStartDateTime()

    ' Set the data and time
    Range(m_dateCell).Value = Now
    Range(m_timeCell).Value = Format$(Now, m_timeFormat)

End Sub

''' <summary>   Gets the resistance of the probes. </summary>
''' <value>   [Double]. </value>
Public Property Get ProbesResistance() As Double
    
    ProbesResistance = VBA.CDbl(Range(m_probesResistanceCell).Value)

End Property

''' <summary>   Gets the compensation factor. </summary>
''' <value>   [Double]. </value>
Public Property Get CompensationFactor() As Double
    
    CompensationFactor = VBA.CDbl(Range(m_compensationFactorCell).Value)

End Property

''' <summary>   Gets the initial resistance for clearing the acqisition dataprobes. </summary>
''' <value>   [Double] Equals the <see cref="CompensationFactor"/> times <see cref="ProbesResistance"/>. </value>
Public Property Get InitialResistance() As Double
    
    InitialResistance = Me.CompensationFactor * Me.ProbesResistance

End Property

''' <summary>   Gets the selected channge number range. </summary>
''' <param name="a_channelNo">   [Integer] the selected channel number. </value>
''' <value>   [<see cref="Excel.Range"/>]. </value>
Public Function GetSelectedChannelNumberRange(ByVal a_channelNo As Integer) As Excel.Range

    Dim p_row As Integer
    p_row = IIf(m_byRow, m_channelRow + m_displayIncrement * (a_channelNo - 1), m_channelRow)
    Dim p_col As Integer
    p_col = IIf(m_byRow, m_channelCol, m_channelCol + m_displayIncrement * (a_channelNo - 1))
    Set GetSelectedChannelNumberRange = Cells(p_row, p_col)
    
End Function

''' <summary>   Gets the selected channel number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedChannelNumber() As Integer
    SelectedChannelNumber = Me.GetActiveCellResistanceNo
End Property

''' <summary>   Sets the selected channel number. </summary>
''' <param name="a_value">   [Integer] the selected channel number. </param>
Public Property Let SelectedChannelNumber(ByVal a_value As Integer)
    
    If Not Me.ProtectContents Then
        ActiveCell = Me.GetSelectedChannelNumberRange(a_value)
        ActiveCell.Value = "R" & VBA.CStr(a_value)
    End If

End Property

''' <summary>   Parses the resistance channel number value from the value of the active cell. </summary>
''' <returns>   [Integer] Returns 0 if invalid resistance cell was selected. </returns>
Public Function GetActiveCellResistanceNo() As Integer

    Const p_procedureName = "GetActiveCellResistanceNo"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_channelNumber As Integer
    p_channelNumber = UserView.TryParseSelectedChannelNumber(VBA.UCase$(ActiveCell.Value), p_details)
    If 0 = p_channelNumber Then
        MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
            "Invalid Resistance Number"
    End If
    
    GetActiveCellResistanceNo = p_channelNumber
  
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sets the execution Can Execute) status of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoScanToggleExecutable() As Boolean
    AutoScanToggleExecutable = Me.AutoScanToggleButton.Enabled
End Property

''' <summary>   Sets the execution Can Execute) status of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoScanToggleExecutable(ByVal a_value As Boolean)
    Me.AutoScanToggleButton.Enabled = a_value
End Property

''' <summary>   Gets the action state of the Auto Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoScanToggleValue() As Boolean
    AutoScanToggleValue = Me.AutoScanToggleButton.Value
End Property

''' <summary>   Sets the action state of the Auto Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoScanToggleValue(ByVal a_value As Boolean)
    Me.AutoScanToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Auto Scan toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub AutoScanToggleButton_GotFocus()
    This.FocusedControlName = AutoScanToggleButton.Name
    DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Auto Scan toggle button.  </summary>
Private Sub AutoScanToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the Auto Scan toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub AutoScanToggleButton_Change()

    On Error GoTo err_Handler:
    
    UserView.AutoScanToggleValue = AutoScanToggleButton.Value
    
    If This.FocusedControlName = AutoScanToggleButton.Name Then

        On Error GoTo 0

        ' Commands for click
        UserView.OnAutoScanToggleButtonChange
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        UserView.OnAutoScanToggleButtonChange

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    UserView.OnAutoScanToggleButtonChange
    
End Sub

''' <summary>   Gets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is executable. </value>
Public Property Get AutoSingleToggleExecutable() As Boolean
    AutoSingleToggleExecutable = Me.AutoSingleToggleButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let AutoSingleToggleExecutable(ByVal a_value As Boolean)
    Me.AutoSingleToggleButton.Enabled = a_value
End Property

''' <summary>   Gets the action state of the Auto Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get AutoSingleToggleValue() As Boolean
    AutoSingleToggleValue = Me.AutoSingleToggleButton.Value
End Property

''' <summary>   Sets the action state of the Auto Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let AutoSingleToggleValue(ByVal a_value As Boolean)
    Me.AutoSingleToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Auto Single toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub AutoSingleToggleButton_GotFocus()
    This.FocusedControlName = AutoSingleToggleButton.Name
    DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Auto Single toggle button.  </summary>
Private Sub AutoSingleToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the Auto Single toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub AutoSingleToggleButton_Change()

    On Error GoTo err_Handler:
    
    UserView.AutoSingleToggleValue = AutoSingleToggleButton.Value
    
    If This.FocusedControlName = AutoSingleToggleButton.Name Then

        On Error GoTo 0
        
        ' Commands for click
        UserView.OnAutoSingleToggleButtonChange
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        UserView.OnAutoSingleToggleButtonChange

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    UserView.OnAutoSingleToggleButtonChange
    
End Sub

''' <summary>   Gets the execution (Can Execute) status of the Manual Scan Toggle Button. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ManualScanToggleExecutable() As Boolean
    ManualScanToggleExecutable = Me.ManualScanToggleButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Scan Toggle Button. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualScanToggleExecutable(ByVal a_value As Boolean)
    Me.ManualScanToggleButton.Enabled = a_value
End Property

''' <summary>   Gets the action state of the Manual Scan Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualScanToggleValue() As Boolean
    ManualScanToggleValue = Me.ManualScanToggleButton.Value
End Property

''' <summary>   Sets the action state of the Manual Scan Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualScanToggleValue(ByVal a_value As Boolean)
    Me.ManualScanToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Manual Scan toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub ManualScanToggleButton_GotFocus()
    This.FocusedControlName = ManualScanToggleButton.Name
    DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Manual Scan toggle button.  </summary>
Private Sub ManualScanToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the Manual Scan toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub ManualScanToggleButton_Change()

    On Error GoTo err_Handler:
    
    UserView.ManualScanToggleValue = ManualScanToggleButton.Value
    
    If This.FocusedControlName = ManualScanToggleButton.Name Then

        On Error GoTo 0

        ' Commands for click
        UserView.OnManualScanToggleButtonChange
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        UserView.OnManualScanToggleButtonChange

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    UserView.OnManualScanToggleButtonChange
    
End Sub

''' <summary>   Gets the execution (Can Execute) status of the Manual Single Toggle command. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ManualSingleToggleExecutable() As Boolean
    ManualSingleToggleExecutable = Me.ManualSingleToggleButton.Enabled
End Property

''' <summary>   Sets the execution (Can Execute) status of the Manual Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is executable. </param>
Public Property Let ManualSingleToggleExecutable(ByVal a_value As Boolean)
    Me.ManualSingleToggleButton.Enabled = a_value
End Property

''' <summary>   Gets the action state of the Manual Single Toggle command. </summary>
''' <value>   [Boolean] true if the command is active. </value>
Public Property Get ManualSingleToggleValue() As Boolean
    ManualSingleToggleValue = Me.ManualSingleToggleButton.Value
End Property

''' <summary>   Sets the action state of the Manual Single Toggle command. </summary>
''' <param name="a_value">   [Boolean] true if the command is active. </param>
Public Property Let ManualSingleToggleValue(ByVal a_value As Boolean)
    Me.ManualSingleToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Manual Single toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub ManualSingleToggleButton_GotFocus()
    This.FocusedControlName = ManualSingleToggleButton.Name
    DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Manual Single toggle button.  </summary>
Private Sub ManualSingleToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    DoEvents
End Sub

''' <summary>   Handles the Manual Single toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub ManualSingleToggleButton_Change()

    On Error GoTo err_Handler:
    
    UserView.ManualSingleToggleValue = ManualSingleToggleButton.Value
    
    If This.FocusedControlName = ManualSingleToggleButton.Name Then

        On Error GoTo 0

        ' Commands for click
        UserView.OnManualSingleToggleButtonChange
        
    Else

        On Error GoTo 0

        ' Commands for change from within the sheet
        UserView.OnManualSingleToggleButtonChange

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    On Error GoTo 0
    
    ' Commands for change from outside the sheet
    UserView.OnManualSingleToggleButtonChange
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Address = m_probesResistanceCell Then
        UserView.ReadingOffset = Me.InitialResistance
    ElseIf Target.Address = m_compensationFactorCell Then
        UserView.ReadingOffset = Me.InitialResistance
    End If

End Sub


