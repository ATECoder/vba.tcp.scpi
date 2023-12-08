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
Private Const m_dateFormat As String = "mm/dd/yyyy"
Private Const m_dateCell As String = "G10"
Private Const m_timeCell As String = "G11"
Private Const m_probesResistanceCell As String = "C10"
Private Const m_compensationFactorCell As String = "C9"

Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 14
Private Const m_dutNumberRow As Integer = 2
Private Const m_dutNumberCol As Integer = 13
Private Const m_byRow As Boolean = True
Private Const m_displayIncrement As Integer = 1

Private Type this_
    FocusedControlName As String
    DesignMode As Boolean
End Type

Private This As this_

''' <summary>   Initializes and returns an instance of this object. </summary>
Public Function Initialize() As UserSheet
    
    Me.AutoScanToggleExecutable = False
    Me.AutoScanToggleValue = False
    Me.AutoSingleToggleExecutable = False
    Me.AutoSingleToggleValue = False
    Me.ManualScanToggleExecutable = False
    Me.ManualScanToggleValue = False
    Me.ManualSingleToggleExecutable = False
    Me.ManualSingleToggleValue = False
    
    ' initialize the state of the command buttons
    Set Initialize = Me

End Function

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

''' <summary>   Gets the sheet range (Cell) that corresponds to the specified DUT number. </summary>
''' <param name="a_dutNumber">   [Integer] the DUT number. </value>
''' <returns>   [<see cref="Excel.Range"/>]. </returns>
Public Function GetDutNumberCaptionRange(ByVal a_dutNumber As Integer) As Excel.Range

    Dim p_row As Integer
    p_row = IIf(m_byRow, m_dutNumberRow + m_displayIncrement * (a_dutNumber - 1), m_dutNumberRow)
    Dim p_col As Integer
    p_col = IIf(m_byRow, m_dutNumberCol, m_dutNumberCol + m_displayIncrement * (a_dutNumber - 1))
    Set GetDutNumberCaptionRange = UserSheet.Cells(p_row, p_col)
    
End Function

''' <summary>   Activates DUT capttion cell of the specified DUT number. </summary>
''' <param name="a_value">   [Integer] the DUT number. </param>
Public Sub ActivateDutNumberCaption(ByVal a_value As Integer)

    Const p_procedureName = "ActivateDutNumberCaption"

    If (0 <> a_value) And Not Me.ProtectContents Then
        If ThisWorkbook.ActiveSheet.Name <> UserSheet.Name Then
            UserSheet.Activate
            VBA.DoEvents
        End If
        ' set this cell as the active cell
        Me.GetDutNumberCaptionRange(a_value).Activate
        ' set the active cell number.
        ' 20231204: Prevent writing the DUT number.
        ' ActiveCell.Value = DataView.DutNumberCaptionPrefix & VBA.CStr(a_value)
    End If

End Sub

''' <summary>   Gets the selected DUT Number caption. </summary>
''' <value>   [String]. </returns>
Public Property Get SelectedDutNumberCaption() As String
    SelectedDutNumberCaption = ActiveCell.Value
End Property

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
    This.FocusedControlName = Me.AutoScanToggleButton.Name
    VBA.DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Auto Scan toggle button.  </summary>
Private Sub AutoScanToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    VBA.DoEvents
End Sub

''' <summary>   Handles the Auto Scan toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub AutoScanToggleButton_Change()

    On Error GoTo err_Handler:
    
    ' this prevents execution when the button is disabled, which is required in order
    ' to correctly initialize the state of the button.
    If This.DesignMode Or Not Me.AutoScanToggleButton.Enabled Then Exit Sub
    
    UserView.AutoScanToggleValue = Me.AutoScanToggleButton.Value
    
    If This.FocusedControlName = Me.AutoScanToggleButton.Name Then

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
    This.FocusedControlName = Me.AutoSingleToggleButton.Name
    VBA.DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Auto Single toggle button.  </summary>
Private Sub AutoSingleToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    VBA.DoEvents
End Sub

''' <summary>   Handles the Auto Single toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub AutoSingleToggleButton_Change()

    On Error GoTo err_Handler:
    
    ' this prevents execution when the button is disabled, which is required in order
    ' to correctly initialize the state of the button.
    If This.DesignMode Or Not Me.AutoSingleToggleButton.Enabled Then Exit Sub
    
    UserView.AutoSingleToggleValue = Me.AutoSingleToggleButton.Value
    
    If This.FocusedControlName = Me.AutoSingleToggleButton.Name Then

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
    If Me.ManualScanToggleButton.Value <> a_value Then _
        Me.ManualScanToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Manual Scan toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub ManualScanToggleButton_GotFocus()
    This.FocusedControlName = Me.ManualScanToggleButton.Name
    VBA.DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Manual Scan toggle button.  </summary>
Private Sub ManualScanToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    VBA.DoEvents
End Sub

''' <summary>   Handles the Manual Scan toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub ManualScanToggleButton_Change()

    On Error GoTo err_Handler:
    
    ' this prevents execution when the button is disabled, which is required in order
    ' to correctly initialize the state of the button.
    If This.DesignMode Or Not Me.ManualScanToggleButton.Enabled Then Exit Sub
    
    UserView.ManualScanToggleValue = Me.ManualScanToggleButton.Value
    
    If This.FocusedControlName = Me.ManualScanToggleButton.Name Then

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
    If Me.ManualSingleToggleButton.Value <> a_value Then _
        Me.ManualSingleToggleButton.Value = a_value
End Property

''' <summary>   Handles the Got Focus event of the Manual Single toggle button.  </summary>
''' <remarks>   Identifies the active toggle button by saving the button name. </remarks>
Private Sub ManualSingleToggleButton_GotFocus()
    This.FocusedControlName = Me.ManualSingleToggleButton.Name
    VBA.DoEvents
End Sub

''' <summary>   Handles the Lost Focus event of the Manual Single toggle button.  </summary>
Private Sub ManualSingleToggleButton_LostFocus()
    This.FocusedControlName = VBA.vbNullString
    VBA.DoEvents
End Sub

''' <summary>   Handles the Manual Single toggle button. </summary>
''' <remarks>   The button change event also causes a click event.
''' This method issues the click event on actual click.
''' Note: Make sure to not implement the click event,
Private Sub ManualSingleToggleButton_Change()

    On Error GoTo err_Handler:
    
    ' this prevents execution when the button is disabled, which is required in order
    ' to correctly initialize the state of the button.
    If This.DesignMode Or Not Me.ManualSingleToggleButton.Enabled Then Exit Sub
    
    UserView.ManualSingleToggleValue = Me.ManualSingleToggleButton.Value
    
    If This.FocusedControlName = Me.ManualSingleToggleButton.Name Then

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


