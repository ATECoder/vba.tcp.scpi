VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The K2700 View for linking the K2700 observer with actual user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_readingFormat As String = "C21"
Private Const m_card1 As String = "C23"
Private Const m_card2 As String = "C24"
Private Const m_card1ScanList As String = "C25"
Private Const m_card2ScanList As String = "C26"
Private Const m_FrontInput As String = "C27"
Private Const m_channelNumber As String = "C28"
Private Const m_functionName As String = "B29"
Private Const m_functionValue As String = "C29"
Private Const m_functionUnit As String = "D29"
Private Const m_channelCount As String = "C30"
Private Const m_externalTriggerStatus As String = "C31"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Controls
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private WithEvents ExternalTriggerOptionButton As MSForms.OptionButton
Attribute ExternalTriggerOptionButton.VB_VarHelpID = -1
Private WithEvents FunctionComboBox As MSForms.ComboBox
Attribute FunctionComboBox.VB_VarHelpID = -1
Private WithEvents ImmediateTriggerOptionButton As MSForms.OptionButton
Attribute ImmediateTriggerOptionButton.VB_VarHelpID = -1
Private WithEvents FrontInputCheckBox As MSForms.CheckBox
Attribute FrontInputCheckBox.VB_VarHelpID = -1
Private WithEvents MeasureButton As MSForms.CommandButton
Attribute MeasureButton.VB_VarHelpID = -1
Private WithEvents MonitorExtTrigToggleButton As MSForms.ToggleButton
Attribute MonitorExtTrigToggleButton.VB_VarHelpID = -1
Private WithEvents QueryInputsButton As MSForms.CommandButton
Attribute QueryInputsButton.VB_VarHelpID = -1
Private WithEvents SingleMultipleCheckBox As MSForms.CheckBox
Attribute SingleMultipleCheckBox.VB_VarHelpID = -1

Private Type this_
    SenseFunctions As Collection
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
' K2700 Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the external trigger option button change and click events. </summary>
Private Sub ExternalTriggerOptionButton_Change()
    This.Observer.OnExternalTriggerOptionChange ExternalTriggerOptionButton.Value
End Sub

Public Property Let ExternalTriggerOptionExecutable(ByVal a_value As Boolean)
    If Not ExternalTriggerOptionButton Is Nothing Then
        ExternalTriggerOptionButton.Enabled = a_value
    End If
End Property

Public Property Let ExternalTriggerOptionButtonValue(ByVal a_value As Boolean)
    If Not ExternalTriggerOptionButton Is Nothing Then
        ExternalTriggerOptionButton.Value = a_value
    End If
End Property

''' <summary>   Selects the measurement function. </summary>
Private Sub FunctionComboBox_Change()
    This.Observer.SenseFunctionName = FunctionComboBox.Text
End Sub

Public Property Let FunctionComboBoxEnabled(ByVal a_value As Boolean)
    If Not FunctionComboBox Is Nothing Then
        FunctionComboBox.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Immediate trigger option button change and click events. </summary>
Private Sub ImmediateTriggerOptionButton_Change()
    This.Observer.OnImmediateTriggerOptionChange ImmediateTriggerOptionButton.Value
End Sub

Public Property Let ImmediateTriggerOptionExecutable(ByVal a_value As Boolean)
    If Not ImmediateTriggerOptionButton Is Nothing Then
        ImmediateTriggerOptionButton.Enabled = a_value
    End If
End Property

Public Property Let ImmediateTriggerOptionButtonValue(ByVal a_value As Boolean)
    If Not ImmediateTriggerOptionButton Is Nothing Then
        ImmediateTriggerOptionButton.Value = a_value
    End If
End Property

''' <summary>   Handles the FrontInput check box change and click events. </summary>
Private Sub FrontInputCheckBox_Change()
    This.Observer.FrontInputsRequired = FrontInputCheckBox.Value
End Sub

Public Property Let FrontInputCheckBoxEnabled(ByVal a_value As Boolean)
    If Not FrontInputCheckBox Is Nothing Then
        FrontInputCheckBox.Enabled = a_value
    End If
End Property

Public Property Let FrontInputEnabled(ByVal a_value As Boolean)
    If Not FrontInputCheckBox Is Nothing Then
        FrontInputCheckBox.Value = a_value
    End If
End Property

Private Sub MeasureButton_Click()
    This.Observer.OnMeasure
End Sub

Public Property Let MeasureExecutable(ByVal a_value As Boolean)
    If Not MeasureButton Is Nothing Then
        MeasureButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the external trigger toggle button. </summary>
Private Sub MonitorExtTrigToggleButton_Change()
    This.Observer.OnMonitorExtTrigToggleButtonChange MonitorExtTrigToggleButton.Value
End Sub

Public Property Let MonitorExtTrigExecutable(ByVal a_value As Boolean)
    If Not MonitorExtTrigToggleButton Is Nothing Then
        MonitorExtTrigToggleButton.Enabled = a_value
    End If
End Property

Public Property Let MonitorExtTrigEnabled(ByVal a_value As Boolean)
    If Not MonitorExtTrigToggleButton Is Nothing Then
        MonitorExtTrigToggleButton.Value = a_value
    End If
End Property

Private Sub QueryInputsButton_Click()
    This.Observer.OnQueryInputs
End Sub

Public Property Let QueryInputsExecutable(ByVal a_value As Boolean)
    If Not QueryInputsButton Is Nothing Then
        QueryInputsButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the SingleMultiple check box change and click events. </summary>
Private Sub SingleMultipleCheckBox_Change()
    This.Observer.SingleReadEnabled = SingleMultipleCheckBox.Value
End Sub

Public Property Let SingleMultipleCheckBoxEnabled(ByVal a_value As Boolean)
    If Not SingleMultipleCheckBox Is Nothing Then
        SingleMultipleCheckBox.Enabled = a_value
    End If
End Property

Public Property Let SingleReadEnabled(ByVal a_value As Boolean)
    If Not SingleMultipleCheckBox Is Nothing Then
        SingleMultipleCheckBox.Value = a_value
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
        If Not a_value Is Nothing And Not FunctionComboBox Is Nothing Then
            Dim p_fourWireResistanceIndex As Integer
            FunctionComboBox.Clear
            Dim p_item As Variant
            Dim p_senseFunction As String
            For Each p_item In a_value
                p_senseFunction = VBA.CStr(p_item)
                FunctionComboBox.AddItem p_senseFunction
                If "FRES" = p_senseFunction Then
                    p_fourWireResistanceIndex = FunctionComboBox.Count
                End If
            Next
            FunctionComboBox.ListIndex = p_fourWireResistanceIndex
        End If
    End If
End Property

