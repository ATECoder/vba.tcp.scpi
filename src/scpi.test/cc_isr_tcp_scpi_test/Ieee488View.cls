VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The IEEE488 View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

' Controls

Private WithEvents CLSButton As MSForms.CommandButton
Attribute CLSButton.VB_VarHelpID = -1

Private Type this_
    Observer As K2700Observer
End Type

Private This As this_

Private Sub Class_Initialize()
    Set This.Observer = K2700Observer
End Sub

Private Sub Class_Terminate()
    Set This.Observer = Nothing
End Sub


' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' IEEE488 Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
    This.Observer.OnClearExecutionState
End Sub

Public Property Let CLSButtonEnabled(ByVal a_value As Boolean)
    If Not CLSButton Is Nothing Then
        CLSButton.Enabled = a_value
    End If
End Property


