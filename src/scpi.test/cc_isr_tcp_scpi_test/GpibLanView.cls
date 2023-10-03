VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GpibLanView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The GPIB-Lan View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

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



