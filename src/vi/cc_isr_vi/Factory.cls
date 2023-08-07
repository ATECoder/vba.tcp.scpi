VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private m_k2700 As K2700

''' <summary>   Instantiates the <see cref="K2700"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function NewK2700() As K2700
    Set NewK2700 = New K2700
End Function

''' <summary>   Gets an instance of the <see cref="ResistanceNumberEventArgs"/> class. </summary>
''' <value>   [<see cref="ResistanceNumberEventArgs"/>]. </returns>
Public Property Get NewResistanceNumberEventArgs() As ResistanceNumberEventArgs
    Set NewResistanceNumberEventArgs = New ResistanceNumberEventArgs
End Property

''' <summary>   Instantiates the <see cref="ScpiSystem"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="ScpiSystem"/>]. </returns>
Public Function NewScpiSystem() As ScpiSystem
    Set NewScpiSystem = New ScpiSystem
End Function

