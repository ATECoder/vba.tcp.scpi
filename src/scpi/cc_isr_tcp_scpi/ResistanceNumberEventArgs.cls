VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResistanceNumberEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type ThisData
    ResistanceNumber As Integer
    Resistance As Double
End Type

Private This As ThisData

''' <summary>   Initializes and returns the event arguments. </summary>
''' <para name="a_resistanceNo">   [integer] The resistance number. </param>
''' <para name="a_resistance">     [double] The resistance. </param>
''' <value>   [<see cref="ResistanceNumberEventArgs"/>]. </returns>
Public Function Initialize(ByVal a_resistanceNo As Integer, ByVal a_resistance As Double) As ResistanceNumberEventArgs
    This.ResistanceNumber = a_resistanceNo
    This.Resistance = a_resistance
    Set Initialize = Me
End Function

''' <summary>   Gets the resistance number. </summary>
''' <value>   A integer. </value>
Public Property Get ResistanceNumber() As Integer
    ResistanceNumber = This.ResistanceNumber
End Property

''' <summary>   Gets the resistance. </summary>
''' <value>   A double. </value>
Public Property Get Resistance() As Double
    Resistance = This.Resistance
End Property

