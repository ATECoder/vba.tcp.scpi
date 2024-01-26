VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DutReadingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Dut reading event argument. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    DutNumber As Integer
    Reading As String
End Type

Private This As this_

''' <summary>   Initializes and returns the event arguments. </summary>
''' <para name="a_dutNumber">   [Integer] The DUT number. </param>
''' <para name="a_reading">     [String] The reading. </param>
''' <value>   [<see cref="DutReadingEventArgs"/>]. </returns>
Public Function Initialize(ByVal a_dutNumber As Integer, ByVal a_reading As String) As DutReadingEventArgs
    This.DutNumber = a_dutNumber
    This.Reading = a_reading
    Set Initialize = Me
End Function

''' <summary>   Gets the DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumber() As Integer
    DutNumber = This.DutNumber
End Property

''' <summary>   Gets the reading. </summary>
''' <value>   [String]. </value>
Public Property Get Reading() As String
    Reading = This.Reading
End Property

