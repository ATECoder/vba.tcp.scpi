VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiplexerCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Multiplexer card. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type ThisData
    Name As String
    Capacity As Integer
    slotNumber As Integer
    ScanList As String
End Type

Private This As ThisData

''' <summary>   Initializes a new card. </summary>
''' <param name="a_name">       The card name. </param>
''' <param name="a_capacity">   The card capacity in number of channels. </param>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Sub Initialize(ByVal a_name As String, ByVal a_capacity As Integer)
    This.Capacity = a_capacity
    This.Name = a_name
    This.slotNumber = 0
End Sub

''' <summary>   IMakes a deep copy of this object. </summary>
''' <param name="a_card">   The card to clone </param>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Function Clone(ByVal a_card As multiplexerCard) As multiplexerCard
    
    Dim p_clone As New multiplexerCard
    Clone = p_clone.Initialize(a_card.Name, a_card.Capacity)
    
End Function

''' <summary>   Gets the multiplexer card name. </summary>
''' <value> [String]. </value>
Public Property Get Name() As String
        
    Name = This.Name

End Property

''' <summary>   Gets the multiplexer card Capacity. </summary>
''' <value> [Integer]. </value>
Public Property Get Capacity() As Integer
        
    Capacity = This.Capacity

End Property

''' <summary>   Gets the multiplexer card Slot Number. </summary>
''' <value> [Integer]. </value>
Public Property Get slotNumber() As Integer
        
    slotNumber = This.slotNumber

End Property

''' <summary>   Sets the multiplexer card Slot Number. </summary>
''' <param SlotNumber="a_value">   [Integer] The Slot Number of the card. </param>
Public Property Let slotNumber(ByVal a_value As Integer)
    
    This.slotNumber = a_value

End Property

''' <summary>   Gets the scan list. </summary>
''' <value> [String]. </value>
Public Property Get ScanList() As String
        
    ScanList = This.ScanList

End Property

''' <summary>   Sets the scan list. </summary>
''' <param name="a_value">   [String] The scan list. </param>
''' <value> [String]. </value>
Public Property Let ScanList(ByVal a_value As String)
    This.ScanList = a_value
End Property

''' <summary>   Builds default scan list. </summary>
''' <remarks>   2023-08-17. </remarks>
''' <param name="measurementCode">   The measurement code. </param>
''' <returns>   A string. </returns>
Public Function BuildDefaultScanList(ByVal a_measurementCode As String) As String

    Dim firstChannelNumber As Integer
    firstChannelNumber = This.slotNumber * 100 + 1
    
    Dim lastChannelNumber As Integer
    lastChannelNumber = firstChannelNumber + This.Capacity - 1
    This.ScanList = ":FUNC '" & measurementCode & _
            "',(@" & CStr(firstChannelNumber) & _
            ":" & CStr(lastChannelNumber) & ")"
    
End Function

''' <summary>   Gets or sets the 'route' command. </summary>
''' <value> The 'route' command. </value>
Public Property Get RouteCommand() As String
    
    RouteCommand = ":ROUT:MULT:CLOS (@" & CStr(This.slotNumber) & CStr(This.Capacity + 4) & "," & _
                                          CStr(This.slotNumber) & CStr(This.Capacity + 5) & ")"

End Property

    

