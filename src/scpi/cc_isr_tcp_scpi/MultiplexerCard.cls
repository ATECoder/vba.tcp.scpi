VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiplexerCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Multiplexer card. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type ThisData
    Name As String
    Capacity As Integer
    SlotNumber As Integer
    SlotCapacity As Integer
    ScanList As String
    FirstChannelNumber As Integer
    LastChannelNumber As Integer
End Type

Private This As ThisData

''' <summary>   Initializes a new card. </summary>
''' <param name="a_name">       The card name. </param>
''' <param name="a_capacity">   The card capacity in number of channels. </param>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Function Initialize(ByVal a_name As String, ByVal a_capacity As Integer) As MultiplexerCard
    This.Capacity = a_capacity
    This.Name = a_name
    This.SlotNumber = 0
    This.SlotCapacity = 100
    This.ScanList = VBA.vbNullString
    This.FirstChannelNumber = 0
    This.LastChannelNumber = 0
    Set Initialize = Me
End Function

''' <summary>   Makes a deep copy of this object. </summary>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Function Clone() As MultiplexerCard
    
    Dim p_clone As New MultiplexerCard
    p_clone.Initialize This.Name, This.Capacity
    p_clone.SlotNumber = This.SlotNumber
    p_clone.ScanList = This.ScanList
    p_clone.FirstChannelNumber = This.FirstChannelNumber
    p_clone.LastChannelNumber = This.LastChannelNumber
    Set Clone = p_clone

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
Public Property Get SlotNumber() As Integer
        
    SlotNumber = This.SlotNumber

End Property

''' <summary>   Sets the multiplexer card Slot Number. </summary>
''' <param SlotNumber="a_value">   [Integer] The Slot Number of the card. </param>
Public Property Let SlotNumber(ByVal a_value As Integer)
    
    This.SlotNumber = a_value

End Property

''' <summary>   Gets the multiplexer card First Channel Number. </summary>
''' <value> [Integer]. </value>
Public Property Get FirstChannelNumber() As Integer
        
    FirstChannelNumber = This.FirstChannelNumber

End Property

''' <summary>   Sets the multiplexer card First Channel Number. </summary>
''' <param FirstChannelNumber="a_value">   [Integer] The First Channel Number of the card. </param>
Public Property Let FirstChannelNumber(ByVal a_value As Integer)
    
    This.FirstChannelNumber = a_value
    Me.LastChannelNumber = a_value + Me.Capacity - 1

End Property

''' <summary>   Gets the multiplexer card Last Channel Number. </summary>
''' <value> [Integer]. </value>
Public Property Get LastChannelNumber() As Integer
        
    LastChannelNumber = This.LastChannelNumber

End Property

''' <summary>   Sets the multiplexer card Last Channel Number. </summary>
''' <param LastChannelNumber="a_value">   [Integer] The Last Channel Number of the card. </param>
Public Property Let LastChannelNumber(ByVal a_value As Integer)
    
    This.LastChannelNumber = a_value

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
''' <param name="a_measurement">   The measurement function, e.g., RES. </param>
''' <returns>   A string. </returns>
Public Function BuildDefaultScanList(ByVal a_measurement As String) As String

    Dim p_firstChannelNumber As Integer
    p_firstChannelNumber = This.SlotNumber * 100 + 1
    
    Dim LastChannelNumber As Integer
    LastChannelNumber = FirstChannelNumber + This.Capacity - 1
    This.ScanList = ":FUNC '" & a_measurement & _
            "',(@" & CStr(FirstChannelNumber) & _
            ":" & CStr(LastChannelNumber) & ")"
    
End Function

''' <summary>   Gets the 'route' command. </summary>
''' <value> The 'route' command. </value>
Public Property Get RouteCommand() As String
    
    RouteCommand = ":ROUT:MULT:CLOS (@" & _
            CStr(This.SlotNumber * This.SlotCapacity + This.Capacity + 4) & "," & _
            CStr(This.SlotNumber * This.SlotCapacity + This.Capacity + 5) & ")"

End Property

''' <summary>   Gets the scan list for a channel and a slot. </summary>
''' <param name="a_slotNumber">      [Integer] The slot number. </param>
''' <param name="a_slotCapacity">    [Integer] The slot capacity for this instrument familty. </param>
''' <param name="a_channelNumber">   [Integer] The channel number within the board.
'''                                  capacity. Note that the channel number for
'''                                  the route command may exceed to card capacity
'''                                  e.g., channels 24 and 25 for 20 channels board. </param>
''' <returns> [String]. </returns>
Public Function BuildSlotScanChannel(ByVal a_slotNumber As Integer, ByVal a_slotCapacity As Integer, ByVal a_channelNumber As Integer) As String

    Dim p_channelNumber As Integer
    p_channelNumber = a_slotNumber * a_slotCapacity + a_channelNumber
    BuildSlotScanChannel = CStr(p_channelNumber)
    
End Function

''' <summary>   Gets the scan list for this card and the specified channel. </summary>
''' <param name="a_channelNumber">   [Integer] The channel number.
'''                                  must be between ths first and last channels
'''                                  of this board. </param>
''' <returns> [String]. </returns>
Public Function BuildScanChannel(ByVal a_channelNumber As Integer) As String

    If a_channelNumber < This.FirstChannelNumber Or _
        a_channelNumber > This.LastChannelNumber Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
            ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".BuildScanChannel", _
            "Argument: 'a_ChannelNumber' = " & CStr(a_channelNumber) & _
            " is out of range for this '" & This.Name & _
            "' with channels in the range [" & CStr(This.FirstChannelNumber) & _
            "," & CStr(This.LastChannelNumber) & "]."
    End If
    Dim p_channelNumber As Integer
    p_channelNumber = This.SlotNumber * This.SlotCapacity + a_channelNumber - This.FirstChannelNumber + 1
    BuildScanChannel = CStr(p_channelNumber)
    
End Function


