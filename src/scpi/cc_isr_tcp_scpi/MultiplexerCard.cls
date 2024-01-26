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

Private Type this_
    Name As String
    Capacity As Integer
    SlotNumber As Integer
    SlotCapacity As Integer
    DeviceFirstChannel As Integer
    DeviceLastChannel As Integer
    FunctionScanList As String
    FunctionalCapacity As Integer
End Type

Private This As this_

''' <summary>   Initializes a new card. </summary>
''' <param name="a_name">       The card name. </param>
''' <param name="a_capacity">   The card capacity in number of channels. </param>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Function Initialize(ByVal a_name As String, ByVal a_capacity As Integer) As MultiplexerCard
    This.Capacity = a_capacity
    This.Name = a_name
    This.SlotNumber = 0
    This.SlotCapacity = 100
    This.DeviceFirstChannel = 0
    This.DeviceLastChannel = 0
    This.FunctionScanList = VBA.vbNull
    This.FunctionalCapacity = This.Capacity
    Set Initialize = Me
End Function

''' <summary>   Makes a deep copy of this object. </summary>
''' <returns>   [<see cref="MultiplexerCard"/>]</returns>
Public Function Clone() As MultiplexerCard
    
    Dim p_clone As New MultiplexerCard
    p_clone.Initialize This.Name, This.Capacity
    p_clone.SlotNumber = This.SlotNumber
    p_clone.DeviceFirstChannel = This.DeviceFirstChannel
    p_clone.DeviceLastChannel = This.DeviceLastChannel
    Set Clone = p_clone

End Function

''' <summary>   Gets the multiplexer card name. </summary>
''' <value>   [String]. </value>
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
''' <param name="a_value">   [Integer] The Slot Number of the card. </param>
Public Property Let SlotNumber(ByVal a_value As Integer)
    
    This.SlotNumber = a_value

End Property

''' <summary>   Gets the multiplexer functional channel count. </summary>
''' <value> [Integer]. </value>
Public Property Get FunctionalCapacity() As Integer
        
    FunctionalCapacity = This.FunctionalCapacity

End Property

''' <summary>   Gets the device-level channel number that corresponds to the
'''             multiplexer card First Channel Number. </summary>
''' <value> [Integer]. </value>
Public Property Get DeviceFirstChannel() As Integer
        
    DeviceFirstChannel = This.DeviceFirstChannel

End Property

''' <summary>   Sets the device-level channel number that corresponds to the
'''             multiplexer card first channel number. </summary>
''' <param name="a_value">   [Integer] The device channel that corresponds to
'''                          the first channel of this card. </param>
Public Property Let DeviceFirstChannel(ByVal a_value As Integer)
    
    This.DeviceFirstChannel = a_value
    This.DeviceLastChannel = a_value + Me.FunctionalCapacity - 1

End Property

''' <summary>   Gets the device-level channel number that corresponds to the
'''             multiplexer card last channel number. </summary>
''' <value> [Integer]. </value>
Public Property Get DeviceLastChannel() As Integer
        
    DeviceLastChannel = This.DeviceLastChannel

End Property

''' <summary>   Sets the device-level channel number that corresponds to the
'''             multiplexer card last channel number. </summary>
''' <param name="a_value">   [Integer] The device channel that corresponds to
'''                          the last channel of this card. </param>
Public Property Let DeviceLastChannel(ByVal a_value As Integer)
    
    This.DeviceLastChannel = a_value

End Property

''' <summary>   Gets the measurement function scan list, e.g., </summary>
''' <value>   [String]. </value>
Public Property Get FunctionScanList() As String
        
    FunctionScanList = This.FunctionScanList

End Property

''' <summary>   Builds the function scan list, e.g., <c>:FUNC 'RES',(@101,120)</c>. </summary>
''' <remarks>   2023-08-17. </remarks>
''' <param name="a_senseFunctionName">   The sense function name, e.g., RES. </param>
''' <returns>   [String]. </returns>
Public Function BuildFunctionScanList(ByVal a_senseFunctionName As String) As String

    ' update the function channel count
    This.FunctionalCapacity = IIf("FRES" = a_senseFunctionName, This.Capacity / 2, This.Capacity)
    
    Dim p_firstChannelNumber As Integer
    p_firstChannelNumber = This.SlotNumber * 100 + 1
    
    Dim p_lastChannelNumber As Integer
    p_lastChannelNumber = p_firstChannelNumber + This.FunctionalCapacity - 1
    
    This.FunctionScanList = ":FUNC '" & a_senseFunctionName & _
            "',(@" & VBA.CStr(p_firstChannelNumber) & _
            "," & VBA.CStr(p_lastChannelNumber) & ")"
            
    BuildFunctionScanList = This.FunctionScanList
    
End Function

''' <summary>   Gets the 'route' multiple close command,
'''             e.g., <c>:ROUT:MULT:CLOS (@124,125)</c>. </summary>
''' <value>   [String]. </value>
Public Property Get RouteMultipleCloseCommand() As String
        
    RouteMultipleCloseCommand = ":ROUT:MULT:CLOS (@" & _
            VBA.CStr(This.SlotNumber * This.SlotCapacity + This.Capacity + 4) & "," & _
            VBA.CStr(This.SlotNumber * This.SlotCapacity + This.Capacity + 5) & ")"

End Property

''' <summary>   Gets the scan channel for a channel and a slot. </summary>
''' <param name="a_slotNumber">           [Integer] The slot number. </param>
''' <param name="a_slotCapacity">         [Integer] The slot capacity for this instrument familty. </param>
''' <param name="a_boardChannelNumber">   [Integer] The channel number within the board.
'''                                       capacity. Note that the channel number for
'''                                       the route command may exceed to card capacity
'''                                       e.g., channels 24 and 25 for 20 channels board. </param>
''' <returns>   [String]. </returns>
Public Function GetSlotScanChannel(ByVal a_slotNumber As Integer, _
    ByVal a_slotCapacity As Integer, ByVal a_boardChannelNumber As Integer) As String

    Dim p_channelNumber As Integer
    p_channelNumber = a_slotNumber * a_slotCapacity + a_boardChannelNumber
    GetSlotScanChannel = VBA.CStr(p_channelNumber)
    
End Function

''' <summary>   Gets the scan channel, e.g., '113', for this card and the specified channel. </summary>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number.
'''                                        must be between ths first and last device channels
'''                                        that correspond tothe first and last channels
'''                                        of this board. </param>
''' <returns>   [String]. </returns>
Public Function BuildScanChannel(ByVal a_deviceChannelNumber As Integer) As String

    If a_deviceChannelNumber < This.DeviceFirstChannel Or _
        a_deviceChannelNumber > This.DeviceLastChannel Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
            ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".BuildScanChannel", _
            "Argument: 'a_deviceChannelNumber' = " & VBA.CStr(a_deviceChannelNumber) & _
            " is out of range for this '" & This.Name & _
            "' with channels in the range [" & VBA.CStr(This.DeviceFirstChannel) & _
            "," & VBA.CStr(This.DeviceLastChannel) & "]."
    End If
    Dim p_channelNumber As Integer
    p_channelNumber = This.SlotNumber * This.SlotCapacity + a_deviceChannelNumber - This.DeviceFirstChannel + 1
    BuildScanChannel = VBA.CStr(p_channelNumber)
    
End Function

''' <summary>   Gets the scan list, e.g., <c>(@113)</c>, for this card and the specified channel. </summary>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number.
'''                                        must be between ths first and last device channels
'''                                        that correspond to the first and last channels
'''                                        of this board. </param>
''' <returns>   [String]. </returns>
Public Function BuildChannelScanList(ByVal a_deviceChannelNumber As Integer) As String

    BuildChannelScanList = "(@" & Me.BuildScanChannel(a_deviceChannelNumber) & ")"
    
End Function

''' <summary>   Builds the channel 'route' close command,
'''             e.g., <c>:ROUT:CLOS (@124)</c>. </summary>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number.
'''                                        must be between ths first and last device channels
'''                                        that correspond to the first and last channels
'''                                        of this board. </param>
''' <value>   [String]. </value>
Public Function BuildChannelRouteCloseCommand(ByVal a_deviceChannelNumber As Integer) As String
        
    BuildChannelRouteCloseCommand = ":ROUT:CLOS (@" & Me.BuildScanChannel(a_deviceChannelNumber) & ")"

End Function



