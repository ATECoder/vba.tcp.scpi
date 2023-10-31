VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeasureMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A public creatable measure definition. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Values that represent mesaurement mode options. </summary>
Public Enum MeasurementModeOption
    ''' <summary>   An enum constant representing the none option. </summary>
    None = 0
    ''' <summary>   An enum constant representing the continuous trigger option. </summary>
    Continuous = 1
    ''' <summary>   An enum constant representing the immediate trigger option. </summary>
    Immediate = 2
    ''' <summary>   An enum constant representing the external trigger option. </summary>
    External = 3
    ''' <summary>   An enum constant representing the external trigger monitoring option. </summary>
    Monitoring = 4
End Enum

Private Type this_
    AutoIncrementChannel As Boolean
    BeepEnabled As Boolean
    ChannelNumber As Integer
    FrontInputs As Boolean
    Mode As cc_isr_Tcp_Scpi.MeasurementModeOption
    SenseFunction As String
    TimerInterval As Integer
    SingleRead As Boolean
End Type

Private This As this_

''' <summary>   Gets the Auto Increment Channel Number condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrementChannel() As Boolean
    AutoIncrementChannel = This.AutoIncrementChannel
End Property

''' <summary>   Sets the Auto Increment Channel Number condition. </summary>
''' <param name="a_value">   [Boolean] the Auto Increment Channel Number condition value. </value>
Public Property Let AutoIncrementChannel(ByVal a_value As Boolean)
    This.AutoIncrementChannel = a_value
End Property

''' <summary>   Gets the beep enabled on read condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get BeepEnabled() As Boolean
    BeepEnabled = This.BeepEnabled
End Property

''' <summary>   Sets the beep enabled on read condition. </summary>
''' <param name="a_value">   [Boolean] true if a beep sound should be issued on reading. </value>
Public Property Let BeepEnabled(ByVal a_value As Boolean)
    This.BeepEnabled = a_value
End Property


''' <summary>   Gets the Channel Number. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelNumber() As Integer
    ChannelNumber = This.ChannelNumber
End Property

''' <summary>   Sets the Channel Number. </summary>
''' <param name="a_value">   [Integer] the Channel Number value. </value>
Public Property Let ChannelNumber(ByVal a_value As Integer)
    This.ChannelNumber = a_value
End Property

''' <summary>   Gets the Front Inputs condition. </summary>
''' <remarks>   True if front inputs; otherwise, Rear inputs. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get FrontInputs() As Boolean
    FrontInputs = This.FrontInputs
End Property

''' <summary>   Sets the Front Inputs condition. </summary>
''' <param name="a_value">   [Boolean] true if Front Inputs; otherwise Rear inputs. </value>
Public Property Let FrontInputs(ByVal a_value As Boolean)
    This.FrontInputs = a_value
End Property

''' <summary>   Gets the measurement mode. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Get Mode() As cc_isr_Tcp_Scpi.MeasurementModeOption
    Mode = This.Mode
End Property

''' <summary>   Sets the measurement mode. </summary>
''' <param name=-"a_value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>] the value. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.MeasurementModeOption"/>]. </value>
Public Property Let Mode(ByVal a_value As cc_isr_Tcp_Scpi.MeasurementModeOption)
    This.Mode = a_value
End Property

''' <summary>   Gets the name of the sense function. </summary>
''' <value>   [String]. </value>
Public Property Get SenseFunction() As String
    SenseFunction = This.SenseFunction
End Property

''' <summary>   Sets the name of the sense function. </summary>
''' <param name="a_value">   [String] the name of the sense function. </value>
Public Property Let SenseFunction(ByVal a_value As String)
    This.SenseFunction = a_value
End Property

''' <summary>   Gets the Single Read condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get SingleRead() As Boolean
    SingleRead = This.SingleRead
End Property

''' <summary>   Sets the Single Read condition. </summary>
''' <param name="a_value">   [Boolean] true if a single read; otherwise, multiple. </value>
Public Property Let SingleRead(ByVal a_value As Boolean)
    This.SingleRead = a_value
End Property

