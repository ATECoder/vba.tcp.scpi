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
    AutoIncrement As Boolean
    BeepEnabled As Boolean
    DutCount As Integer
    DutNumber As Integer
    FrontInputs As Boolean
    MaximumDutCount As Integer
    Mode As cc_isr_Tcp_Scpi.MeasurementModeOption
    ReadingOffset As Double
    SenseFunction As String
    SingleRead As Boolean
    TimerInterval As Integer
End Type

Private This As this_

''' <summary>   Gets the auto increment DUT number condition. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoIncrement() As Boolean
    AutoIncrement = This.AutoIncrement
End Property

''' <summary>   Sets the auto increment DUT number condition. </summary>
''' <param name="a_value">   [Boolean] True to increment the DUT number
''' upon the completion of each measurement. </value>
Public Property Let AutoIncrement(ByVal a_value As Boolean)
    This.AutoIncrement = a_value
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

''' <summary>   Gets the DUT number. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutNumber() As Integer
    DutNumber = This.DutNumber
End Property

''' <summary>   Sets the DUT number. </summary>
''' <param name="a_value">   [Integer] the DUT number value. </value>
Public Property Let DutNumber(ByVal a_value As Integer)
    This.DutNumber = a_value
End Property

''' <summary>   Gets the number of DUTs to be measured. </summary>
''' <value>   [Integer]. </value>
Public Property Get DutCount() As Integer
    DutCount = This.DutCount
End Property

''' <summary>   Sets the number of DUTs to be measured. </summary>
''' <param name="a_value">   [Integer] the number of DUT to be measured. </value>
Public Property Let DutCount(ByVal a_value As Integer)
    This.DutCount = a_value
End Property

''' <summary>   Gets the maximum number of DUTs. </summary>
''' <value>   [Integer]. </value>
Public Property Get MaximumDutCount() As Integer
    MaximumDutCount = This.MaximumDutCount
End Property

''' <summary>   Sets the maximum number of DUTs. </summary>
''' <param name="a_value">   [Integer] the maximum number of DUTs. </value>
Public Property Let MaximumDutCount(ByVal a_value As Integer)
    This.MaximumDutCount = a_value
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

''' <summary>   Gets the Reading Offset representing, for example,
''' the probe resistance or voltage bias. </summary>
''' <value>   [Double]. </value>
Public Property Get ReadingOffset() As Double
    ReadingOffset = This.ReadingOffset
End Property

''' <summary>   Sets the Reading Offset representing, for example,
''' the probe resistance or voltage bias. </summary>
''' <param name="a_value">   [Double] the Reading Offset. </value>
Public Property Let ReadingOffset(ByVal a_value As Double)
    This.ReadingOffset = a_value
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

''' <summary>   Gets the Timer Interval. </summary>
''' <value>   [Integer]. </value>
Public Property Get TimerInterval() As Integer
    TimerInterval = This.TimerInterval
End Property

''' <summary>   Sets the Timer Interval. </summary>
''' <param name="a_value">   [Integer] the Timer Interval value. </value>
Public Property Let TimerInterval(ByVal a_value As Integer)
    This.TimerInterval = a_value
End Property


