VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700SenseSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   K2700 sense Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    SenseSystem As SenseSystem
    SenseFunction As cc_isr_Tcp_Scpi.K2700SenseFunction
End Type

Private This As this_

''' <summary>   Values that represent trigger sources. </summary>
Public Enum K2700SenseFunction
    ''' <summary>   An enum constant representing the none option. </summary>
    None = 0
    ''' <summary>   An enum constant representing the DC Voltage sense function. </summary>
    DcVoltage = 1
    AcVoltage = 2
    DcCurrent = 3
    AcCurrent = 4
    Resistance = 5
    FourWireResistance = 6
    Temperature = 7
    Frequency = 8
    Period = 9
    Continuity = 10
End Enum

''' <summary>   Returns a reference to the <see cref="SenseSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_Ieee488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="SenseSystem"/>]. </returns>
Public Function Initialize(ByVal a_system As SenseSystem) As K2700SenseSystem
    Set This.SenseSystem = a_system
    Set Initialize = Me
End Function

''' <summary>   Gets the Sense Function/ </summary>
Public Property Get SenseSystem() As SenseSystem
    Set SenseSystem = This.SenseSystem
End Property

''' <summary>   Gets the Sense Function/ </summary>
Public Property Get SenseFunction() As cc_isr_Tcp_Scpi.K2700SenseFunction
    SenseFunction = This.SenseFunction
End Property

''' <summary>   Sets the <see cref="K2700SenseFunction"/>. </summary>
''' <param name="p_value">   [<see cref="K2700SenseFunction"/>] the K2700 sense function. </param>
Public Sub K2700SenseFunctionSetter(ByVal p_value As cc_isr_Tcp_Scpi.K2700SenseFunction)

    Dim p_senseFunctionName As String:  p_senseFunctionName = FromK2700SenseFunction(p_value)
    This.SenseSystem.SenseFunctionSetter p_senseFunctionName, "'"
    This.SenseFunction = p_value

End Sub

''' <summary>   Gets the <see cref="K2700SenseFunction"/>. </summary>
''' <param name="p_value">   [<see cref="K2700SenseFunction"/>] the sense function. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.K2700SenseFunction"/>]. </value>
Public Function K2700SenseFunctionGetter() As cc_isr_Tcp_Scpi.K2700SenseFunction

    This.SenseFunction = ToK2700SenseFunction(This.SenseSystem.SenseFunctionGetter(""""))
    K2700SenseFunctionGetter = This.SenseFunction

End Function

''' <summary>   Converts a sense function <see cref="K2700SenseFunction"/>
''' to a sense function name. </summary>
''' <param name="p_value">   [<see cref="K2700SenseFunction"/>] the sense function. </param>
''' <returns>   [String]. </returns>
Public Function FromK2700SenseFunction(ByVal p_value As cc_isr_Tcp_Scpi.K2700SenseFunction) As String

    Select Case p_value
        Case K2700SenseFunction.DcVoltage
            FromK2700SenseFunction = "VOLT:DC"
        Case K2700SenseFunction.AcVoltage
            FromK2700SenseFunction = "VOLT:AC"
        Case K2700SenseFunction.DcCurrent
            FromK2700SenseFunction = "CURR:DC"
        Case K2700SenseFunction.AcCurrent
            FromK2700SenseFunction = "CURR:AC"
        Case K2700SenseFunction.Resistance
            FromK2700SenseFunction = "RES"
        Case K2700SenseFunction.FourWireResistance
            FromK2700SenseFunction = "FRES"
        Case K2700SenseFunction.Temperature
            FromK2700SenseFunction = "TEMP"
        Case K2700SenseFunction.Frequency
            FromK2700SenseFunction = "FREQ"
        Case K2700SenseFunction.Period
            FromK2700SenseFunction = "PER"
        Case K2700SenseFunction.Continuity
            FromK2700SenseFunction = "CONT"
        Case Else
            FromK2700SenseFunction = ""
    End Select
    
End Function

''' <summary>   Converts a sense function name to a <see cref="K2700SenseFunction"/>
''' sense function. </summary>
''' <param name="p_value">   [String] the sense function name. </param>
''' <returns>   [<see cref="cc_isr_Tcp_Scpi.K2700SenseFunction"/>]. </returns>
Public Function ToK2700SenseFunction(ByVal p_value As String) As cc_isr_Tcp_Scpi.K2700SenseFunction

    Select Case p_value
        Case "VOLT"
            ToK2700SenseFunction = K2700SenseFunction.DcVoltage
        Case "VOLT:DC"
            ToK2700SenseFunction = K2700SenseFunction.DcVoltage
        Case "VOLT:AC"
            ToK2700SenseFunction = K2700SenseFunction.AcVoltage
        Case "CURR"
            ToK2700SenseFunction = K2700SenseFunction.DcCurrent
        Case "CURR:DC"
            ToK2700SenseFunction = K2700SenseFunction.DcCurrent
        Case "CURR:AC"
            ToK2700SenseFunction = K2700SenseFunction.AcCurrent
        Case "RES"
            ToK2700SenseFunction = K2700SenseFunction.Resistance
        Case "FRES"
            ToK2700SenseFunction = K2700SenseFunction.FourWireResistance
        Case "TEMP"
            ToK2700SenseFunction = K2700SenseFunction.Temperature
        Case "FREQ"
            ToK2700SenseFunction = K2700SenseFunction.Frequency
        Case "PER"
            ToK2700SenseFunction = K2700SenseFunction.Period
        Case "CONT"
            ToK2700SenseFunction = K2700SenseFunction.Continuity
        Case Else
            ToK2700SenseFunction = K2700SenseFunction.None
    End Select
    
End Function

