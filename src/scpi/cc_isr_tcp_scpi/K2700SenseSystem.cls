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
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="SenseSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_Ieee488.TcpSession"/>] A connected session. </param>
''' <returns>   [<see cref="SenseSystem"/>]. </returns>
Public Function Initialize(ByVal a_system As SenseSystem) As K2700SenseSystem
    Set This.SenseSystem = a_system
    Set Initialize = Me
End Function

''' <summary>   Gets the underlying Sense System/ </summary>
''' <value>   [<see cref="SenseSystem"/>]. </value>
Public Property Get SenseSystem() As SenseSystem
    Set SenseSystem = This.SenseSystem
End Property

''' <summary>   Sets the <see cref="SenseFunctionOption"/>. </summary>
''' <param name="a_value">   [<see cref="SenseFunctionOption"/>] the sense function option. </param>
Public Sub SenseFunctionSetter(ByVal a_value As cc_isr_Tcp_Scpi.SenseFunctionOption)

    Dim p_senseFunctionName As String: p_senseFunctionName = Me.FromSenseFunction(a_value)
    This.SenseSystem.SenseFunctionSetter p_senseFunctionName, "'"
    This.SenseSystem.SenseFunction = a_value

End Sub

''' <summary>   Gets the <see cref="SenseFunctionOption"/>. </summary>
''' <param name="a_value">   [<see cref="SenseFunctionOption"/>] the sense function option. </param>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>]. </value>
Public Function SenseFunctionGetter() As cc_isr_Tcp_Scpi.SenseFunctionOption

    This.SenseSystem.SenseFunction = ToSenseFunction(This.SenseSystem.SenseFunctionGetter(""""))
    SenseFunctionGetter = This.SenseSystem.SenseFunction

End Function

''' <summary>   Converts a sense function <see cref="SenseFunctionOption"/>
''' to a sense function name. </summary>
''' <param name="a_value">   [<see cref="SenseFunctionOption"/>] the sense function. </param>
''' <returns>   [String]. </returns>
Public Function FromSenseFunction(ByVal a_value As cc_isr_Tcp_Scpi.SenseFunctionOption) As String

    Select Case a_value
        Case SenseFunctionOption.DcVoltage
            FromSenseFunction = "VOLT:DC"
        Case SenseFunctionOption.AcVoltage
            FromSenseFunction = "VOLT:AC"
        Case SenseFunctionOption.DcCurrent
            FromSenseFunction = "CURR:DC"
        Case SenseFunctionOption.AcCurrent
            FromSenseFunction = "CURR:AC"
        Case SenseFunctionOption.Resistance
            FromSenseFunction = "RES"
        Case SenseFunctionOption.FourWireResistance
            FromSenseFunction = "FRES"
        Case SenseFunctionOption.Temperature
            FromSenseFunction = "TEMP"
        Case SenseFunctionOption.Frequency
            FromSenseFunction = "FREQ"
        Case SenseFunctionOption.Period
            FromSenseFunction = "PER"
        Case SenseFunctionOption.Continuity
            FromSenseFunction = "CONT"
        Case Else
            FromSenseFunction = ""
    End Select
    
End Function

''' <summary>   Converts a sense function name to a <see cref="SenseFunctionOption"/>
''' sense function. </summary>
''' <param name="a_value">   [String] the sense function name. </param>
''' <returns>   [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>]. </returns>
Public Function ToSenseFunction(ByVal a_value As String) As cc_isr_Tcp_Scpi.SenseFunctionOption

    Select Case a_value
        Case "VOLT"
            ToSenseFunction = SenseFunctionOption.DcVoltage
        Case "VOLT:DC"
            ToSenseFunction = SenseFunctionOption.DcVoltage
        Case "VOLT:AC"
            ToSenseFunction = SenseFunctionOption.AcVoltage
        Case "CURR"
            ToSenseFunction = SenseFunctionOption.DcCurrent
        Case "CURR:DC"
            ToSenseFunction = SenseFunctionOption.DcCurrent
        Case "CURR:AC"
            ToSenseFunction = SenseFunctionOption.AcCurrent
        Case "RES"
            ToSenseFunction = SenseFunctionOption.Resistance
        Case "FRES"
            ToSenseFunction = SenseFunctionOption.FourWireResistance
        Case "TEMP"
            ToSenseFunction = SenseFunctionOption.Temperature
        Case "FREQ"
            ToSenseFunction = SenseFunctionOption.Frequency
        Case "PER"
            ToSenseFunction = SenseFunctionOption.Period
        Case "CONT"
            ToSenseFunction = SenseFunctionOption.Continuity
        Case Else
            ToSenseFunction = SenseFunctionOption.None
    End Select
    
End Function

