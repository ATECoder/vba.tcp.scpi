VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SenseSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A generic Sense Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Session As cc_isr_Ieee488.TcpSession
    AutoRangeEnabled  As Boolean
    PowerLineCycles As Double
    SenseFunctionName As String
    SenseFunction As cc_isr_Tcp_Scpi.SenseFunctionOption
End Type

Private This As this_

''' <summary>   Values that represent sense function options. </summary>
Public Enum SenseFunctionOption
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
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As SenseSystem
    Set This.Session = a_session
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_Ieee488.TcpSession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.Session Is Nothing Then
        Connected = False
    Else
        Connected = This.Session.Connected
    End If

End Property

''' <summary>   Gets the last set or get Auto Range enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoRangeEnabled() As Boolean
    AutoRangeEnabled = This.AutoRangeEnabled
End Property

''' <summary>   Gets the Auto Range enabled. </summary>
''' <returns>   [Boolean]. </returns>
Public Function AutoRangeEnabledGetter() As Boolean
    Dim p_value As String
    p_value = This.Session.QueryLine(":" & This.SenseFunctionName & ":RANG:AUTO?")
    This.AutoRangeEnabled = IIf(p_value = "1", True, False)
    AutoRangeEnabledGetter = This.AutoRangeEnabled
End Function

''' <summary>   Sets the Auto Range enabled. </summary>
''' <param name="a_value">   [Boolean] The AutoRange enabled to set. </param >
Public Sub AutoRangeEnabledSetter(ByVal a_value As Boolean)
    This.Session.WriteLine ":" & This.SenseFunctionName & ":RANG:AUTO " & IIf(a_value, "1", "0")
    This.AutoRangeEnabled = a_value
End Sub

''' <summary>   Gets the last set or get Power Line Cycles. </summary>
''' <value>   [Double]. </value>
Public Property Get PowerLineCycles() As Double
    PowerLineCycles = This.PowerLineCycles
End Property

''' <summary>   Gets the integration rate in Power Line Cycles. </summary>
''' <returns>   [Double]. </returns>
Public Function PowerLineCyclesGetter() As Double
    
    Dim p_command As String: p_command = ":" & This.SenseFunctionName & ":NPLC?"
    Dim p_details As String
    Dim p_value As Double
    If Not cc_isr_Core.StringExtensions.TryParseDouble(This.Session.QueryLine(p_command), _
            p_value, p_details) Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".PowerLineCyclesGetter", _
            " Failed getting the sense system power line cycles using the '" & _
            p_command & "'. Details: " & p_details
    End If
    This.PowerLineCycles = p_value
    PowerLineCyclesGetter = This.PowerLineCycles
    
End Function

''' <summary>   Sets the integration rate in Power Line Cycles. </summary>
''' <param name="a_value">   [Double] The integration rate in Power Line Cyclest. </param >
Public Sub PowerLineCyclesSetter(ByVal a_value As Double)
    This.Session.WriteLine ":" & This.SenseFunctionName & ":NPLC " & VBA.CStr(a_value)
    This.PowerLineCycles = a_value
End Sub

''' <summary>   Gets the [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>]. </summary>
''' <value>   [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>]. </value>
Public Property Get SenseFunction() As cc_isr_Tcp_Scpi.SenseFunctionOption
    SenseFunction = This.SenseFunction
End Property

''' <summary>   Sets the [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>]. </summary>
''' <param name="a_value">    [<see cref="cc_isr_Tcp_Scpi.SenseFunctionOption"/>] the sense function. </param>
Friend Property Let SenseFunction(ByVal a_value As cc_isr_Tcp_Scpi.SenseFunctionOption)
    This.SenseFunction = a_value
End Property

''' <summary>   Sets the Sense Function name. </summary>
''' <value>    [String]. </value>
Public Property Get SenseFunctionName() As String
    SenseFunctionName = This.SenseFunctionName
End Property

''' <summary>   Sets the Sense Function name. </summary>
''' <param name="a_value">    [String] the sense function name. </param>
Friend Property Let SenseFunctionName(ByVal a_value As String)
    This.SenseFunctionName = a_value
End Property

''' <summary>   Gets the sense function. </summary>
''' <param name="a_decorator">   [String] the decorator to trigger from the function name. </param>
Public Function SenseFunctionGetter(Optional ByVal a_decorator As String = """") As String

    This.SenseFunctionName = This.Session.QueryLine(":SENS:FUNC?")
    This.SenseFunctionName = cc_isr_Core.StringExtensions.TrimLeft(This.SenseFunctionName, a_decorator)
    This.SenseFunctionName = cc_isr_Core.StringExtensions.TrimRight(This.SenseFunctionName, a_decorator)
    SenseFunctionGetter = This.SenseFunctionName
    
End Function

''' <summary>   Sets the sense function. </summary>
''' <param name="a_value">       [String] the sense function. </param>
''' <param name="a_decorator">   [String] the decorator to append before and after the function name. </param>
Public Sub SenseFunctionSetter(ByVal a_value As String, _
        Optional ByVal a_decorator As String = "'")

    This.SenseFunctionName = a_value
    This.Session.WriteLine ":SENS:FUNC " & a_decorator & This.SenseFunctionName & a_decorator

End Sub

