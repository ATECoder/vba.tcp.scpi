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
    ViSession As cc_isr_Ieee488.ViSession
    AutoRangeEnabled  As Boolean
    PowerLineCycles As Double
    SenseFunction As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="SenseSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>] A connected session. </param>
''' <value>   [<see cref="SenseSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As SenseSystem
    Set This.ViSession = a_session
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_IEEE488.VISession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.ViSession Is Nothing Then
        Connected = False
    Else
        Connected = This.ViSession.Connected
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
    p_value = This.ViSession.QueryLine(":" & This.SenseFunction & ":RANG:AUTO?")
    This.AutoRangeEnabled = IIf(p_value = "1", True, False)
    AutoRangeEnabledGetter = This.AutoRangeEnabled
End Function

''' <summary>   Sets the Auto Range enabled. </summary>
''' <param name="p_value">   [Boolean] The AutoRange enabled to set. </param >
Public Function AutoRangeEnabledSetter(ByVal p_value As Boolean)
    This.ViSession.WriteLine (":" & This.SenseFunction & ":RANG:AUTO " & IIf(p_value, "1", "0"))
    This.AutoRangeEnabled = p_value
End Function

''' <summary>   Gets the last set or get Power Line Cycles. </summary>
''' <value>   [Double]. </value>
Public Property Get PowerLineCycles() As Double
    PowerLineCycles = This.PowerLineCycles
End Property

''' <summary>   Gets the integration rate in Power Line Cycles. </summary>
''' <returns>   [Double]. </returns>
Public Function PowerLineCyclesGetter() As Double
    
    Dim p_command As String: p_command = ":" & This.SenseFunction & ":NPLC?"
    Dim p_details As String
    Dim p_value As Double
    If Not cc_isr_Core.StringExtensions.TryParseDouble(This.ViSession.QueryLine(p_command), _
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
''' <param name="p_value">   [Double] The integration rate in Power Line Cyclest. </param >
Public Function PowerLineCyclesSetter(ByVal p_value As Double)
    This.ViSession.WriteLine (":" & This.SenseFunction & ":NPLC " & VBA.CStr(p_value))
    This.PowerLineCycles = p_value
End Function

''' <summary>   Gets the Sense Function/ </summary>
Public Property Get SenseFunction() As String
    SenseFunction = This.SenseFunction
End Property

''' <summary>   Gets the sense function. </summary>
''' <param name="a_decorator">   [String] the decorator to trigger from the function name. </param>
Public Function SenseFunctionGetter(Optional ByVal a_decorator As String = """") As String

    This.SenseFunction = This.ViSession.QueryLine(":SENS:FUNC?")
    This.SenseFunction = cc_isr_Core.StringExtensions.TrimLeft(This.SenseFunction, a_decorator)
    This.SenseFunction = cc_isr_Core.StringExtensions.TrimRight(This.SenseFunction, a_decorator)
    SenseFunctionGetter = This.SenseFunction
    
End Function

''' <summary>   Sets the sense function. </summary>
''' <param name="p_value">       [String] the sense function. </param>
''' <param name="a_decorator">   [String] the decorator to append before and after the function name. </param>
Public Sub SenseFunctionSetter(ByVal p_value As String, _
        Optional ByVal a_decorator As String = "'")

    This.SenseFunction = p_value
    This.ViSession.WriteLine ":SENS:FUNC " & a_decorator & This.SenseFunction & a_decorator

End Sub

