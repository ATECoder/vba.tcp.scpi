VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriggerSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Trigger Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Session As cc_isr_Ieee488.TcpSession
    ContinuousEnabled As Boolean
    SampleCount As Long
    Source As cc_isr_Tcp_Scpi.TriggerSourceOption
    TriggerCount As Long
End Type

Private This As this_

''' <summary>   Values that represent trigger source options. </summary>
Public Enum TriggerSourceOption
    ''' <summary>   An enum constant representing the none trigger source. </summary>
    None = 0
    ''' <summary>   An enum constant representing the immediate trigger source. </summary>
    Immediate = 1
    ''' <summary>   An enum constant representing the external trigger source. </summary>
    External = 2
End Enum

''' <summary>   Returns a reference to the <see cref="TriggerSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_Ieee488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="TriggerSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As TriggerSystem
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Actions
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Aborts triggering. </summary>
Public Sub Abort()
    This.Session.WriteLine "ABORT"
End Sub


''' <summary>   Initiates an immediate trigger. </summary>
Public Sub Initiate()
    This.Session.WriteLine "INIT:IMM"
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Properties
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the last set or get continuous enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ContinuousEnabled() As Boolean
    ContinuousEnabled = This.ContinuousEnabled
End Property

''' <summary>   Gets the continuous enabled. </summary>
''' <returns>   [Boolean]. </returns>
Public Function ContinuousEnabledGetter() As Boolean
    Dim p_value As String
    p_value = This.Session.QueryLine("INIT:CONT?")
    This.ContinuousEnabled = IIf(p_value = "ON", True, False)
    ContinuousEnabledGetter = This.ContinuousEnabled
End Function

''' <summary>   Sets the continuous enabled. </summary>
''' <param name="a_value">   [Boolean] The continuous enabled to set. </param >
Public Sub ContinuousEnabledSetter(ByVal a_value As Boolean)
    This.Session.WriteLine "CONT " & IIf(a_value, "ON", "OFF")
    This.ContinuousEnabled = a_value
End Sub

''' <summary>   Gets the last set or get sample count. </summary>
''' <value>   [Long]. </value>
Public Property Get SampleCount() As Long
    SampleCount = This.SampleCount
End Property

''' <summary>   Gets the sample count. </summary>
''' <returns>   [Long]. </returns>
Public Function SampleCountGetter() As Long
    Dim p_command As String: p_command = "SAMP:COUN?"
    Dim p_details As String
    Dim p_value As Long
    If Not cc_isr_Core.StringExtensions.TryParseLong(This.Session.QueryLine(p_command), _
            p_value, p_details) Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SampleCountGetter", _
            " Failed getting the sample count using '" & p_command & "'. Details: " & p_details
    End If
    This.SampleCount = p_value
    SampleCountGetter = This.SampleCount
End Function

''' <summary>   Sets the sample count. </summary>
''' <param name="a_value">   [Long] The sample count to set. </param >
Public Sub SampleCountSetter(ByVal a_value As Long)
    This.Session.WriteLine "SAMP:COUN " & VBA.CStr(a_value)
    This.SampleCount = a_value
End Sub


''' <summary>   Gets the <see cref="TriggerSourceOption"/>. </summary>
''' <value>   <see cref="TriggerSourceOption"/>. </returns>
Public Property Get Source() As cc_isr_Tcp_Scpi.TriggerSourceOption
    Source = This.Source
End Property

''' <summary>   Sets the <see cref="TriggerSourceOption"/>. </summary>
''' <param name="a_value">   [<see cref="TriggerSourceOption"/>] the trigger source. </param>
Public Sub SourceSetter(ByVal a_value As cc_isr_Tcp_Scpi.TriggerSourceOption)

    If cc_isr_Tcp_Scpi.TriggerSourceOption.Immediate = a_value Then
        This.Session.WriteLine ":TRIG:SOUR IMM"
    ElseIf cc_isr_Tcp_Scpi.TriggerSourceOption.External = a_value Then
        ' for some reason the old application had the trigger source toggled
        ' if requesting an external trigger source
        ' This.Session.WriteLine ":TRIG:SOUR IMM"
        This.Session.WriteLine ":TRIG:SOUR EXT"
    End If
    This.Source = a_value
    
End Sub

''' <summary>   Gets the <see cref="TriggerSourceOption"/>. </summary>
''' <param name="a_value">   [<see cref="TriggerSourceOption"/>] the trigger source. </param>
Public Function SourceGetter() As cc_isr_Tcp_Scpi.TriggerSourceOption

    Dim p_source As String
    p_source = This.Session.QueryLine(":TRIG:SOUR?")
    Select Case p_source
        Case "IMM"
            This.Source = TriggerSourceOption.Immediate
        Case "EXT"
            This.Source = TriggerSourceOption.External
        Case Else
            This.Source = TriggerSourceOption.None
    End Select
    SourceGetter = This.Source
    
End Function

''' <summary>   Gets the last set or get Trigger count. </summary>
''' <value>   [Long]. </value>
Public Property Get TriggerCount() As Long
    TriggerCount = This.TriggerCount
End Property

''' <summary>   Gets the Trigger count. </summary>
''' <returns>   [Long]. </returns>
Public Function TriggerCountGetter() As Long
    
    Dim p_command As String: p_command = "TRIG:COUN?"
    Dim p_details As String
    Dim p_value As Long
    If Not cc_isr_Core.StringExtensions.TryParseLong(This.Session.QueryLine(p_command), _
            p_value, p_details) Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SampleCountGetter", _
            " Failed getting the sample count using '" & p_command & "'. Details: " & p_details
    End If
    
    This.TriggerCount = p_value
    TriggerCountGetter = This.TriggerCount
    
End Function

''' <summary>   Sets the Trigger count. </summary>
''' <param name="a_value">   [Long] The Trigger count to set. </param >
Public Sub TriggerCountSetter(ByVal a_value As Long)
    This.Session.WriteLine "TRIG:COUN " & VBA.CStr(a_value)
    This.TriggerCount = a_value
End Sub





