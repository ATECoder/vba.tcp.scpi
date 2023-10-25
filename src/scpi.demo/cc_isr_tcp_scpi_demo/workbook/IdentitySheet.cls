VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_repeatCountCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_receiveTimeoutCell As String = "B4"
Private Const m_errorMessageCell As String = "B5"

Private Const m_socketAddressCell As String = "C2"
Private Const m_connectedSocketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "G2"
Private Const m_receivedMessageCell As String = "H2"
Private Const m_statusByteCell As String = "F3"
Private Const m_averageElapsedTimeCell As String = "H3"
Private Const m_elapsedTimeCell As String = "H4"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_identityCell As String = "H2"

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim p_host As String: p_host = Range(m_inputHostCell).Value
    Dim p_port As Integer: p_port = Range(m_inputPortCell).Value
    Dim p_repeatCount As Integer: p_repeatCount = Range(m_repeatCountCell).Value
    Dim p_timeout As Integer: p_timeout = Range(m_receiveTimeoutCell).Value
    
    ReadIdentity p_host, p_port, p_timeout, p_repeatCount
   
End Sub


''' <summary>   Uses the 2700 Virtual instrument to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="a_host">          The host IPv4 dotted address. </param>
''' <param name="a_port">          the host port number. </param>
''' <param name="a_timeout">       session timeout. </param>
''' <param name="a_repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentity(ByVal a_host As String, ByVal a_port As Long, _
    ByVal a_timeout As Long, ByVal a_repeatCount As Integer)

    Const p_procedureName As String = "ReadIdentity"
    On Error GoTo Finally
    
    Range(m_socketAddressCell) = vbNullString
    Range(m_connectedSocketIdCell).Value = vbNullString
    Range(m_connectedCell).Value = vbNullString
    Range(m_sentMessageCell).Value = vbNullString
    Range(m_receivedMessageLengthCell).Value = vbNullString
    Range(m_receivedMessageCell).Value = vbNullString
    Range(m_elapsedTimeCell).Value = vbNullString
    Me.LastError = vbNullString

    Dim p_stopper As cc_isr_Core_IO.Stopwatch: Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    Dim p_totalMilliseconds As Double
    
    Dim p_device As K2700
    Set p_device = cc_isr_Tcp_Scpi.Factory.NewK2700.Initialize()
    
    p_device.Connectable.OpenConnection a_host & ":" & VBA.CStr(a_port), a_timeout
    
    ' report the socket number
    Range(m_connectedSocketIdCell).Value = p_device.Session.Socket.SocketId
   
    ' report the connection state
    Range(m_connectedCell).Value = p_device.Connected
    Range(m_socketAddressCell).Value = p_device.Session.SocketAddress
    
    If a_repeatCount > 0 And p_device.Connected Then
    
        Dim p_command As String: p_command = "*IDN?"
        
        Range(m_sentMessageCell).Value = p_command
        
        Dim p_loopCount As Integer
        Do While p_loopCount < a_repeatCount
            
            VBA.DoEvents
           
            p_loopCount = p_loopCount + 1
            p_stopper.Restart
            
            p_device.Device.QueryIdentity
            
            p_totalMilliseconds = p_totalMilliseconds + p_stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).Value = Len(p_device.Device.Identity)
            Range(m_receivedMessageCell).Value = p_device.Device.Identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).Value = Format(p_totalMilliseconds / p_loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).Value = Format(p_totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf a_repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).Value = "testing connect and disconnect; disconnected."
    
    Else
        Me.LastError = "connection failed without reporting an exception"
    End If
    
Finally:

    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
            cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject a_suffix:="Reading identity failed."
        End If
    End If
    
    On Error Resume Next
    If Not p_device Is Nothing Then
        If p_device.Connected Then
            p_device.Connectable.CloseConnection
        End If
    End If
    
    ' terminate the session
    Set p_device = Nothing
    
    Me.TraceErrorQueue
    
    On Error GoTo 0

End Sub

''' <summary>   Displays the last error. </summary>
Public Property Let LastError(ByVal a_value As String)
    
    Range(m_errorMessageCell).Value = a_value

End Property

''' <summary>   Dequeues and displays the
''' <see cref="cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue">errors queue</see>. </summary>
Public Sub TraceErrorQueue()
    
    If cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount > 0 Then _
        Me.LastError = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(VBA.vbCrLf)

End Sub

