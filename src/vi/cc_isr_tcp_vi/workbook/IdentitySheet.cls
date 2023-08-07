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

''' <summary>   Initialize this worksheet. </summary>
Public Sub Initialize()
End Sub

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim p_host As String: p_host = Range(m_inputHostCell).value
    Dim p_port As Integer: p_port = Range(m_inputPortCell).value
    Dim p_repeatCount As Integer: p_repeatCount = Range(m_repeatCountCell).value
    Dim p_timeout As Integer: p_timeout = Range(m_receiveTimeoutCell).value
    
    ReadIdentity p_host, p_port, p_repeatCount
   
End Sub


''' <summary>   Uses the 2700 Virtual instrument to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="a_host">   The host IPv4 dotted address. </param>
''' <param name="a_port">   the host port number. </param>
''' <param name="a_repeatCount">   The number of time to repeat the query. </param>
Private Sub ReadIdentity(ByVal a_host As String, ByVal a_port As Long, ByVal a_repeatCount As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = vbNullString
    Range(m_connectedSocketIdCell).value = vbNullString
    Range(m_connectedCell).value = vbNullString
    Range(m_sentMessageCell).value = vbNullString
    Range(m_receivedMessageLengthCell).value = vbNullString
    Range(m_receivedMessageCell).value = vbNullString
    Range(m_elapsedTimeCell).value = vbNullString
    Range(m_errorMessageCell).value = vbNullString

    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Factory.NewStopWatch()
    Dim p_totalMilliseconds As Double
    
    Dim p_device As New K2700
    
    p_device.OpenConnection a_host, a_port
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = p_device.Session.VISession.SocketId
   
    ' report the connection state
    Range(m_connectedCell).value = p_device.Connected
    
    If a_repeatCount > 0 And p_device.Connected Then
    
        Dim p_command As String: p_command = "*IDN?"
        
        Range(m_sentMessageCell).value = p_command
        
        Dim p_loopCount As Integer
        Do While p_loopCount < a_repeatCount
            
            DoEvents
           
            p_loopCount = p_loopCount + 1
            p_stopper.Restart
            
            p_device.Session.QueryIdentity
            
            p_totalMilliseconds = p_totalMilliseconds + p_stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(p_device.Session.Identity)
            Range(m_receivedMessageCell).value = p_device.Session.Identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(p_totalMilliseconds / p_loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(p_totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf a_repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
    
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim p_errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            p_errorMessage = "Reading identity failed using Tcp session " & UserDefinedErrors.FormattedErr & _
                             " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    If Not p_device Is Nothing Then
        If p_device.Connected Then
            p_device.CloseConnection
        End If
    End If
    
    ' terminate the session
    Set p_device = Nothing
    
    If Len(p_errorMessage) > 0 Then
        Range(m_errorMessageCell).value = p_errorMessage
    End If
    
    On Error GoTo 0

End Sub


