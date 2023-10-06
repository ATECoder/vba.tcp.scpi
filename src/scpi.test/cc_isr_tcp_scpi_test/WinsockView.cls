VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinsockView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The IEEE488 View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sessionTimeout As String = "B1"
Private Const m_inputAddress As String = "B2"
Private Const m_inputPort As String = "B3"
Private Const m_socketAddress As String = "C2"
Private Const m_socketId As String = "D2"
Private Const m_connected As String = "E2"

Private Const m_timerInterval As String = "C22"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Controls
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private WithEvents ConnectToggleButton As MSForms.ToggleButton
Attribute ConnectToggleButton.VB_VarHelpID = -1

Private Type this_
    SocketAddress As String
    SessionTimeout As Integer
    Observer As K2700Observer
End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Sub Class_Initialize()
    Set This.Observer = K2700Observer
End Sub

Private Sub Class_Terminate()
    Set This.Observer = Nothing
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Controls
' + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the connect toggle button. </summary>
Private Sub ConnectToggleButton_Change()
    This.Observer.OnConnect
End Sub

Public Property Let ConnectToggleExecutable(ByVal a_value As Boolean)
    If Not ConnectToggleButton Is Nothing Then
        ConnectToggleButton.Enabled = a_value
    End If
End Property

Public Property Let ConnectToggleButtonValue(ByVal a_value As Boolean)
    If Not ConnectToggleButton Is Nothing Then
        ConnectToggleButton.Value = a_value
    End If
End Property

''' <summary>   Gets the Socket Address. </summary>
''' <value>   [String]. </value>
Public Property Get SocketAddress() As String
    SocketAddress = This.SocketAddress
End Property

''' <summary>   Sets the Socket Address. </summary>
''' <param name="a_value">   [String] the message value. </value>
Friend Property Let SocketAddress(ByVal a_value As String)
    If This.SocketAddress <> a_value Then
        This.SocketAddress = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.SocketAddress = a_value
    End If
End Property

''' <summary>   Gets the socket and session timeout in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get SessionTimeout() As Integer
    SessionTimeout = This.SessionTimeout
End Property

''' <summary>   Sets the socket and session timeout in milliseconds. </summary>
''' <param name="a_value">   [Integer] the timeout value in milliseconds. </value>
Public Property Let SessionTimeout(ByVal a_value As Integer)
    If This.SessionTimeout <> a_value Then
        This.SessionTimeout = a_value
        If Not This.Observer Is Nothing Then _
            This.Observer.SessionTimeout = a_value
    End If
End Property


