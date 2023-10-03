VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The K2700 View for linking the K2700 observer with actual user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sessionTimeout As String = "B1"
Private Const m_inputAddress As String = "B2"
Private Const m_inputPort As String = "B3"
Private Const m_socketAddress As String = "C2"
Private Const m_socketId As String = "D2"
Private Const m_connected As String = "E2"
Private Const m_sentMessage As String = "F2"
Private Const m_receivedMessageLength As String = "J2"
Private Const m_receivedMessage As String = "K2"
Private Const m_elapsedTime As String = "L2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_lastMessage As String = "B4"
Private Const m_lastErrorMessage As String = "B33"

Private Const m_serialPoll As String = "E11"
Private Const m_statusByte As String = "F11"
Private Const m_standardByte As String = "H11"

Private Const m_srq As String = "C10"
Private Const m_gpibAddress As String = "C12"
Private Const m_readTimeout As String = "C15"
Private Const m_AutoAssertTalk As String = "C18"

Private Const m_readingFormat As String = "C21"
Private Const m_timerInterval As String = "C22"
Private Const m_card1 As String = "C23"
Private Const m_card2 As String = "C24"
Private Const m_card1ScanList As String = "C25"
Private Const m_card2ScanList As String = "C26"
Private Const m_inputs As String = "C27"
Private Const m_channelNumber As String = "C28"
Private Const m_functionName As String = "B29"
Private Const m_functionValue As String = "C29"
Private Const m_functionUnit As String = "D29"
Private Const m_channelCount As String = "C30"
Private Const m_externalTriggerStatus As String = "C31"

' Controls

Private Type this_
    Observer As K2700Observer
End Type

Private This As this_

Private Sub Class_Initialize()
    Set This.Observer = K2700Observer
End Sub

Private Sub Class_Terminate()
    Set This.Observer = Nothing
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' K2700 Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

