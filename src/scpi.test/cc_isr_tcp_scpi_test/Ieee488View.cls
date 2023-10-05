VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The IEEE488 View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sentMessage As String = "F2"
Private Const m_receivedMessageLength As String = "J2"
Private Const m_receivedMessage As String = "K2"
Private Const m_elapsedTime As String = "L2"
Private Const m_elapsedTimeFormat As String = "0.0"

Private Const m_serialPoll As String = "E11"
Private Const m_statusByte As String = "F11"
Private Const m_standardByte As String = "H11"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Controls
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private WithEvents CLSButton As MSForms.CommandButton
Attribute CLSButton.VB_VarHelpID = -1
Private WithEvents CommandsComboBox As MSForms.ComboBox
Attribute CommandsComboBox.VB_VarHelpID = -1
Private WithEvents QueryButton As MSForms.CommandButton
Attribute QueryButton.VB_VarHelpID = -1
Private WithEvents QueryStatsuByteButton As MSForms.CommandButton
Attribute QueryStatsuByteButton.VB_VarHelpID = -1
Private WithEvents ReadButton As MSForms.CommandButton
Attribute ReadButton.VB_VarHelpID = -1
Private WithEvents ReadStandardEventStatusButton As MSForms.CommandButton
Attribute ReadStandardEventStatusButton.VB_VarHelpID = -1
Private WithEvents ReadStatusAfterWriteCheckBox As MSForms.CheckBox
Attribute ReadStatusAfterWriteCheckBox.VB_VarHelpID = -1
Private WithEvents RSTButton As MSForms.CommandButton
Attribute RSTButton.VB_VarHelpID = -1
Private WithEvents WriteButton As MSForms.CommandButton
Attribute WriteButton.VB_VarHelpID = -1

Private Type this_
    ReceivedMessage As String
    SentMessage As String
    StatusByte As Integer
    StandardByte As Integer
    ScpiCommands As Collection
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
' IEEE488 Controls
' + + + + + + + + + + + + + + + + + + + + + + + + + + + +

Private Sub CLSButton_Click()
    This.Observer.OnClearExecutionState
End Sub

Public Property Let CLSExecutable(ByVal a_value As Boolean)
    If Not CLSButton Is Nothing Then
        CLSButton.Enabled = a_value
    End If
End Property

''' <summary>   Selects the command. </summary>
Private Sub CommandsComboBox_Change()
    This.Observer.MessageToSend = CommandsComboBox.Text
End Sub

Public Property Let CommandsComboBoxEnabled(ByVal a_value As Boolean)
    If Not CommandsComboBox Is Nothing Then
        CommandsComboBox.Enabled = a_value
    End If
End Property

Private Sub QueryButton_Click()
    This.Observer.OnQuery
End Sub

Public Property Let QueryExecutable(ByVal a_value As Boolean)
    If Not QueryButton Is Nothing Then
        QueryButton.Enabled = a_value
    End If
End Property

Private Sub QueryStatsuByteButton_Click()
    This.Observer.OnQueryStatusByte
End Sub

Public Property Let QueryStatsuByteExecutable(ByVal a_value As Boolean)
    If Not QueryStatsuByteButton Is Nothing Then
        QueryStatsuByteButton.Enabled = a_value
    End If
End Property

Private Sub ReadButton_Click()
    This.Observer.OnRead
End Sub

Public Property Let ReadExecutable(ByVal a_value As Boolean)
    If Not ReadButton Is Nothing Then
        ReadButton.Enabled = a_value
    End If
End Property

Private Sub ReadStandardEventStatusButton_Click()
    This.Observer.OnReadStandardEventStatus
End Sub

Public Property Let ReadStandardEventStatusExecutable(ByVal a_value As Boolean)
    If Not ReadStandardEventStatusButton Is Nothing Then
        ReadStandardEventStatusButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the read status after write check box change and click events. </summary>
Private Sub ReadStatusAfterWriteCheckBox_Change()
    This.Observer.ReadStatusAfterWriteEnabled = ReadStatusAfterWriteCheckBox.Value
End Sub

Public Property Let ReadStatusAfterWriteCheckBoxEnabled(ByVal a_value As Boolean)
    If Not ReadStatusAfterWriteCheckBox Is Nothing Then
        ReadStatusAfterWriteCheckBox.Enabled = a_value
    End If
End Property

Public Property Let ReadStatusAfterWriteEnabled(ByVal a_value As Boolean)
    If Not ReadStatusAfterWriteCheckBox Is Nothing Then
        ReadStatusAfterWriteCheckBox.Value = a_value
    End If
End Property

Private Sub RSTButton_Click()
    This.Observer.OnResetKnownState
End Sub

Public Property Let RSTExecutable(ByVal a_value As Boolean)
    If Not RSTButton Is Nothing Then
        RSTButton.Enabled = a_value
    End If
End Property

Private Sub WriteButton_Click()
    This.Observer.OnWrite
End Sub

Public Property Let WriteExecutable(ByVal a_value As Boolean)
    If Not WriteButton Is Nothing Then
        WriteButton.Enabled = a_value
    End If
End Property

''' <summary>   Gets the Sent Message. </summary>
''' <value>   [String]. </value>
Public Property Get SentMessage() As String
    SentMessage = This.SentMessage
End Property

''' <summary>   Sets the Sent Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let SentMessage(ByVal a_value As String)
    If This.SentMessage <> a_value Then
        This.SentMessage = a_value
    End If
End Property


''' <summary>   Gets the Received Message. </summary>
''' <value>   [String]. </value>
Public Property Get ReceivedMessage() As String
    ReceivedMessage = This.ReceivedMessage
End Property

''' <summary>   Sets the Received Message. </summary>
''' <param name="a_value">   [String] the message value. </value>
Public Property Let ReceivedMessage(ByVal a_value As String)
    If This.ReceivedMessage <> a_value Then
        This.ReceivedMessage = a_value
    End If
End Property


''' <summary>   Gets the Status Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StatusByte() As Integer
    StatusByte = This.StatusByte
End Property

''' <summary>   Sets the Status Byte. </summary>
''' <param name="a_value">   [Integer] the Status Byte value. </value>
Public Property Let StatusByte(ByVal a_value As Integer)
    If This.StatusByte <> a_value Then
        This.StatusByte = a_value
    End If
End Property

''' <summary>   Gets the Standard Byte. </summary>
''' <value>   [Integer]. </value>
Public Property Get StandardByte() As Integer
    StandardByte = This.StandardByte
End Property

''' <summary>   Sets the Standard Byte. </summary>
''' <param name="a_value">   [Integer] the Standard Byte value. </value>
Public Property Let StandardByte(ByVal a_value As Integer)
    If This.StandardByte <> a_value Then
        This.StandardByte = a_value
    End If
End Property

''' <summary>   Gets the Scpi Commands. </summary>
''' <value>   [Collection]. </value>
Public Property Get ScpiCommands() As Collection
    Set ScpiCommands = This.ScpiCommands
End Property

''' <summary>   Sets the Scpi Commands. </summary>
''' <param name="a_value">   [Collection] the Scpi Commands value. </value>
Public Property Let ScpiCommands(ByVal a_value As Collection)
    If Not cc_isr_Core.CollectionExtensions.AreEqual(This.ScpiCommands, a_value) Then
        Set This.ScpiCommands = a_value
        If Not a_value Is Nothing And Not CommandsComboBox Is Nothing Then
            CommandsComboBox.Clear
            Dim p_item As Variant
            For Each p_item In a_value
                CommandsComboBox.AddItem VBA.CStr(p_item)
            Next
        End If
    End If
End Property


