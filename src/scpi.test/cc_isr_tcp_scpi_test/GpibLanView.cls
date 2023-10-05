VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GpibLanView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The GPIB-Lan View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_serialPoll As String = "E11"
Private Const m_statusByte As String = "F11"
Private Const m_srq As String = "C10"
Private Const m_gpibAddress As String = "C12"
Private Const m_readTimeout As String = "C15"
Private Const m_AutoAssertTalk As String = "C18"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Controls
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private WithEvents AutoAssertTalkGetButton As MSForms.CommandButton
Attribute AutoAssertTalkGetButton.VB_VarHelpID = -1
Private WithEvents AutoAssertTalkSetButton As MSForms.CommandButton
Attribute AutoAssertTalkSetButton.VB_VarHelpID = -1
Private WithEvents GoToLocalButton As MSForms.CommandButton
Attribute GoToLocalButton.VB_VarHelpID = -1
Private WithEvents GpibAddressGetButton As MSForms.CommandButton
Attribute GpibAddressGetButton.VB_VarHelpID = -1
Private WithEvents GpibAddressSetButton As MSForms.CommandButton
Attribute GpibAddressSetButton.VB_VarHelpID = -1
Private WithEvents LocalLockoutButton As MSForms.CommandButton
Attribute LocalLockoutButton.VB_VarHelpID = -1
Private WithEvents ReadTimeoutGetButton As MSForms.CommandButton
Attribute ReadTimeoutGetButton.VB_VarHelpID = -1
Private WithEvents ReadTimeoutSetButton As MSForms.CommandButton
Attribute ReadTimeoutSetButton.VB_VarHelpID = -1
Private WithEvents ReadStatusByteButton As MSForms.CommandButton
Attribute ReadStatusByteButton.VB_VarHelpID = -1
Private WithEvents SDCButton As MSForms.CommandButton
Attribute SDCButton.VB_VarHelpID = -1
Private WithEvents SRQButton As MSForms.CommandButton
Attribute SRQButton.VB_VarHelpID = -1

Private Type this_
    AutoAssertTalk As Boolean
    GpibLanControllerPort As Long
    StatusByte As Integer
    ServiceRequested As Integer
    GpibAddress As Integer
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

''' <summary>   Handles the Auto Asset TALK Get button click event. </summary>
Private Sub AutoAssertTalkGetButton_Click()
    This.Observer.OnAutoAssertTalkGet
End Sub

Public Property Let AutoAssertTalkGetExecutable(ByVal a_value As Boolean)
    If Not AutoAssertTalkGetButton Is Nothing Then
        AutoAssertTalkGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Auto Asset TALK Set button click event. </summary>
Private Sub AutoAssertTalkSetButton_Click()
    This.Observer.OnAutoAssertTalkSet
End Sub

Public Property Let AutoAssertTalkSetExecutable(ByVal a_value As Boolean)
    If Not AutoAssertTalkSetButton Is Nothing Then
        AutoAssertTalkSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()
    This.Observer.OnGoToLocal
End Sub

Public Property Let GoToLocalExecutable(ByVal a_value As Boolean)
    If Not GoToLocalButton Is Nothing Then
        GoToLocalButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GpibAddressGetButton_Click()
    This.Observer.OnGpibAddressGet
End Sub

Public Property Let GpibAddressGetExecutable(ByVal a_value As Boolean)
    If Not GpibAddressGetButton Is Nothing Then
        GpibAddressGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the GPIB Address Set button click event. </summary>
Private Sub GpibAddressSetButton_Click()
    This.Observer.OnGpibAddressSet
End Sub

Public Property Let GpibAddressSetExecutable(ByVal a_value As Boolean)
    If Not GpibAddressSetButton Is Nothing Then
        GpibAddressSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()
    This.Observer.OnLocalLockout
End Sub

Public Property Let LocalLockoutExecutable(ByVal a_value As Boolean)
    If Not LocalLockoutButton Is Nothing Then
        LocalLockoutButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()
    This.Observer.OnReadTimeoutGet
End Sub

Public Property Let ReadTimeoutGetExecutable(ByVal a_value As Boolean)
    If Not ReadTimeoutGetButton Is Nothing Then
        ReadTimeoutGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()
    This.Observer.OnReadTimeoutSet
End Sub

Public Property Let ReadTimeoutSetExecutable(ByVal a_value As Boolean)
    If Not ReadTimeoutSetButton Is Nothing Then
        ReadTimeoutSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Status Byte (serial poll) button click event. </summary>
Private Sub ReadStatusByteButton_Click()
    This.Observer.OnReadStatusByte
End Sub

Public Property Let ReadStatusByteExecutable(ByVal a_value As Boolean)
    If Not ReadStatusByteButton Is Nothing Then
        ReadStatusByteButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    This.Observer.OnSDC
End Sub

Public Property Let SDCExecutable(ByVal a_value As Boolean)
    If Not SDCButton Is Nothing Then
        SDCButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()
    This.Observer.OnSRQ
End Sub

Public Property Let SRQExecutable(ByVal a_value As Boolean)
    If Not SRQButton Is Nothing Then
        SRQButton.Enabled = a_value
    End If
End Property

''' <summary>   Gets the Read After Write Enabled. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Sets the Read After Write Enabled. </summary>
''' <param name="a_value">   [Boolean] the value. </value>
Public Property Let AutoAssertTalk(ByVal a_value As Boolean)
    If This.AutoAssertTalk <> a_value Then
        This.AutoAssertTalk = a_value
    End If
End Property

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    If This.GpibLanControllerPort <> a_value Then
        This.GpibLanControllerPort = a_value
        If Not This.Observer Is Nothing Then _
                This.Observer.GpibLanControllerPort = a_value
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
    If This.StatusByte <> a_value Then _
        This.StatusByte = a_value
End Property

''' <summary>   Gets the Service Requested. </summary>
''' <value>   [Integer]. </value>
Public Property Get ServiceRequested() As Integer
    ServiceRequested = This.ServiceRequested
End Property

''' <summary>   Sets the Service Requested. </summary>
''' <param name="a_value">   [Integer] the Service Requested value. </value>
Public Property Let ServiceRequested(ByVal a_value As Integer)
    If This.ServiceRequested <> a_value Then _
        This.ServiceRequested = a_value
End Property

''' <summary>   Gets the Gpib Address. </summary>
''' <value>   [Integer]. </value>
Public Property Get GpibAddress() As Integer
    GpibAddress = This.GpibAddress
End Property

''' <summary>   Sets the Gpib Address. </summary>
''' <param name="a_value">   [Integer] the Gpib Address value. </value>
Public Property Let GpibAddress(ByVal a_value As Integer)
    If This.GpibAddress <> a_value Then
        This.GpibAddress = a_value
    End If
End Property
