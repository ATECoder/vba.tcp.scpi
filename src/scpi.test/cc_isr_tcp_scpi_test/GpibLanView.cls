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

Public Property Let AutoAssertTalkGetButtonEnabled(ByVal a_value As Boolean)
    If Not AutoAssertTalkGetButton Is Nothing Then
        AutoAssertTalkGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Auto Asset TALK Set button click event. </summary>
Private Sub AutoAssertTalkSetButton_Click()
    This.Observer.OnAutoAssertTalkSet
End Sub

Public Property Let AutoAssertTalkSetButtonEnabled(ByVal a_value As Boolean)
    If Not AutoAssertTalkSetButton Is Nothing Then
        AutoAssertTalkSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Go To Local button click event. </summary>
Private Sub GoToLocalButton_Click()
    This.Observer.OnGoToLocal
End Sub

Public Property Let GoToLocalButtonEnabled(ByVal a_value As Boolean)
    If Not GoToLocalButton Is Nothing Then
        GoToLocalButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the GPIB Address Get button click event. </summary>
Private Sub GpibAddressGetButton_Click()
    This.Observer.OnGpibAddressGet
End Sub

Public Property Let GpibAddressGetButtonEnabled(ByVal a_value As Boolean)
    If Not GpibAddressGetButton Is Nothing Then
        GpibAddressGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the GPIB Address Set button click event. </summary>
Private Sub GpibAddressSetButton_Click()
    This.Observer.OnGpibAddressSet
End Sub

Public Property Let GpibAddressSetButtonEnabled(ByVal a_value As Boolean)
    If Not GpibAddressSetButton Is Nothing Then
        GpibAddressSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the local lockout button click event. </summary>
Private Sub LocalLockoutButton_Click()
    This.Observer.OnLocalLockout
End Sub

Public Property Let LocalLockoutButtonEnabled(ByVal a_value As Boolean)
    If Not LocalLockoutButton Is Nothing Then
        LocalLockoutButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Timeout get button click event. </summary>
Private Sub ReadTimeoutGetButton_Click()
    This.Observer.OnReadTimeoutGet
End Sub

Public Property Let ReadTimeoutGetButtonEnabled(ByVal a_value As Boolean)
    If Not ReadTimeoutGetButton Is Nothing Then
        ReadTimeoutGetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Timeout set button click event. </summary>
Private Sub ReadTimeoutSetButton_Click()
    This.Observer.OnReadTimeoutSet
End Sub

Public Property Let ReadTimeoutSetButtonEnabled(ByVal a_value As Boolean)
    If Not ReadTimeoutSetButton Is Nothing Then
        ReadTimeoutSetButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Read Status Byte (serial poll) button click event. </summary>
Private Sub ReadStatusByteButton_Click()
    This.Observer.OnReadStatusByte
End Sub

Public Property Let ReadStatusByteButtonEnabled(ByVal a_value As Boolean)
    If Not ReadStatusByteButton Is Nothing Then
        ReadStatusByteButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    This.Observer.OnSDC
End Sub

Public Property Let SDCButtonEnabled(ByVal a_value As Boolean)
    If Not SDCButton Is Nothing Then
        SDCButton.Enabled = a_value
    End If
End Property

''' <summary>   Handles the SRQ button click event. </summary>
Private Sub SRQButton_Click()
    This.Observer.OnSRQ
End Sub

Public Property Let SRQButtonEnabled(ByVal a_value As Boolean)
    If Not SRQButton Is Nothing Then
        SRQButton.Enabled = a_value
    End If
End Property

