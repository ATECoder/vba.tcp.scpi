VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    Adds menu items. </summary>
'''
''' The marco will not execute. Error: cannot run macro. Go Figure
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    MenuItemName As String
    MenuBarName As String
End Type

Private This As this_

Public Sub RemoveToolsMenu()
    On Error Resume Next
    Application.CommandBars(This.MenuBarName).Controls(This.MenuItemName).Delete
    On Error GoTo 0
End Sub

Public Sub AddToolsMenu(Optional ByVal a_menuItemName As String = "$isr.tools")
    Dim p_popUp As CommandBarPopup
    Dim p_button As CommandBarButton
    This.MenuBarName = "Worksheet Menu Bar"
    This.MenuItemName = a_menuItemName
    With Application.CommandBars(This.MenuBarName)
        On Error Resume Next
        .Controls(This.MenuItemName).Delete
        On Error GoTo 0
        Set p_popUp = .Controls.Add( _
            Type:=msoControlPopup, _
            Before:=.Controls.Count, _
            Temporary:=True)
    End With
    p_popUp.Caption = This.MenuItemName
    Set p_button = p_popUp.Controls.Add
    With p_button
        .Caption = "Close Editor Windows"
        .OnAction = "cc_isr_Core_IO.EditorExtensions.CloseEditorWindows"
        .Style = msoButtonCaption
    End With
    Set p_button = p_popUp.Controls.Add
    With p_button
        .Caption = "TBD"
        .OnAction = ""
        .Style = msoButtonCaption
    End With
End Sub
