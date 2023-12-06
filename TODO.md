# TODO

## GIT

## TODO

remove:
    DutNumberCaptionPrefix As String
    DutTitle As String
	
	from user sheet and views.
	get them only from the data sheet.
	
        
Add to data sheet
    DutTitle As String
        
''' <summary>   Gets a DUT number from the DUT number caption. </summary>
''' <param name="a_value">   [String] the selected DUT number caption, e.g., R2. </value>
''' <param name="a_details">   [Out, String] Details the failure information if any. </param>
''' <returns>   [Integer] Returns 0 if invalid DUT number caption was selected. </returns>
Public Function TryParseSelectedDutNumber(ByVal a_value As String, ByRef a_details As String) As Integer

    Const p_procedureName = "TryParseSelectedDutNumber"
    
    TryParseSelectedDutNumber = This.ViewModel.TryParseDutNumberCaption(a_value, _
        Me.DutNumberCaptionPrefix, Me.DutTitle, a_details)
    
End Function

''' <summary>   Parses the DUT number value from the value of the active cell. </summary>
''' <param name="a_maximumDutNumber">   [Integer] the maximum DUT number. </param>
''' <returns>   [Integer] Returns 0 if invalid resistance cell was selected. </returns>
Public Function GetActiveCellDutNumber(ByVal a_maximumDutNumber As Integer) As Integer

    Const p_procedureName = "GetActiveCellDutNumber"
    
    Dim p_details As String: p_details = VBA.vbNullString
    Dim p_dutNumber As Integer
    p_dutNumber = UserView.TryParseSelectedDutNumber(VBA.UCase$(ActiveCell.Value), _
        a_maximumDutNumber, p_details)
    If 0 = p_dutNumber Then
        MsgBox p_details, VBA.VbMsgBoxStyle.vbOKOnly Or VBA.VbMsgBoxStyle.vbExclamation, _
            "Invalid device under test number"
    End If
    
    GetActiveCellDutNumber = p_dutNumber
  
End Function


replace is valid channel number with view model is valid dut number

Remove Selectws Dut Number from the user sheet and use Get and Set Active Cell Dut Number

''' <summary>   Gets the selected DUT number from the active cell value. </summary>
''' <value>   [Integer]. </value>
Public Property Get SelectedDutNumber() As Integer
    SelectedDutNumber = Me.GetActiveCellDutNumber
End Property

user view:
Public Property Let SelectedDutNumber(ByVal a_value As Integer)
    
    If This.SelectedDutNumber <> a_value Then
        
        This.SelectedDutNumber = a_value
        
        ' emulate selecting a DUT number if within range.
        If Me.IsValidChannelNumber(a_value) Then _
            Me.SelectedDutNumberCaption = This.DutNumberCaptionPrefix & VBA.CStr(a_value)
        
        If Not This.UserSheet Is Nothing Then _
            This.UserSheet.SetActiveDutNumberCell a_value

    End If

End Property

        p_mode.DutNumber = This.UserSheet.GetActiveCellDutNumber(p_mode.DutCount)

user view:
Public Property Let SelectedDutNumber(ByVal a_value As Integer)
    
    If This.SelectedDutNumber <> a_value Then
        
        This.SelectedDutNumber = a_value
        
        ' emulate selecting a DUT number if within range.
        If a_value < 0 And a_value < This.DataSheet.MaxDutCount Then _
            Me.SelectedDutNumberCaption = This.DutNumberCaptionPrefix & VBA.CStr(a_value)
        
        If a_value < 0 And a_value < This.DataSheet.MaxDutCount And _
            Not This.UserSheet Is Nothing Then _
            This.UserSheet.SetActiveDutNumberCell a_value

    End If

End Property

## Tests

## Fixes

## Updates

Upload release to GitHub:
	* add deploy and localize scripts.
	* [release a build artifact asset on git hub]
	* [gh release upload]
	* [gh release create]
	
[release a build artifact asset on git hub]: release https://stackoverflow.com/questions/5207269/how-to-release-a-build-artifact-asset-on-github-with-a-script
[gh release upload]: https://cli.github.com/manual/gh_release_upload
[gh release create]: https://cli.github.com/manual/gh_release_create
