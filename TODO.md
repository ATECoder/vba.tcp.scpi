# TODO

Core:
Add support for property changed events.

## Beta 202308

* 2700 View Model: Add Can Execute for all buttons using property change to efectuate,.
* figure out how to detect Sheet cell change.
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("H5")) Is Nothing Then Macro
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
	' this uses fewer resources than Intersect, which will be helpful if your worksheet changes a lot.
    IF Target.Address = "$D$2" Then
        MsgBox("Cell D2 Has Changed.")
    End If
End Sub

https://stackoverflow.com/questions/18124853/excel-vba-checkbox-click-and-change-events-the-same

 the change event also causes a click event. Use this to issue the click event and do not
 implement the click event
Private Sub CheckBox1_Change()
On Error Goto Err:
If ActiveControl.Name = CheckBox1.Name Then 
    On Error Goto 0        
    'Commands for click
    Exit Sub
Else
    On Error Goto 0 
    'Commands for change from within the Userform
    Exit Sub
Err:On Error Goto 0 
    'Commands for change from outside the Userform
End Sub    

* figure out if check box changed occurs before checkbox click.

* move button code to functions and arrange those according to topics.


* issue reset clear whenever switching model. 
* Demo:
	Use 2700 to implement all SCPI commands.
	Use trigger system to toggle the trigger source.
	? use trigger system to send the init SCPI message ':INIT ..."
* questions:
  * See if External Trigger Option Button commands can be moved to the 2700 instrument method such as
    Configure External Trigger Monitoring.


* flying:
* on device initialized
	* make sure the device is in the correct state;
	* detect if the device was disconnected. 
	* ensure the 2700 device is set to read after write false.
* issue reset clear whenever switching model. 
* add operation completion queries to settings commands.


## Tests

## Fixes

## Updates
Tests:
* Add unit tests:
	* Add 2700 unit tests:
		* connect on before all
			* determine if Prologix connected
			* how to quickly determine if instrument is connected
		* inconclusive if not connected.
		* use read raw
		* use buffer read
		* use TCP client.
MVVM:
	* use MVVM View Model for the unit test sheet
	* use MVVM View Model for the 2700 sheet.
Upload release to GitHub:
	* add deploy and localize scripts.
	* [release a build artifact asset on git hub]
	* [gh release upload]
	* [gh release create]
	
[release a build artifact asset on git hub]: release https://stackoverflow.com/questions/5207269/how-to-release-a-build-artifact-asset-on-github-with-a-script
[gh release upload]: https://cli.github.com/manual/gh_release_upload
[gh release create]: https://cli.github.com/manual/gh_release_create
