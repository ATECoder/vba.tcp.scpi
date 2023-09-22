# TODO
=IF(LEN(C4)>0,"ERR                               ","Module")


https://groups.io/g/HP-Agilent-Keysight-equipment/topic/86224398

## GIT:

## TODO

? Add tests for resetting the gpib-lan controller.

SCPI:
Update the code and ran all tests.

Flying
Power on reset the GPIB Lan device once each session.
-- Store the Gpib Lan power on reset condition at the IEEE488 workbook level, e.g., in a Project Singleton Settings Class.


Must use +AUTO MODE at 0 otherwise serial poll does not work. 
Change await reading to await device reading
Add Await Controller Reading.
Set gpib-lan timeout to session timeout. 
Set the limits to the GPIB-Lan timeout (1, 3000)
Reset the controller on connection changed. 
After reset test toggling timeout to see that we are getting the correct values.





It looks like the GPIB Lan device is failing to return the correct value.
We need to run more unit tests on the GPIB Lan device testing getting Read after write and serial poll consistently.



the instrument is left with 16/n  on the output. possible a serial pool reading.


!! it looks like we have an issue with reading device errors and that the serial poll is left handling in the queue. 
possibly we need a longer timeout to read the serial poll from the GPIB Lan.

!!! trace using device test where the leftover reading comes from.



### SCPI View Model Test:
continue checking the restore.  See if the code that clears the queue fixes this issue.
then move reading the errors from the restore until after restoring the gpib lan.
then add a test where the read after write is changed and then restored.
it might be a good idea to clear the buffer on restore after connection is established and before anything else is done.
Check the restore with reading the errors with the wrong settings. 
then use the device to restore and then restore the function mode. 
* repeat the tests a few times.
* test the monitoring.
* see if we can use a loop to wait for a couple of triggers or a timeout after 5 seconds....

2700 Demo
* Use the observer to update the demo
* test using the view model.


## Beta 202308

https://stackoverflow.com/questions/18124853/excel-vba-checkbox-click-and-change-events-the-same

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
