# TODO

## GIT:

Core:
Add parsing methods and unit tests. 

IEEE488
Add vi session receive string to be used for clearing the instrument queue against query interrupted errors.
Add Restore Known State to the device.
Use string extension parsing to parse data from the instrument.
Fix restoring GPIB Lan state.

SCPI:
Fix document typos.
Use string extension parsing when reading device values.
Use the IEEE 488 device for determining the restore of the gpib lan device.



## TODO

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
