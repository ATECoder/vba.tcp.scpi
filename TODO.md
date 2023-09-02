# TODO

## GIT:

Core_IO: Fix raising, enqueuing and cloning user defined errors. 
Replace the show method in the work book activation as the referenced object may not be fully referenced yet.

## TODO

IEEE488:
Consider removing the error tracers in favor for actually reporting the errors in the user interface.
See how this works in the test sheet and then in the IEEE 488 sheet.
? remove the error tracers in favor of logging or raising the error.


2700:  
update the tests per the ieee tests.

run the scpi test. 
Sense System Test
* Add and test the function write and query.
View Model test:
* Run the view model test to ensure inconclusive works.
* Test connecting.
* Test no errors upon connecting.
* Test restoring known state.
* Test reading and parsing errors.
* Test all view model commands.
* test switching modes.
* Test getting and setting sense function. 
2700 Demo
* fix and test


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
