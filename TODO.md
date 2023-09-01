# TODO

## GIT:

Core_IO:
Add Ieee488 Device Error

ieee 488:
Test: Clear error state before each test.
GPIB Lan: 
add error handling to the connection changed event handler;
Add Assert talk after write required condition.
Remove test of current state from the read after write setter.
Rename setter property value to a_value by value.

Scpi:
K2700: Set the Gpib Lan Assert Talk option to false. 
K2700 and View Model: Remove setting the read after write.
View Model:
Raise Device error when detecting those so that gets reported as part of the last error.
Test: add detection of connection failure as in the IEEE488 tests;
Clear error state before each test.




## TODO

2700:  
run the scpi test. 
Add the sense test and test function query.
Run the view model test to ensure inconclusive works.
Test connecting.
Test no errors upon connecting.
Test restoring known state.
Test reading and parsing errors.
Test all view model commands.
test switching modes.
Test getting and setting sense function. 

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
