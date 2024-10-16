# Change log
All notable changes to these libraries will be documented in this file in a format based on [Keep a Change log]

## [1.0.8746] - 2023-12-12
* Add DUT Count to the measure mode and use it to set the DUT count.
* Add maximum DUT count to the Measure Mode class and set the View Model value upon configuration.
* Remove Reading Offset and Timer Interval from the command arguments as these are set upon configurations.
* Update selecting and parsing the selected DUT number.
* Define rear and front input sense functions.
* Select DUT upon the DUT measured event in manual single reading mode.

## [1.0.8735] - 2023-12-01
* Open the beta 202312 branch.
* Add enumerating DUT.
* Set manual mode Maximum DUT number to 48.

## [1.0.8708] - 2023-11-03
* Fix and test the demo.

## [1.0.8707] - 2023-11-02
* Complete fixing and running unit tests.
* All Tests passed.

## [1.0.8704] - 2023-10-31
* Add delays to the write and query line methods.
* Fix the query inputs method by prefixing the message with *OPC.
* Fix how measurements are propagated from the view model.
* Update the sheets and views.
* Not tested.

## [1.0.8702] - 2023-10-29
* View model: change accessibility of most let properties to Friend;
* Add arguments to view model commands.
* K2700: Set Device properties when setting the K2700 properties.
* K2700 Sheet: Initialize the socket address.
* Not tested.

## [1.0.8698] - 2023-10-25
* Fix the demo.
* Add document info.
* Remove beta branches and old LFS files.

## [1.0.8620] - 2023-08-08
* passed tests.
* merged into main.

## [1.0.8619] - 2023-08-07
* fork of [VBA IOT TCP]. 

&copy;  2023 Integrated Scientific Resources, Inc. All rights reserved.

[1.0.8746]: https://github.com/ATECoder/vba.tcp.scpi
[Keep a Change log]: https://keepachangelog.com/en/1.0.0/
[VBA IOT TCP]: https://github.com/ATECoder/vba.iot.tcp