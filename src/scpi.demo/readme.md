# About

[cc.isr.tcp.scpi.demo] is an Excel workbook for demonstrating control and communication with some SCPI based instruments over TCP/IP.

Presently supported is the Keithley 2700 instrument either as an LXI instrument or a GPIB instrument by way of a GPIB-Lan controller such as the [Prologix] GPIB to LAN device.

## Workbook references

* [cc.isr.tcp.scpi] - Controls and queries specific virtual instruments such as the Keithley 2700.
* [cc.isr.tcp.Ieee488] - Controls and queries instruments that support the IEEE 488.2 standard.
* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Worksheets

* Identity -- To query the instrument identity using the *IDN? command.
* K2700    -- To command and query the Keithley 2700 scanning multimeter.

## Integration Testing

### Identity Worksheet Testing

Follow this procedure for reading the instrument identity string:

* Select the Identity sheet.
* Enter the instrument dotted IP address, such as `192.168.252`;
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Click _Read Identity_ to read the instrument identity using the `*IDN?` query command:
  * Check the following options:
	* ___Use VI Session___ to test the `ViSssion` class;
	* ___Use IEEE 488 Session___ to test the IEEE488 session class.

### K2700 Worksheet Testing

* The K2700 Worksheet is used to command and query the 2700 scanning multimeter using IEEE488.2 and SCPI commands and queries as well as resistance measurements controller manual and via external trigger.

#### Query Unterminated Errors and the GPIB-Lan controller

The GPIB-Lan controller _Read-After-Write_ feature addresses the instrument to talk after sending messages to the instrument.
Instruments such as the Keithley 2700 Scanning Multimeter throw Query Unterminated errors when addressed to talk when not 
having data to send. This can be addressed by turning off _Read-After-Write_ and using the controller's `++read` command for reading from the instrument. 

Turning _Read-After-Write_ on addresses the instrument to talk and, therefore, could could cause a Query Unterminated error. 

Here are some issues to keep in mind when using the IEEE488 test sheet:

* By default, the Controller is initialized with _Read-After-Write_ turned off.
	* Thus, the _Read-After-Write_ state is `False` upon connecting to the instrument.
	* Internally, the program uses the controller's `++read` command to get the readings from the instrument. 
* Toggling _Read-After-Write_ may cause Query Unterminated errors.
* Following instrument errors, commands, which check the status byte for errors, would fail to run because of the error status of the instrument.
* Issuing the `*CLS` command clears this error condition provided the command is appended with `*OPC?`, which turns the command into a query thus avoiding the Query Unterminated error on the bare `*CLS`.
* By default, as implemented by the __CLS__ and __RST__ buttons, the program appends `*OPC?` to its implementation of the `*CLS` and `*RST` commands thus keeping the program in sync with the instrument and avoiding Query Unterminated errors even if the instrument is set for _Read-After-Write_.
* When _Read-After-Write_ is turned on from the test sheet __SET__ command button:
	* The program is set to turn off _Read-After-Write_ on the next `Write` to prevent the Query Unterminated error.
	* The program then updates the state of the _Read-After-Write_ value on the sheet.
	* In other words, with this implementation, instrument communication is largely aimed at avoiding Query Unterminated by turning _Read-After-Write_ off.

#### Connecting and Disconnecting

Follow the procedures below for connecting and disconnecting the instrument:

* Enter the instrument dotted IP address such as `192.168.0.252`.
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Len controller such as the Prologix controller.
* Depress the ___Toggle Connection___ button to connect the instrument.
	* The instrument connection information such as the _Socket Address_ and _Id_ display at the top row;
	* Control buttons are enabled.
* Release the ___Toggle Connection___ button to disconnect the instrument.
	* Control buttons are disabled.

#### Errors

The last error is displayed to the right of the _Last Error_ row heading.  

Commands issued after an error will be sent to the instrument after clearing the instrument to its known state using the ___CLS___ button.

#### Testing IEEE 488.2 Commands

Follow this procedure to exercise the IEEE 488.2 command:

* Connect the instrument as described above;
* Click the ___RST___ to reset the instrument to its known state. Notice that the reset takes over a second. 
	* Some query commands take a bit longer to execute. The extended time is handled by awaiting for the result for a timeout specified by the session timeout interval, which is different from the socket receive timeout and the GPIB-Lan timeout. 
* Click the ___CLS___ button to clear the instrument to its know state clearing any existing errors;
* Select a command from the ___Command___ drop down list;
	* If a query command, ending with a _?_ is selected, click ___Write___ and then ___Read___ or ___Query___, otherwise click _Read_.
* For example, select the _*IDN?_ command and click ___Query___. The instrument identity should display under the _Received_ heading. 
* The elapsed time for each command is displayed under the _ms_ heading.
* Check the ___Read Status After Write___ check box to automatically read the instrument status byte. 
	* With Tcp control of LXI instruments, the status byte can be queried only after non-query commands. 
	* The GPIB-Lan controller is capable of reading the status byte using _Serial Poll_ even after a query write.
	* The _Read Status After Write_ uses the GPIB-Lan to query the status byte when using the GPIB-Lan controller. 
	* The serial polled value is displayed under _Spoll_ and the value read using ___*STB?___ is displayed under _SRQ_.
	* ___*ESR?___ reads the standard event status which helps determine which event turned on the Requesting Status (RQS) bit of the status register.
	
#### Testing the GPIB_Lan Controller

Follow this procedure to exercise the GPIB-Lan controller:

* Connect the instrument as described above;
* The GPIB-Lan controller buttons are enabled if connecting with the controller on port 1234.
* Once enabled, the command buttons can be used to:
	* ___GTL___: Go to Local sending the instrument to local. The instrument automatically switches to remote on the next command.
	* ___LLO___: Local Lockout to lock the _Local_ instrument button;
	* ___SDC___: Selective device clear;
	* Toggle _Read-After-Write_;
		* Note that if _Read-After-Write_ is `True`, it directs the instrument to 'talk', which automatically sets the instrument to talk after any command. With some instruments (e.g., the Keithley 2700), this causes an instrument Query Unterminated error. This error state lingers until the next `*CLS` command.
	* ___SPOLL___: Serial poll to read the status byte;
	* ___SRQ___: to tell if the Requesting Service signal (Bit 6) of the service request register of the instrument is set;
	* Get or set the _GPIB Address_;
	* Get and set the controller _Read Timeout_ for reading the instrument.
		* Note that the `Ieee488Session` class commands the controller to read the message from the instrument if auto Read-After-Write is turned of. This timeout affects such reading.

#### Testing the Resistance Measurements

The K2700 worksheet includes a set of commands and data for measuring resistances that are numbered to match the scan card channels. 

* Click _Read Cards_ to read and display the scanning cards installed in the instrument.
* Click _Set Scans_ to set the scan lists that are used when scanning all cards in sequence.
* Click _Query Inputs_ to read the status of the _Inputs_ button ont he instrument front panel.

The resistance measurements are controlled by command and option buttons as follows:

#### Single read of a specific resistance from the front or rear scanning cards.

* Toggle the Single/Scan button to Single;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to either manual or auto to select the specified resistance number and optionally automatically increment the resistance after each reading.
* Click _Read R_ to take a single reading from either the front panel or via the scanning cards.

#### Multi-resistance reading from the rear scanning cards.

* Toggle the Single/Scan button to Scan;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to auto to automatically increment the resistance after each reading.
* Click _Read R_ to take readings from all resistances via the scanning cards.

##### Multi-resistance reading from the rear scanning cards controlled by external triggering.

* Toggle the Single/Scan button to Scan;
* Toggle the Front/Rear button to either Front or Rear;
* Toggle the Manual/Auto button to auto to automatically increment the resistance after each reading.
* Depress _Ext Trig_ to start monitoring the external trigger event.
	* Externally trigger to take sequential measurements on all resistances;
	* Release the _Ext Trig_ button to end monitoring the external trigger event.

# Feedback

[cc.isr.tcp.scpi] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.tcp.scpi] repository.

[cc.isr.tcp.scpi]: https://github.com/ATECoder/vba.tcp.scpi
[cc.isr.tcp.scpi.test]: https://github.com/ATECoder/vba.tcp.iv/src/test
[cc.isr.tcp.scpi.demo]: https://github.com/ATECoder/vba.tcp.scpi/src/demo

[cc.isr.tcp.ieee488]: https://github.com/ATECoder/vba.tcp.ieee488
[cc.isr.winsock]: https://github.com/ATECoder/vba.winsock/src/
[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testfx

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>

[Prologix]: https://prologix.biz/product/gpib-ethernet-controller/
[Prologix GPIB-Lan controller]: https://prologix.biz/product/GPIB-ethernet-controller/
