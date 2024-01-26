# About

[cc.isr.tcp.scpi] is an Excel workbook for controlling and querying SCPI based instruments over TCP/IP.

Presently supported is the Keithley 2700 instrument either as an LXI instrument or a GPIB instrument by way of a GPIB-Lan controller such as the [Prologix] GPIB to LAN device.

## Workbook references

* [cc.isr.tcp.Ieee488] - Controls and queries instruments that support the IEEE 488.2 standard.
* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Key Features

* Provides commands and queries for communicating with IEEE488.2 instruments.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.
* Provides an extended sets of methods to control the Keithley 2700 instrument.
* Provides a custom sets of methods to control the Keithley 2700 instrument for measuring 4-wire resistances from the front or read panel using internal or external triggers.

## Main Types

The main types provided by this library are:

* _K2700_ Implements some basic 2700 scanning multimeter functionality.
* _scpi system_ Implements some basic SCPI SSystem subsystem commands.

## Unit Testing

See [cc.isr.tcp.scpi.test]

## Integration Testing

See [cc.isr.tcp.scpi.demo]

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

[unit test]: ./unit.test.lnk
[deploy]: ./deploy.ps1
[localize]: ./localize.ps1

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>

[Prologix]: https://prologix.biz/product/gpib-ethernet-controller/
[Prologix GPIB-Lan controller]: https://prologix.biz/product/GPIB-ethernet-controller/
