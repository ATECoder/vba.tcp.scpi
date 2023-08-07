### About

[cc.isr.VI] is an Excel workbook for controlling and querying SCPI based instruments over TCP/IP and supporting higher level [ISR] workbooks.

Presently supported is the Keithley 2700 instrument either as an LXI instrument or a GPIB instrument by way of a GPIB-Lan controller such as the [Prologix] GPIB to LAN device.

#### Dependencies

The [cc.isr.VI] workbook depends on the following Workbooks:

* [cc.isr.Core] - Includes core Visual Basic for Applications classes and modules.
* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.
* [cc.isr.Ieee488] - Controls and queries instruments that support the IEEE 488.2 standard.

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.VI] workbook includes the following worksheets:

* Identity -- To query the instrument identity using the *IDN? command.
* K2700    -- To command and query the Keithley 2700 scanning multimeter.

## Key Features

* Provides commands and queries for communicating with IEEE488.2 instruments.
* Uses Windows Winsock32 calls to construct sockets for communicating with the instrument by way of a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Provides GPIB-Lan commands and queries for communicating with the GPIB-Lan controller.
* Provides an extended sets of methods to control the Keithley 2700 instrument.
* Provides a custom sets of methods to control the Keithley 2700 instrument for measuring 4-wire resistances from the front or read panel using internal or external triggers.

## Main Types

The main types provided by this library are:

* _K2700_ Implements some basic 2700 scanning multimeter functionality.

## Scripts

* _Deploy_: copies files to the build folder.

### Testing

Testing information is included in the [Testing] document.

## Scripts

* Build: copies files to the build folder and remove the existing references.
* Deploy: copies files to the build folder.

### Feedback

[cc.isr.vi] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.vi] repository.

[cc.isr.vi]: https://github.com/ATECoder/vba.iot.tcp/src/vi
[cc.isr.Core]: https://github.com/ATECoder/vba.iot.tcp/src/core
[cc.isr.Winsock]: https://github.com/ATECoder/vba.iot.tcp/src/Winsock
[cc.isr.Ieee488]: https://github.com/ATECoder/vba.iot.tcp/src/ieee488
[Testing]: ./cc.isr.vi.testing.md

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Prologix]: https://prologix.biz/product/gpib-ethernet-controller/
[Prologix GPIB-Lan controller]: https://prologix.biz/product/GPIB-ethernet-controller/
[ISR]: https://www.integratedscientificresources.com
