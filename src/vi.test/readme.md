# About

[cc.isr.tcp.vi.test] is an Excel workbook for testing the [cc.isr.tcp.vi] workbook.

Presently supported is the Keithley 2700 instrument either as an LXI instrument or a GPIB instrument by way of a GPIB-Lan controller such as the [Prologix] GPIB to LAN device.

## Workbook references

* [cc.isr.tcp.vi] - Controls and queries specific virtual instruments such as the Keithley 2700.
* [cc.isr.tcp.Ieee488] - Controls and queries instruments that support the IEEE 488.2 standard.
* [cc.isr.Winsock] - Implements TCP Client and Server classes with Windows Winsock API.
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.
* [cc.isr.test.fx] - Test framework workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Worksheets

* UnitTestSheet -- To run unit tests (pending).

## Scripts

* [unit test]: shortcut to run unit tests.
* [deploy]: copy the workbook and its dependencies to the deployment `..\..\bin` folder.
* [localize]: sets workbook references to the local deployment `..\..\bin` folder.

## Unit Testing

At this time, the [cc.isr.tcp.ieee488] workbooks exclusively employs integration testing using the IEEE488 and Identity worksheets. 

Units testing will be added in future releases.

# Feedback

[cc.isr.tcp.vi] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.tcp.vi] repository.

[cc.isr.tcp.vi]: https://github.com/ATECoder/vba.tcp.vi
[cc.isr.tcp.vi.test]: https://github.com/ATECoder/vba.tcp.iv/src/test
[cc.isr.tcp.vi.demo]: https://github.com/ATECoder/vba.tcp.vi/src/demo

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
