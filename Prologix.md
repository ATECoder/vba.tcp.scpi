# Using the Prologix GPIB-ETHERNET (GPIB-Lan) Controller

Prologix GPIB-ETHERNET controller converts any computer with a network port into 
a GPIB Controller or Device. 

In Controller mode, Prologix GPIB-ETHERNET controller can remotely control GPIB 
enabled instruments such as Oscilloscopes, Logic Analyzers, and Spectrum Analyzers. 

In Device mode, Prologix GPIB-ETHERNET controller converts the computer into a 
GPIB peripheral for downloading data and screen plots from the instrument front panel. 

## Required Applications

* [Netfinder]
* [Prologix GPIB Configurer]

Additional resources are located at the [Prologix] web site.

## Setup

1) Connect the PC to the local network and record it IP address;
2) Connect the Prologix to the local network or directly to the computer using a direct or crossover cable.
3) Open Netfinder;
4) Click _Search_;
5) Netfinder will locate the device and display it's default IP as 0.0.0.0;
6) Click _Asign IP_;
7) Assuming the PC IP address is 192.168.0.100, enter a static IP as follows:
	* IP Address: 192.168.0.252
	* Subnet Mask: 255.255.255.0
	* Default Gateway: 192.168.0.1
8) Open the GPIB Configurer;
9) Select the Prologix in the _Select Device_ panel;
10) Enter the instrument GPIB address, e.g., 16.
11) Enter the identity command, *IDN? to the left of the _Send_ button.
12) Click _Send_;
13) The instrument identity is displayed in the _Terminal_ panel.

## Controller Mode

The Prologix GPIB-Lan device is used in its Controller Mode.

In Controller mode, the GPIB-ETHERNET Controller acts as the Controller-In-Charge 
(CIC) on the GPIB bus. When the controller receives a command over the network port 
terminated by the network terminator – CR (ASCII 13) or LF (ASCII 10) – it addresses 
the GPIB instrument at the currently specified address (See `++addr` command) to listen, 
and passes along the received data. 

When Read-After-Write feature is enabled (See `++auto` command) the controller 
addresses the instrument to talk after sending a command, 
in order to read its response. All data received from instruments over GPIB is sent to the host over the network. Thus, the Read-After-Write feature simplifies communication with instruments. You send commands and read responses without consideration for low level GPIB protocol details. 

The Read-After-Write feature causes unterminated errors (-410) on instruments such as the Keithley 2700 Multimeter Scanner, in which case it needs to be disabled.

When Read-After-Write feature is not enabled the controller 
does not automatically address the instrument to talk. In this case, the `++read`
command is sent to the controller to read data from the instrument. 

[Prologix]: https://prologix.biz/resources/
[Prologix GPIB Configurer]: http://www.ke5fx.com/gpib/readme.htm
[Netfinder]: https://prologix.biz/downloads/netfinder.exe