# Conform Universal
Conform Universal (ConformU) is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification, and that Alpaca devices conform to the Alpaca Protocol specification. ConformU runs natively on Linux, Arm, MacOS and Windows and will supersede the original Windows Forms based Conform application.

# Features
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform.
  * Alpaca devices that support the Alpaca discovery protocol will automatically be presented for selection.
  * An Alpaca device URL, port number and device number can be set manually to target devices that don't support discovery.
  * When running on Windows, COM Drivers will be presented for testing as well as Alpaca devices.
* Tests Alpaca device conformance with the Alpaca API protocol.
* Tests Alpaca device conformance with the Alpaca Discovery protocol.
* Tests Alpaca devices for conformance with the Alpaca protocol.
* The application can run:
  * in the user's default browser as a Blazor single page application.
  * as a command line utility without any graphical UI.
* Test outcomes are reported:
  * in the browser UI and console.
  * as a human readable log file
  * as a structured machine readable JSON report file.
* The Alpaca Discovery Map provides an Alpaca Device view of discovered ASCOM devices and a unique ASCOM Device view of discovered Alpaca devices.
* The application will advise when a new update is available on GitHub.
* On Windows ConformU can be configured to run as a 32bit application when installed on a 64bit OS.

# Command line
The command line format is: `conformu [COMMAND] [OPTIONS]` and the primary commands are:

| __Command__ | __Description__ |
| --- | --- |
| conformance&nbsp;`DEVICE_IDENTIFIER` | Run a full conformance check on the specified COM driver or Alpaca device.|
| alpacaprotocol `ALPACA_URI` | Run an Alpaca protocol check on the specified Alpaca device. |
| conformance-settings | Run a full conformance check on the specified COM driver or Alpaca device configured in the supplied settings file. |
| alpacaprotocol | Run an Alpaca protocol check on the device configured in the supplied settings file. |

`DEVICE_IDENTIFIER` can either be a COM ProgID such as ASCOM.Simulator.Camera or an Alpaca URI such as http://192.168.1.42:11111/api/v1/telescope/0. The application will
automatically determine the technology type and device type from the device identifier.

`ALPACA_URI` is an Alpaca URI such as http://192.168.1.42:11111/api/v1/telescope/0. The device type will be inferred from the URI.

`Options` vary depending on the command being run.

Full details of commands and options are available by entering `conformu --help` at a command prompt.

# Supported Operating Systems and Hardware
* Linux 64bit
* Arm (tested on Raspberry Pi 3B and 5)
  * 32bit OS
  * 64bit OS
* Windows
  * x64
  * x86
  * ARM
* MacOS
  * Apple silicon
  * Intel silicon

# Installation and support requirements
The applications are self-contained and do not require any additional support components to be installed on the target machine. To install ConformU:
* **Windows**
  * Download and run the Windows installer from the latest release.
  * The Windows installer supports both 32bit and 64bit OS and will create the expected start menu entries.
  * Optionally it will create a desktop icon.
* **MacOS**
  * Download the MacOS installer from the latest release.
  * The installer is a DMG file that self-mounts when double clicked and displays a basic GUI.
  * Use this GUI to drag and drop the Conform Universal application to the Applications folder.
  * The application can then be started as normal from the Applications folder.
* **Linux**
  * Download the appropriate Linux x64, ARM 32bit or ARM 64bit compressed archive for your environment from the latest release to a convenient folder.
  * Expand the archive using "tar -xf" or a similar utility.
  * The application can then be run from the command line by the command: ./conformu

# Mode of Operation
Conform Universal is a .NET single page Blazor application, which runs a web server on the localhost IP address (127.0.0.1). Each time the application starts it selects an unused IP port unless configured to always use the same port.

When run, the executable starts the command line server application and, if operating in the default GUI mode, automatically starts the default browser and creates a tab to hold the application window. If operating in command line mode, the browser is not started.