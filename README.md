# Conform Universal
Conform Universal (ConformU) is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification, and that Alpaca devices conform to the Alpaca Protocol specification. ConformU runs natively on Linux, Arm, MacOS and Windows and will supersede the original Windows Forms based Conform application.

# Features
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform
  * Alpaca devices that support the Alpaca discovery protocol will automatically be presented for selection.
  * An Alpaca device URL, port number and device number can be set manually to target devices that don't support discovery.
  * When running on Windows, COM Drivers will be presented for testing as well as Alpaca devices.
* Tests Alpaca device conformance with the Alpaca API protocol
* Tests Alpaca device conformance with the Alpaca Discovery protocol
* The application can run:
  * in the user's default browser as a Blazor single page application
  * as a command line utility without any graphical UI
* Test outcomes are reported:
  * in the browser UI and console
  * as a human readable log file
  * as a structured machine readable JSON report file
* The Alpaca Discovery Map provides an Alpaca Device view of discovered ASCOM devices and a unique ASCOM Device view of discovered Alpaca devices.
* The application will advise when a new update is available on GitHub.
* On Windows ConformU can be configured to run as a 32bit application when installed on a 64bit OS.

# Changes from current Conform behaviour
* An issue summary is appended at the end of the conformance report.
* All conformance deviations are reported as "issues" rather than as a mix of "errors" and "issues".
* Settings can be reset to default if required.
* Navigation is locked while Alpaca discovery and device testing is underway, but the conformance test run can still be interrupted on demand.

# Command line options
* '--commandline' - Run as a command line application with no graphical interface using a pre-configured settings file.
* '--settings fullyqualifiedfilename' - Fully qualified file name of the application configuration file. Leave blank to use the default location.
* '--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,
* '--logfile filename' - If filename has no directory/folder component it will be appended to the log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter will be ignored. Leave filename blank to use automatic file naming, where the file name will be based on the file creation time.
* '--resultsfile fullyqualifiedfilename' - Fully qualified file name of the results file.
* '--debugdiscovery' - Write discovery debug information to the log.
* '--debugstartup' - Write start-up debug information to the log.
* '--help' - Display a help screen.
* '--version' - Display version information.

# Supported Operating Systems and Hardware
* Linux 64bit
* Arm (tested on Raspberry Pi 3B and 4)
  * 32bit OS
  * 64bit OS
* Windows
  * x64
  * x86
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