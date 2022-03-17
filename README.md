# ConformU - Beta 0.0.1
ConformU is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification. ConformU runs natively on Linux, Arm, MacOS and Windows and, when out of beta, will supersede the original Windows Forms based Conform application.

# Features
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform
  * Alpaca devices that support the Alpaca discovery protocol will automatically be presented for selection.
  * An Alpaca device URL, port number and device number can be set manually to target devices that don't support discovery.
  * When running on Windows, COM Drivers will be presented for testing as well as Alpaca devices.
* The application can run
  * in the user's default browser as a Blazor single page application
  * as a command line utility without any graphical UI
* Test outcomes are reported:
  * in the browser UI and console
  * as a human readable log file
  * as a structured machine readable JSON report file

# Changes from Windows Conform behaviour
* An issue summary is appended at the end of the conformance report.
* All conformance deviations are reported as "issues" rather than as a mix of "errors" and "issues".
* Includes an Alpaca Discovery Map feature that shows both an Alpaca Device view of discovered ASCOM devices and a unique ASCOM Device view of discovered Alpaca devices.
* Settings can be reset to default if required.
* Navigation is locked while Alpaca discovery and device testing is underway, but the conformance test run can be interrupted on demand.

## Command line options
* '--commandline' - Run as a command line application with no graphical interface using a pre-configured settings file.
* '--settings fullyqualifiedfilename' - Fully qualified file name of the application configuration file. Leave blank to use the default location.
* '--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,
* '--logfile filename' - If filename has no directory/folder component it will be appended to the log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter will be ignored. Leave filename blank to use automatic file naming, where the file name will be based on the file creation time.
* '--resultsfile fullyqualifiedfilename' - Fully qualified file name of the results file.
* '--debugdiscovery' - Write discovery debug information to the log.
* '--debugstartup' - Write start-up debug information to the log.
* '--help' - Display a help screen.
* '--version' - Display version information.

# Hardware and OS Platform availability
* Windows (tested)
* Linux-64 (tested)
* Arm-32 (tested)
* Arm-64 (tested)
* MacOS (not tested)

# Installation and support requirements
The application is self-contained and does not require that any additional support components be installed on the target machine. To install ConformU:
* Create a folder on your device to hold the ConformU application.
* Expand the appropriate "zip" or "tar" archive that corresponds to your device's hardware / OS architecture into the newly created folder.
* Linux environments only: Give the executable run permissions with the 'chmod 755 ./conformu' command.
* The application can now be run directly from a file browser or from the command line.

# Known Issues
* The server console application does not always close immediately the browser window is closed. Pressing CTRL/C in the console window will terminate the server application.
* Raspberry Pi3B - If the Chromium browser is the default browser and is not already running, the ConformU UI will not be displayed, and Chromium error messages will in the console. If Chromium is already running, the ConformU UI is displayed as expected. The Firefox browser displays the ConformU UI correctly regardless of whether or not it is already running. The issue with Chromium does not appear on the 64bit Pi4.