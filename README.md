# Conform Universal
Conform Universal (ConformU) is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification. ConformU runs natively on Linux, Arm, MacOS and Windows and will supersede the original Windows Forms based Conform application.

# Features
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform
  * Alpaca devices that support the Alpaca discovery protocol will automatically be presented for selection.
  * An Alpaca device URL, port number and device number can be set manually to target devices that don't support discovery.
  * When running on Windows, COM Drivers will be presented for testing as well as Alpaca devices.
* The application can run:
  * in the user's default browser as a Blazor single page application
  * as a command line utility without any graphical UI
* Test outcomes are reported:
  * in the browser UI and console
  * as a human readable log file
  * as a structured machine readable JSON report file
* The Alpaca Discovery Map provides an Alpaca Device view of discovered ASCOM devices and a unique ASCOM Device view of discovered Alpaca devices.
* The application will advise when a new update is available on GitHub.

# Changes from current Conform behaviour
* An issue summary is appended at the end of the conformance report.
* All conformance deviations are reported as "issues" rather than as a mix of "errors" and "issues".
* Settings can be reset to default if required.
* Navigation is locked while Alpaca discovery and device testing is underway, but the conformance test run can still be interrupted on demand.

## Command line options
* '--commandline' - Run as a command line application with no graphical interface using a pre-configured settings file.
* '--settings fullyqualifiedfilename' - Fully qualified file name of the application configuration file. Leave blank to use the default location.
* '--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,
* '--logfile filename' - If filename has no directory/folder component it will be appended to the log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter will be ignored. Leave filename blank to use automatic file naming, where the file name will be based on the file creation time.
* '--resultsfile fullyqualifiedfilename' - Fully qualified file name of the results file.
* '--debugdiscovery' - Write discovery debug information to the log.
* '--debugstartup' - Write start-up debug information to the log.
* '--help' - Dislay a help screen.
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
* Create a folder on your device to hold the ConformU application.
* Expand the appropriate "zip" or "tar" archive that corresponds to your device's hardware / OS architecture into the newly created folder.
* Linux environments only: Give the executable run permissions with the 'chmod 755 ./conformu' command.
* The application executable (conformu (Linux) or conformu.exe (Windows))can now be run directly from a file browser or from the command line.
* When run, the executable starts a command line server application and, if operating in the default GUI mode, automatically starts the default browser and creates a new tab to hold the application window.

# Known Issues
* The server console application does not always close immediately the browser window is closed. Pressing CTRL/C in the console window will terminate the server application.
* Raspberry Pi3B - If the Chromium browser is the default browser and is not already running, the ConformU UI will not be displayed, and Chromium error messages will appear in the console. If Chromium is already running, the ConformU UI is displayed as expected. The Firefox browser displays the ConformU UI correctly regardless of whether or not it is already running. The issue with Chromium does not appear on the 64bit Pi4.
* When testing COM Video drivers, the Video.LastVideoFrame property incorrectly reports a PropertyNotImplemented error. This will be fixed in the next release.
