# ConformU
ConformU is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification. It supersedes the original Windows based Conform application.

# Features
* The application runs as a Blazor single page application in the user's default browser and can also be used directly from the command line without any UI.
* Test outcomes are reported:
  * in the browser UI and console
  * as a human readable log file
  * as a structured machine readable JSON report file
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform
  * Alpaca devices that support the Alpaca discovery protocol will automatically be presented for selection.
  * An Alpaca device URL, port number and device number can be set manually to target devices that don't support discovery.
  * COM Drivers installed on the host WIndows PC will automatically be presented for selection
* The graphical UI adjusts responsively to window size.
* Navigation is locked while Alpaca discovery and device testing is underway.
* A conformance test run can be interrupted on demand.
* Settings can be reset to default on demand.

## Command line options
* '--commandline' - Run as a command line application with no graphical interface using a pre-configured settings file.
* '--settings fullyqualifiedfilename' - Fully qualified file name of the application configuration file. Leave blank to use the default location.
* '--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,
* '--logfilename filename' - If filename has no directory/folder component it will be appended to the log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter wil be ignored. Leave filename blank to use automatic file naming, where the file name will be based on the file creation time.
* '--debugdiscovery' - Write discovery debug information to the log.
* '--debugstartup' - Write start-up debug information to the log.
* '--help' - Display this help screen.
* '--version' - Display version information.

# Platform availability
* Windows (tested)
* Linux-64 (tested)
* Arm-32 (tested)
* Arm-64 (not tested)
* MacOS (not tested)

# Installation and support requirements
The application is self-contained and does not require that any additional support components be installed on the target machine. To install ConformU:
* Unzip the relevant OS archive to a folder of your choice on the target machine.
* In Linux environments you will need to give the executable run permissions with the 'chmod 755 ./conformu' command.
* The application can now be run by directly from a file browser or from the command line.