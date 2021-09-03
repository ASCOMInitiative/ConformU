# ConformU
ConformU is a cross-platform tool to validate that Alpaca Devices and ASCOM Drivers conform to the ASCOM interface specification. 

# Features
* The application runs as a Blazor single page application in the user's default browser and can also be used directly from the command line without any UI.
* Test outcomes are reported in:
  * the UI
  * a human readable log file
  * a structured machine readable JSON report file
* Tests Alpaca devices on all Platforms and COM Drivers on the Windows platform

## Command line options
* '--commandline' - Run Conform from the command line.
* '--settings fullyqualifiedfilename' - Fully qualified file name of the configuration file. Leave blank to use the default location.
* '--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,
* '--logfilename filename' - If filename has no directory/folder component it will be appended to the log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter wil be ignored. Leave filename blank to use automatic file naming, where the file name will be based on the file creation time.
* '--debugdiscovery' - Write discovery debug information to the log.
* '--debugstartup' - Write start-up debug information to the log.
* '--help' - Display this help screen.
* '--version' - Display version information.

# Compatibility
* Windows (tested)
* Linux-64 (tested)
* Arm-32 (tested)
* Arm-64 (not tested)
* MacOS (not tested)