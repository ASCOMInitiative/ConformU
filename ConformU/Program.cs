// Ignore Spelling: com prog SHOWMINIMIZED

using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.CommandLine.Builder;
using System.CommandLine.Parsing;
using System.CommandLine;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static ConformU.Globals;
using System.Text.RegularExpressions;
using ASCOM.Common;
using ASCOM.Common.Alpaca;

namespace ConformU
{
    public partial class Program
    {
        private const string RED_TEXT = "\u001b[91m";
        private const string WHITE_TEXT = "\u001b[0m";
        private static string[] commandLineArguments;
        private static List<string> argList;
        private static ConformLogger conformLogger;
        private static SessionState conformStateManager;
        private static ConformConfiguration conformConfiguration;

        #region Windows DLL imports

#if WINDOWS
        [LibraryImport("kernel32.dll")]
        internal static partial IntPtr GetConsoleWindow();

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_SHOWMINIMIZED = 2;
#endif

        #endregion

        /// <summary>
        /// Entry point for the application
        /// </summary>
        /// <param name="args">COmmand line arguments supplied by the user</param>
        /// <returns>Integer 0 for success, negative values for errors and positive values indicating the number of Conformance or Alpaca Protocol check errors and issues</returns>
        public static async Task<int> Main(string[] args)
        {
            try
            {
                // Save the command line arguments so they can be reused if a 32bit application is required
                commandLineArguments = args;

                #region Command definitions

                // ROOT command
                var rootCommand = new RootCommand(
                    $"Conform Universal {Update.ConformuVersionDisplayString}\r\nCopyright (c) 2021-{DateTime.Now.Year} Peter Simpson\r\n\r\nEnter conformu [command] -h for information on options available in each command.\r\n\r\nIf no command or options are provided Conform Universal will start as a GUI application using default parameters."
                );

                // CONFORMANCE command
                var conformanceCommand = new Command("conformance", "Check the specified device for ASCOM device interface conformance with all tests enabled.");

                // ALPACA PROTOCOL command
                var alpacaProtocolCommand = new Command("alpacaprotocol", "Check the specified Alpaca device for Alpaca protocol conformance");

                // CONFORMANCE USING SETTINGS command
                var conformanceUsingSettingsCommand = new Command("conformance-settings", "Check the device configured in the settings file for ASCOM device interface conformance");

                // ALPACA PROTOCOL USING SETTINGS command
                var alpacaUsingSettingsCommand = new Command("alpacaprotocol-settings", "Check the device configured in the settings file for Alpaca protocol conformance");

                // START AS GUI command
                var startAsGuiCommand = new Command("gui", "Start Conform Universal as a GUI application with options to change the log file, results file and settings file locations");

                #endregion

                #region Argument definitions

                // COM_PROGID or ALPACA_URI argument 
                Argument<string> deviceArgument = new(
                    name: "COM_PROGID_or_ALPACA_URI",
                    description: "The device's COM ProgID (e.g. ASCOM.Simulator.Telescope) or its Alpaca URI (e.g. http://[host]:[port]/api/v1/[DeviceType]/[DeviceNumber]).\r\n" +
                    "The device technology (COM or Alpaca) and the ASCOM device type will be inferred from the supplied ProgID or URI.");

                // Add a validator to handle both COM ProgIds and Alpaca URIs
                deviceArgument.AddValidator((commandResult) =>
                {
                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(deviceArgument), out _, out _, out _, out _, out _, out _);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            commandResult.ErrorMessage = $"The COM ProgID or Alpaca URI argument is invalid: {commandResult.GetValueForArgument(deviceArgument)}";
                            break;

                        case DeviceTechnology.Alpaca:
                            // No action required because Alpaca technology is supported
                            break;

                        case DeviceTechnology.COM:
                            // No action required because COM technology is supported
                            break;

                        default:
                            commandResult.ErrorMessage = $"***** Conform Universal internal error - Unexpected technology type returned: {technology}.";
                            break;
                    }

                });

                // ALPACA URI argument
                Argument<string> alpacaDeviceArgument = new(
                    name: "Alpaca_URI",
                    description: "The device's Alpaca URI (e.g. http://host:port/api/v1/[DeviceType]/[DeviceNumber]).\r\n" +
                    "The ASCOM device type will be inferred from the supplied URI.");

                //Add a validator to handle Alpaca URIs
                alpacaDeviceArgument.AddValidator((commandResult) =>
                {
                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(alpacaDeviceArgument), out _, out _, out _, out _, out _, out _);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            commandResult.ErrorMessage = $"The Alpaca URI argument is invalid: {commandResult.GetValueForArgument(deviceArgument)}";
                            break;

                        case DeviceTechnology.Alpaca:
                            // No action required because Alpaca technology is supported
                            break;

                        case DeviceTechnology.COM:
                            commandResult.ErrorMessage = $"COM ProgIDs are not supported by this command: {commandResult.GetValueForArgument(alpacaDeviceArgument)}";
                            break;

                        default:
                            commandResult.ErrorMessage = $"***** Conform Universal internal error - Unexpected technology type returned: {technology}.";
                            break;
                    }
                });

                #endregion

                #region Option definitions

                // VERSION option
                Option<bool> versionOption = new(
                    aliases: new string[] { "-v", "--version" },
                    description: "Show the Conform Universal version number");

                // TESTCYCLES option
                Option<int> testCyclesOption = new(
                    aliases: new string[] { "-c", "--testcycles" },
                    description: "The number of test cycles to undertake");
                testCyclesOption.SetDefaultValue(1);

                // Add a validator for the number of cycles
                testCyclesOption.AddValidator(result =>
                {
                    // Get the log file path
                    int testCycles = (int)result.GetValueOrDefault();

                    if (testCycles < 1) // Found an invalid number of cycles (must be 1 or greater)
                    {
                        // Set the error message
                        Console.WriteLine($"\r\n{RED_TEXT}The number of test cycles must be 1 or more: {testCycles} {WHITE_TEXT}");
                        result.ErrorMessage = $"\r\nThe number of test cycles must be 1 or more: {testCycles}";
                    }
                });

                // DEBUG DISCOVERY option
                Option<bool> debugDiscoveryOption = new(
                    aliases: new string[] { "-d", "--debugdiscovery" },
                    description: "Write discovery debug information to the log.")
                {
                    IsHidden = true
                };

                // DEBUG STARTUP option
                Option<bool> debugStartUpOption = new(
                    aliases: new string[] { "-t", "--debugstartup" },
                    description: "Write start-up debug information to the log.")
                {
                    IsHidden = true
                };

                // LOG FILE PATH option 
                Option<DirectoryInfo> logFilePathOption = new(
                    aliases: new string[] { "-p", "--logfilepath" },
                    description: "Relative or fully qualified path to the log file folder.\r\n" +
                    "Overrides the default GUI log file path, but is ignored when the --logfilename option is present.")
                {
                    ArgumentHelpName = "PATH"
                };

                // Add a validator for the log file path
                logFilePathOption.AddValidator(result =>
                {
                    // Validate Windows paths
                    if (OperatingSystem.IsWindows())
                    {
                        // Get the log file path
                        DirectoryInfo logFilePath = (DirectoryInfo)result.GetValueOrDefault();

                        // Check each character to see if it matches an invalid character
                        foreach (char invalidCharacter in Path.GetInvalidPathChars())
                        {
                            if (logFilePath.FullName.Contains(invalidCharacter)) // Found an invalid character detected
                            {
                                // Set the error message
                                Console.WriteLine($"\r\n{RED_TEXT}Found invalid log file path character: '{invalidCharacter}' ({(int)invalidCharacter:X2}){WHITE_TEXT}");
                                result.ErrorMessage = $"\r\nLog file path contains invalid characters: {logFilePath.FullName}";
                            }
                        }
                    }
                });

                // LOG FILE NAME option
                Option<FileInfo> logFileNameOption = new(
                    aliases: new string[] { "-n", "--logfilename" },
                    description: "Relative or fully qualified name of the log file.\r\n" +
                    "The default GUI log file name will be used if this option is omitted.")
                {
                    ArgumentHelpName = "FILENAME"
                };

                // Add a validator for the log file name
                logFileNameOption.AddValidator(result =>
                {
                    // Validate Windows file names
                    if (OperatingSystem.IsWindows())
                    {
                        // Get the log file name as a FileInfo
                        FileInfo logFileInfo = (FileInfo)result.GetValueOrDefault();

                        // Check each character to see if it matches an invalid character
                        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
                        {
                            //Console.WriteLine($"Invalid log file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2})");
                            // Ignore colon and backslash, which are marked as invalid in a file name
                            if ((invalidCharacter != '\\') & (invalidCharacter != ':'))
                            {
                                if (logFileInfo.FullName.Contains(invalidCharacter)) // Found an invalid character detected
                                {
                                    // Set the error message
                                    Console.WriteLine($"\r\n{RED_TEXT}Found invalid log file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2}){WHITE_TEXT}");
                                    result.ErrorMessage = $"\r\nLog file name contains invalid characters: {logFileInfo.FullName}";
                                }
                            }
                        }
                    }
                });

                // RESULTS FILE option
                Option<FileInfo> resultsFileOption = new(
                    aliases: new string[] { "-r", "--resultsfile" },
                    description: "Relative or fully qualified name of the machine readable results file.\r\n" +
                    "The default GUI filename and location will be used if this option is omitted.")
                {
                    ArgumentHelpName = "FILENAME"
                };

                // Add a validator for the results file name
                resultsFileOption.AddValidator(result =>
                {
                    // Validate Windows file names
                    if (OperatingSystem.IsWindows())
                    {
                        // Get the log file name as a FileInfo
                        FileInfo resultsFileInfo = (FileInfo)result.GetValueOrDefault();

                        // Check each character to see if it matches an invalid character
                        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
                        {
                            //Console.WriteLine($"Invalid log file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2})");
                            // Ignore colon and backslash, which are marked as invalid in a file name
                            if ((invalidCharacter != '\\') & (invalidCharacter != ':'))
                            {
                                if (resultsFileInfo.FullName.Contains(invalidCharacter)) // Found an invalid character detected
                                {
                                    // Set the error message
                                    Console.WriteLine($"\r\n{RED_TEXT}Found invalid results file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2}){WHITE_TEXT}");
                                    result.ErrorMessage = $"\r\nResults file name contains invalid characters: {resultsFileInfo.FullName}";
                                }
                            }
                        }
                    }
                });

                // SETTINGS FILE option
                Option<FileInfo> settingsFileOption = new(
                    aliases: new string[] { "-s", "--settingsfile" },
                    description: "Relative or fully qualified name of the settings file.\r\n" +
                    "The default GUI application settings file will be used if this option is omitted.")
                {
                    ArgumentHelpName = "FILENAME"
                };

                // Add a validator for the settings file name
                settingsFileOption.AddValidator(result =>
                {
                    // Get the log file name as a FileInfo
                    FileInfo settingsFileInfo = (FileInfo)result.GetValueOrDefault();

                    // Validate Windows file names
                    if (OperatingSystem.IsWindows())
                    {
                        bool fileNameOk = true;

                        // Check each character to see if it matches an invalid character
                        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
                        {
                            //Console.WriteLine($"Invalid log file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2})");
                            // Ignore colon and backslash, which are marked as invalid in a file name
                            if ((invalidCharacter != '\\') & (invalidCharacter != ':'))
                            {
                                if (settingsFileInfo.FullName.Contains(invalidCharacter)) // Found an invalid character detected
                                {
                                    // Set the error message
                                    Console.WriteLine($"\r\n{RED_TEXT}Found invalid settings file name character: '{invalidCharacter}' ({(int)invalidCharacter:X2}){WHITE_TEXT}");
                                    result.ErrorMessage = $"\r\nSettings file name contains invalid characters: {settingsFileInfo.FullName}";
                                    fileNameOk = false;
                                }
                            }
                        }
                        if (!fileNameOk)
                            return;
                    }

                    // Validate that the file exists
                    if (!File.Exists(settingsFileInfo.FullName))
                    {
                        //Console.WriteLine($"\r\n{RED_TEXT}Settings file '{settingsFileInfo.FullName}' does not exist.{WHITE_TEXT}");
                        result.ErrorMessage = $"\r\nSettings file '{settingsFileInfo.FullName}' does not exist.";
                    }
                });

                #endregion

                #region Associate arguments and options with commands

                // ROOT COMMAND - add commands and options 
                rootCommand.AddCommand(conformanceCommand);
                rootCommand.AddCommand(alpacaProtocolCommand);
                rootCommand.AddCommand(conformanceUsingSettingsCommand);
                rootCommand.AddCommand(alpacaUsingSettingsCommand);
                rootCommand.AddCommand(startAsGuiCommand);
                rootCommand.AddOption(versionOption);

                // CONFORMANCE COMMAND - add arguments and options 
                conformanceCommand.AddArgument(deviceArgument);
                conformanceCommand.AddOption(logFileNameOption);
                conformanceCommand.AddOption(logFilePathOption);
                conformanceCommand.AddOption(resultsFileOption);
                conformanceCommand.AddOption(settingsFileOption);
                conformanceCommand.AddOption(debugDiscoveryOption);
                conformanceCommand.AddOption(debugStartUpOption);
                conformanceCommand.AddOption(testCyclesOption);

                // ALPCA COMMAND - add arguments and options
                alpacaProtocolCommand.AddArgument(alpacaDeviceArgument);
                alpacaProtocolCommand.AddOption(logFileNameOption);
                alpacaProtocolCommand.AddOption(logFilePathOption);
                alpacaProtocolCommand.AddOption(resultsFileOption);
                alpacaProtocolCommand.AddOption(settingsFileOption);
                alpacaProtocolCommand.AddOption(debugDiscoveryOption);
                alpacaProtocolCommand.AddOption(debugStartUpOption);

                // ALPCA USING SETTINGS COMMAND - add options
                alpacaUsingSettingsCommand.AddOption(logFileNameOption);
                alpacaUsingSettingsCommand.AddOption(logFilePathOption);
                alpacaUsingSettingsCommand.AddOption(resultsFileOption);
                alpacaUsingSettingsCommand.AddOption(settingsFileOption);
                alpacaUsingSettingsCommand.AddOption(debugDiscoveryOption);
                alpacaUsingSettingsCommand.AddOption(debugStartUpOption);

                // CONFORMANCE USING SETTINGS COMMAND - add options
                conformanceUsingSettingsCommand.AddOption(logFileNameOption);
                conformanceUsingSettingsCommand.AddOption(logFilePathOption);
                conformanceUsingSettingsCommand.AddOption(resultsFileOption);
                conformanceUsingSettingsCommand.AddOption(settingsFileOption);
                conformanceUsingSettingsCommand.AddOption(debugDiscoveryOption);
                conformanceUsingSettingsCommand.AddOption(debugStartUpOption);
                conformanceUsingSettingsCommand.AddOption(testCyclesOption);

                // START AS GUI COMMAND  - add options
                startAsGuiCommand.AddOption(logFileNameOption);
                startAsGuiCommand.AddOption(logFilePathOption);
                startAsGuiCommand.AddOption(resultsFileOption);
                startAsGuiCommand.AddOption(settingsFileOption);
                startAsGuiCommand.AddOption(debugDiscoveryOption);
                startAsGuiCommand.AddOption(debugStartUpOption);

                #endregion

                #region Command handlers

                // ROOT COPMMAND handler - Starts the GUI or displays version information
                rootCommand.SetHandler((bool version) => Task.FromResult(RootCommandHandler(version)), versionOption);

                // CONFORMANCE USING SETTINGS COPMMAND handler
                conformanceUsingSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp, cycles) =>
                {
                    int returnCode = -8888;

                    // Initialise required variables required by several commands
                    if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                        return Task.CompletedTask;

                    // Run the conformance test
                    returnCode = RunConformanceTest(cycles);

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption, testCyclesOption);

                // CONFORMANCE COMMAND handler
                conformanceCommand.SetHandler((device, file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp, cycles) =>
                {
                    int returnCode = 0;

                    DeviceTechnology technology = GetDeviceTechnology(device, out ServiceType? serviceType, out string address, out int port, out int alpacaInterfaceVersion, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"\r\n{RED_TEXT}***** SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD *****\r\nArgument: {device}{WHITE_TEXT}");
                            break;

                        case DeviceTechnology.Alpaca:
                            // Initialise required variables required by several commands
                            if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                                return Task.CompletedTask;

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, alpacaInterfaceVersion, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunConformanceTest(cycles);
                            break;

                        case DeviceTechnology.COM:
                            if (OperatingSystem.IsWindows()) // OK because COM is valid on WindowsOS
                            {
                                // Initialise required variables required by several commands
                                if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                                    return Task.CompletedTask;

                                // Set options to conduct a full conformance test
                                conformConfiguration.SetFullTest();

                                // Set the COM device parameters
                                conformConfiguration.SetComDevice(device, deviceType.Value);

                                // Run the conformance test
                                returnCode = RunConformanceTest(cycles);
                            }
                            else // Not valid because COM is only valid on Windows OS
                            {
                                Console.WriteLine($"\r\n{RED_TEXT}Checking COM devices on this operating system is not supported.{WHITE_TEXT}");
                            }
                            break;

                        default:
                            Console.WriteLine($"\r\n{RED_TEXT}SHOULD NEVER SEE THIS MESSAGE BECAUSE ONLY COM AND ALPACA TECHNOLOGY TYPES ARE SUPPORTED *****\r\nArgument: {device}{WHITE_TEXT}");
                            break;
                    }

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, deviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption, testCyclesOption);

                // ALPACA PROTOCOL COMMAND handler
                alpacaProtocolCommand.SetHandler((alpacaDevice, file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    int returnCode = 0;

                    DeviceTechnology technology = GetDeviceTechnology(alpacaDevice, out ServiceType? serviceType, out string address, out int port, out int alpacainterfaceVersion, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"\r\n{RED_TEXT}SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD{WHITE_TEXT}\r\nArgument: {alpacaDevice}");
                            break;

                        case DeviceTechnology.Alpaca:
                            // Initialise required variables required by several commands
                            if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                                return Task.CompletedTask;

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, alpacainterfaceVersion, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunAlpacaProtocolTest();
                            break;

                        case DeviceTechnology.COM:
                            Console.WriteLine($"\r\n{RED_TEXT}This command can only be used with Alpaca devices. Argument: {alpacaDevice}{WHITE_TEXT}");
                            break;

                        default:
                            Console.WriteLine($"\r\n{RED_TEXT}SHOULD NEVER SEE THIS MESSAGE BECAUSE ONLY COM AND ALPACA TECHNOLOGY TYPES ARE SUPPORTED\r\nArgument: {alpacaDevice}{WHITE_TEXT}");
                            break;
                    }

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, alpacaDeviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // ALPACA USING SETTINGS COMMAND handler
                alpacaUsingSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                        return Task.CompletedTask;

                    // Return a task with the required return code
                    return Task.FromResult(RunAlpacaProtocolTest());

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // START AS GUI COMMAND handler
                startAsGuiCommand.SetHandler((file, path, results, settings, debugDiscovery, debugStartUp) => Task.FromResult(StartGuiHandler(file, path, debugStartUp, debugDiscovery, results, settings)), logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                #endregion

                // Create a custom parser that does not present the default version description
                Parser parser = new CommandLineBuilder(rootCommand)
                       //.UseDefaults()
                       //.UseVersionOption()
                       .UseHelp()
                       .UseEnvironmentVariableDirective()
                       .UseParseDirective()
                       .UseSuggestDirective()
                       .RegisterWithDotnetSuggest()
                       .UseTypoCorrections()
                       .UseParseErrorReporting()
                       .UseExceptionHandler()
                       .CancelOnProcessTermination()
                       .Build();

                // Parse the command line and invoke actions as determined by supplied parameters
                return await parser.InvokeAsync(args);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"\r\n{RED_TEXT}Conform internal error - Exception reading command line or stating application:{WHITE_TEXT}\r\n{ex}");
                return ex.HResult;
            }
        }

        /// <summary>
        /// Start ConformU as a GUI application
        /// </summary>
        /// <param name="version"></param>
        /// <param name="file"></param>
        /// <param name="path"></param>
        /// <param name="debugStartup"></param>
        /// <param name="debugDiscovery"></param>
        /// <param name="resultsFile"></param>
        /// <param name="settingsFile"></param>
        private static int RootCommandHandler(bool version)
        {
            if (version)
            {
                Console.WriteLine($"\r\nConform Universal {Update.ConformuVersionDisplayString}");
            }
            else
            {
                // Initialise required variables required by several commands
                if (InitialiseVariables(null, null, false, false, null, null))
                    return 0;

                // Start a task to check whether any updates are available, if configured to do so.
                // The update check is started here to give the maximum time to get a result before the UI is first displayed
                if (conformConfiguration.Settings.UpdateCheck)
                {
                    Task.Run(async () =>
                    {
                        try
                        {
                            await Update.CheckForUpdates(conformLogger);
                        }
                        catch { } // Ignore exceptions here
                    });
                }

                // In production set the working directory to the application directory
                if (!Debugger.IsAttached)
                {
                    Directory.SetCurrentDirectory(Path.GetDirectoryName(Environment.ProcessPath));
                }

                try
                {

#if WINDOWS
                    // Minimise the console window
                    ShowWindow(GetConsoleWindow(), SW_SHOWMINIMIZED);
#endif

                    Console.WriteLine($"\r\nStarting web server.");
                    Task t = CreateHostBuilder(conformLogger, conformStateManager, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept the / prefix
                         .Build()
                         .RunAsync();

                    t.Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\r\n{RED_TEXT}Exception stating application: {ex.Message}{WHITE_TEXT}");
                    return ex.HResult;
                }
            }
            return 0;
        }

        /// <summary>
        /// Start ConformU as a GUI application
        /// </summary>
        /// <param name="version"></param>
        /// <param name="file"></param>
        /// <param name="path"></param>
        /// <param name="debugStartUp"></param>
        /// <param name="debugDiscovery"></param>
        /// <param name="resultsFile"></param>
        /// <param name="settingsFile"></param>
        private static int StartGuiHandler(FileInfo file, DirectoryInfo path, bool debugStartUp, bool debugDiscovery, FileInfo resultsFile, FileInfo settingsFile)
        {
            // Initialise required variables required by several commands
            if (InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile))
                return 0;

            // Start a task to check whether any updates are available, if configured to do so.
            // The update check is started here to give the maximum time to get a result before the UI is first displayed
            if (conformConfiguration.Settings.UpdateCheck)
            {
                Task.Run(async () =>
                {
                    try
                    {
                        await Update.CheckForUpdates(conformLogger);
                    }
                    catch { } // Ignore exceptions here
                });
            }

            // In production set the working directory to the application directory
            if (!Debugger.IsAttached)
            {
                Directory.SetCurrentDirectory(Path.GetDirectoryName(Environment.ProcessPath));
            }

            try
            {

#if WINDOWS
                // Minimise the console window
                ShowWindow(GetConsoleWindow(), SW_SHOWMINIMIZED);
#endif

                Console.WriteLine($"\r\nStarting web server.");
                Task t = CreateHostBuilder(conformLogger, conformStateManager, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept the / prefix
                     .Build()
                     .RunAsync();

                t.Wait();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\r\n{RED_TEXT}Exception stating application: {ex.Message} {WHITE_TEXT}");
                return ex.HResult;
            }
            return 0;
        }

        public static IHostBuilder CreateHostBuilder(ConformLogger conformLogger, SessionState conformStateManager, ConformConfiguration conformConfiguration, string[] args)
        {
            IHostBuilder builder = null;

            builder = Host.CreateDefaultBuilder(args)

                 .ConfigureLogging(logging =>
                 {
                     logging.ClearProviders();
                     logging.AddSimpleConsole(options =>
                     {
                         options.SingleLine = true;
                         options.IncludeScopes = false;
                         options.TimestampFormat = "HH:mm:ss.fff - ";
                     });
                     logging.AddFilter("Microsoft.AspNetCore.SignalR", LogLevel.Information);
                     logging.AddFilter("Microsoft.AspNetCore.Http.Connections", LogLevel.Information);
                     logging.AddDebug();
                 })

                 .ConfigureServices(servicesCollection =>
                 {
                     // Add the logger component to the list of injectable services
                     servicesCollection.AddSingleton(conformLogger);

                     // Add the state management component to the list of injectable services
                     servicesCollection.AddSingleton(conformStateManager);

                     // Add the configuration component to the list of injectable services
                     servicesCollection.AddSingleton(conformConfiguration);
                 })

                 .ConfigureWebHostDefaults(webBuilder =>
                    {
                        webBuilder.UseStartup<Startup>();

#if RELEASE
#if BUNDLED
                       .UseContentRoot(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location))
#endif
#endif
                        ;

                        // Start Kestrel on localhost:ConfiguredIpPort if not running under Visual Studio
                        webBuilder.UseKestrel(opts =>
                        {
                            if (!Debugger.IsAttached)
                            {
                                opts.Listen(IPAddress.Loopback, conformConfiguration.Settings.ApplicationPort);
                            }
                        });

                    });
            return builder;
        }

        #region Support code

        private static int RunConformanceTest(int numberOfCycles)
        {
            int returnCode;

            // Validate the supplied configuration and only start if there are no settings issues
            string validationMessage = conformConfiguration.Validate();
            if (!string.IsNullOrEmpty(validationMessage)) // There is a configuration issue so present an error message
            {
                Console.WriteLine($"\r\n{RED_TEXT}Cannot start test:{WHITE_TEXT}\r\n{validationMessage}");
                return -99;
            }

            // Setting have validated OK so start the test
            CancellationTokenSource cancellationTokenSource = new(); // Create a task cancellation token source and cancellation token
            CancellationToken cancelConformToken = cancellationTokenSource.Token;

            // Create a test manager instance to oversee the test
            using (ConformanceTestManager tester = new(conformConfiguration, conformLogger, cancellationTokenSource, cancelConformToken))
            {
                try
                {
                    returnCode = tester.TestDevice(numberOfCycles);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\r\n{RED_TEXT}RunConformanceTest - Exception:{WHITE_TEXT}\r\n {ex}");
                    returnCode = ex.HResult;
                }
            }

            GC.Collect();

            return returnCode;
        }

        private static int RunAlpacaProtocolTest()
        {
            int returnCode;

            // Validate the supplied configuration and only start if there are no settings issues
            string validationMessage = conformConfiguration.Validate();
            if (!string.IsNullOrEmpty(validationMessage)) // There is a configuration issue so present an error message
            {
                Console.WriteLine($"\r\n{RED_TEXT}Cannot start test:{WHITE_TEXT}\r\n{validationMessage}");
                return -99;
            }

            // Setting have validated OK so start the test
            CancellationTokenSource cancellationTokenSource = new(); // Create a task cancellation token source and cancellation token
            CancellationToken cancelConformToken = cancellationTokenSource.Token;

            // Create a test manager instance to oversee the test
            using (AlpacaProtocolTestManager alpacaTestManager = new(conformConfiguration, conformLogger, cancellationTokenSource, cancelConformToken))
            {

                try
                {
                    returnCode = alpacaTestManager.TestAlpacaProtocol().Result;
                }
                catch (Exception ex)
                {
                    returnCode = ex.HResult;
                    Console.WriteLine($"\r\n{RED_TEXT}Error running the Alpaca protocol test:{WHITE_TEXT}\r\n{ex}");
                }
            }

            GC.Collect();

            return returnCode;
        }

        /// <summary>
        /// Initialises configuration variables required to run the application
        /// </summary>
        /// <param name="logFileName"></param>
        /// <param name="path"></param>
        /// <param name="debugStartUp"></param>
        /// <param name="debugDiscovery"></param>
        /// <param name="resultsFileInfo"></param>
        /// <param name="settingsFileInfo"></param>
        private static bool InitialiseVariables(FileInfo logFileInfo, DirectoryInfo logPathInfo, bool debugStartUp, bool debugDiscovery, FileInfo resultsFileInfo, FileInfo settingsFileInfo)
        {
            bool closeDown = false; // Close down status: true = close down so that the 32bit process can run on its own, false to continue with this 64bit version

            argList = new();

            // Create and register logger objects and configuration services
            string loggerName;

            // Set log name with casing appropriate to OS
            if (OperatingSystem.IsWindows())
            {
                loggerName = "ConformU";
            }
            else
            {
                loggerName = "conformu";
            }

            string logFileName = logFileInfo?.FullName ?? "";
            string logFilePath = logPathInfo?.FullName ?? "";

            // Use fully qualified file name if present, otherwise use log file path and relative file name
            if (Path.IsPathFullyQualified(logFileName)) // Full file name and path provided so split into path and filename and ignore any supplied log file path
            {
                logFilePath = Path.GetDirectoryName(logFileName);
                logFileName = Path.GetFileName(logFileName);
            }
            else // Relative file name so use supplied log file name and path
            {
                // No action required
            }

            // Create a logger component
            conformLogger = new(logFileName, logFilePath, loggerName, true);

            // Create a state manager component
            conformStateManager = new(conformLogger);

            // Debug reading the configuration file if a command line option requires this
            if (debugStartUp) conformLogger.Debug = true;

            conformConfiguration = new(conformLogger, conformStateManager, settingsFileInfo?.FullName);
            conformLogger.Debug = conformConfiguration.Settings.Debug;

            // Enable logging of Alpaca discovery if a command line option requires this
            if (debugDiscovery) conformConfiguration.Settings.TraceDiscovery = true;

            // Set the results filename if supplied on the command line
            if (!string.IsNullOrEmpty(resultsFileInfo?.FullName)) conformConfiguration.Settings.ResultsFileName = resultsFileInfo.FullName;

            // Pass in the blazor server connection timeout
            argList.Add($"--{COMMAND_OPTION_CONNECTION_TIMEOUT}");
            argList.Add(conformConfiguration.Settings.ConnectionTimeout.ToString());

            #region Run as 32bit on Windows if required

            // Run as 32bit on a 64bit OS if configured to do so (Windows only!)
            if (OperatingSystem.IsWindows()) // OS is Windows
            {
                // Check whether we are running on 64bit Windows
                if (Environment.Is64BitOperatingSystem) // OS is 64bit
                {
                    // Test whether we are running in 64bit mode but the user has configured to start in 32bit mode
                    if ((Environment.Is64BitProcess) & (conformConfiguration.Settings.RunAs32Bit)) // Application is running in 64bit mode but the user has specified 32bit
                    {
                        // Restart ConformU using the 32bit executable

                        string baseFolder64 = AppContext.BaseDirectory;
                        string executable32 = Path.Join(baseFolder64, "32Bit", "conformu.exe");

                        // Don't try to run the 32bit application in the development environment!
                        if (!baseFolder64.Contains("\\bin\\"))
                        {
                            try
                            {
                                ProcessStartInfo processStartInfo = new()
                                {
                                    UseShellExecute = true,
                                    WindowStyle = ProcessWindowStyle.Minimized,
                                    FileName = "CMD.exe",
                                    Arguments = $"/c \"{executable32}\" {string.Join(" ", commandLineArguments)}"
                                };

                                Console.WriteLine($"\r\nStarting 32bit process");
                                Process.Start(processStartInfo);
                                Console.WriteLine($"32bit Process started");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"\r\n{RED_TEXT}Exception:{WHITE_TEXT}\r\n{ex}");
                            }

                            closeDown = true; // Set the close down flag
                        }
                    }
                }
            }

            #endregion

            return closeDown;
        }

        private static DeviceTechnology GetDeviceTechnology(string progIdOrUri, out ServiceType? serviceType, out string address, out int port, out int apiVersion, out DeviceTypes? deviceType, out int deviceNumber)
        {
            // Initialise return values
            DeviceTechnology returnValue = DeviceTechnology.NotSelected;

            serviceType = null;
            address = null;
            port = 0;
            apiVersion = 0;
            deviceType = null;
            deviceNumber = 0;

            string deviceList = "";
            foreach (string device in Devices.DeviceTypeNames())
            {
                deviceList += $"{device.ToLowerInvariant()}|";
            }

            deviceList = deviceList.Trim('|');

            // Test whether the device is an Alpaca device by testing whether the device description matches the expected pattern using a regex expression

            // Alpaca Regex - Groups
            // <Protocol Group> - http or https
            // <Address Group> - IPV4 address OR IPV6 address e.g. 192.168.1.1, 127.0.0.1, [::1] and [fe80:3::1ff:fe23:4567:8901]
            // <Port Group> - IP port number e.g. 32323
            // <DeviceType Group> - Device type e.g. telescope, camera etc.
            // <DeviceNumber Group> - Device number e.g.0, 1 etc.
            string alpacaPattern = $@"^(?i)(?<Protocol>https?):\/\/(?<Address>(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){{3}}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)|(?:\[(?:(([0-9A-Fa-f]{{1,4}}:){{7}}([0-9A-Fa-f]{{1,4}}|:))|(([0-9A-Fa-f]{{1,4}}:){{6}}(:[0-9A-Fa-f]{{1,4}}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}})|:))|(([0-9A-Fa-f]{{1,4}}:){{5}}(((:[0-9A-Fa-f]{{1,4}}){{1,2}})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}})|:))|(([0-9A-Fa-f]{{1,4}}:){{4}}(((:[0-9A-Fa-f]{{1,4}}){{1,3}})|((:[0-9A-Fa-f]{{1,4}})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}}))|:))|(([0-9A-Fa-f]{{1,4}}:){{3}}(((:[0-9A-Fa-f]{{1,4}}){{1,4}})|((:[0-9A-Fa-f]{{1,4}}){{0,2}}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}}))|:))|(?:(?:[0-9A-Fa-f]{{1,4}}:){{2}}(?:(?:(?::[0-9A-Fa-f]{{1,4}}){{1,5}})|((:[0-9A-Fa-f]{{1,4}}){{0,3}}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}}))|:))|(([0-9A-Fa-f]{{1,4}}:){{1}}(((:[0-9A-Fa-f]{{1,4}}){{1,6}})|((:[0-9A-Fa-f]{{1,4}}){{0,4}}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}}))|:))|(:(((:[0-9A-Fa-f]{{1,4}}){{1,7}})|((:[0-9A-Fa-f]{{1,4}}){{0,5}}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){{3}}))|:)))(?:%.+)?)\]):?(?<Port>[0-9]{{0,5}})(?-i)\/api\/v(?<ApiVersion>[0-9]*)\/(?<DeviceType>{deviceList})\/(?<DeviceNumber>[0-9]{{1,3}})\/?(?<Remainder>[0-9a-zA-Z]*)";

            // Test whether there is an IP address regex match
            Match alpacaMatch = Regex.Match(progIdOrUri, alpacaPattern, RegexOptions.CultureInvariant);

            // Check whether the IP address Alpaca match failed. 
            if (!alpacaMatch.Success) // IP address check failed so try a host name check
            {
                alpacaPattern = $@"^(?i)(?<Protocol>https?):\/\/(?<Address>([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{{0,61}}[a-zA-Z0-9])(\.([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{{0,61}}[a-zA-Z0-9]))*)?:?(?<Port>[0-9]{{0,5}})(?-i)\/api\/v(?<ApiVersion>[0-9]*)\/(?<DeviceType>{deviceList})\/(?<DeviceNumber>[0-9]{{1,3}})\/?(?<Remainder>[0-9a-zA-Z]*)";

                // Test whether there is a host name regex match
                alpacaMatch = Regex.Match(progIdOrUri, alpacaPattern, RegexOptions.CultureInvariant);
            }

            if (alpacaMatch.Success) // This is an Alpaca URI either using an IP address or a host name
            {
                // Extract the match groups to convenience variables
                string protocolParameter = alpacaMatch.Groups["Protocol"].Value;
                string addressParameter = alpacaMatch.Groups["Address"].Value;
                string portParameter = alpacaMatch.Groups["Port"].Value;
                string apiversionParameter = alpacaMatch.Groups["ApiVersion"].Value;
                string deviceNumberParameter = alpacaMatch.Groups["DeviceNumber"].Value;
                string deviceTypeParameter = alpacaMatch.Groups["DeviceType"].Value;

                //Console.WriteLine($"Protocol: {protocolParameter}, Address: {addressParameter}, Port: {portParameter}, Device number: {deviceNumberParameter}, Device type: {deviceTypeParameter}");

                serviceType = protocolParameter.ToLowerInvariant() == "http" ? ServiceType.Http : ServiceType.Https;

                // Test whether the Alpaca host address is empty
                if (string.IsNullOrEmpty(addressParameter))
                {
                    Console.WriteLine($"\r\n{RED_TEXT}The device identifier appears to be an Alpaca URI but the address cannot be parsed from it.{WHITE_TEXT}\r\n");
                    return returnValue;
                }
                address = addressParameter;
                if (string.IsNullOrEmpty(portParameter))
                {
                    Console.WriteLine($"\r\nNo Alpaca port number was supplied, assuming port 80.\r\n");
                    portParameter = "80";
                }
                port = Convert.ToInt32(portParameter);

                apiVersion = Convert.ToInt32(apiversionParameter);

                deviceNumber = Convert.ToInt32(deviceNumberParameter);

                deviceType = Devices.StringToDeviceType(deviceTypeParameter);

                returnValue = DeviceTechnology.Alpaca;
            }
            else // Not an Alpaca device so test for a COM ProgID
            {
                string progIdPattern = $@"^(?<DeviceName>[a-z0-9.]*)\.(?<DeviceType>{deviceList})(?<Tail>.*)";
                Match comMatch = Regex.Match(progIdOrUri, progIdPattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);


                if (comMatch.Success) // This is an COM ProgID
                {
                    string comNameParameter = comMatch.Groups["DeviceName"].Value;
                    string comDeviceTypeParameter = comMatch.Groups["DeviceType"].Value;
                    string comTailParameter = comMatch.Groups["Tail"].Value;
                    //Console.WriteLine($"COM name: {comNameParameter}, Device type: {comDeviceTypeParameter}, Tail parameter: {comTailParameter}.");

                    if (string.IsNullOrEmpty(comTailParameter))
                    {
                        deviceType = Devices.StringToDeviceType(comDeviceTypeParameter);
                        returnValue = DeviceTechnology.COM;
                    }
                    else
                    {
                        Console.WriteLine($"\r\n{RED_TEXT}The device type given in the COM ProgID is not a valid device type.{WHITE_TEXT}\r\n");
                    }
                }
                else
                {
                    Console.WriteLine($"\r\n{RED_TEXT}Unable to identify the device identifier as either a COM ProgID or an Alpaca URI. Is there a typo in the device identifier?{WHITE_TEXT}\r\n");
                }
            }

            return returnValue;
        }

        #endregion
    }
}