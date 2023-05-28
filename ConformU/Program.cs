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
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Linq;

namespace ConformU
{
    public partial class Program
    {
        private static string[] commandLineArguments;
        private static List<string> argList;
        private static ConformLogger conformLogger;
        private static ConformStateManager conformStateManager;
        private static ConformConfiguration conformConfiguration;

        #region Windows DLL imports

#if WINDOWS
        [LibraryImport("kernel32.dll")]
        static internal partial IntPtr GetConsoleWindow();

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static internal partial bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_SHOWMINIMIZED = 2;
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
                var rootCommand = new RootCommand($"Conform Universal {Update.ConformuVersionDisplayString}\r\nCopyright (c) 2021-{DateTime.Now.Year} Peter Simpson\r\n\r\n" +
                    $"Use conformu [command] -h for information on options available in each command.");

                // CONFORMANCE command
                var conformanceCommand = new Command("conformance", "Check the specified device for ASCOM device interface conformance");

                // ALPACA PROTOCOL command
                var alpacaProtocolCommand = new Command("alpacaprotocol", "Check the specified device for Alpaca protocol conformance");

                // CONFORMANCE USING SETTINGS command
                var conformanceUsingSettingsCommand = new Command("conformance-settings", "Check the device configured in the settings file for ASCOM device interface conformance");

                // ALPACA PROTOCOL USING SETTINGS command
                var alpacaUsingSettingsCommand = new Command("alpacaprotocol-settings", "Check the device configured in the settings file for Alpaca protocol conformance");

                // START AS GUI command
                var startAsGuiCommand = new Command("gui", "Start Conform Universal as an interactive GUI application with options to change the log file, results file and settings file locations");

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
                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(deviceArgument), out _, out _, out _, out _, out _);

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
                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(alpacaDeviceArgument), out _, out _, out _, out _, out _);

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
                Option<string> logFilePathOption = new(
                    aliases: new string[] { "-p", "--logfilepath" },
                    description: "Fully qualified path to the log file folder.\r\n" +
                    "Overrides the default log file path used by the GUI application, but is ignored when the --logfilename option is used.")
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
                        string logFilePath = (string)result.GetValueOrDefault();

                        // Check each character to see if it matches an invalid character
                        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
                        {
                            if (logFilePath.Contains(invalidCharacter)) // Found an invalid character detected
                            {
                                // Set the error message
                                result.ErrorMessage = $"\r\nLog file path contains invalid characters: {logFilePath}";
                            }
                        }
                    }
                });

                // LOG FILE NAME option
                Option<FileInfo> logFileNameOption = new(
                    aliases: new string[] { "-n", "--logfilename" },
                    description: "Filename of the log file (fully qualified or relative to the current directory).\r\n" +
                    "The default GUI log filename and location will be used if this option is omitted.")
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
                            if (logFileInfo.FullName.Contains(invalidCharacter)) // Found an invalid character detected
                            {
                                // Set the error message
                                result.ErrorMessage = $"\r\nLog file name contains invalid characters: {logFileInfo}";
                            }
                        }
                    }
                });

                // RESULTS FILE option
                Option<FileInfo> resultsFileOption = new(
                    aliases: new string[] { "-r", "--resultsfile" },
                    description: "Filename of the machine readable results file (fully qualified or relative to the current directory).\r\n" +
                    "The default GUI filename and location will be used if this option is omitted.")
                {
                    ArgumentHelpName = "FILENAME"
                };

                // SETTINGS FILE option
                Option<FileInfo> settingsFileOption = new(
                    aliases: new string[] { "-s", "--settingsfile" },
                    description: "Filename of the settings file to use (fully qualified or relative to the current directory).\r\n" +
                    "The GUI application settings file will be used if this option is omitted.")
                {
                    ArgumentHelpName = "FILENAME"
                };

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
                conformanceCommand.AddOption(settingsFileOption);
                conformanceCommand.AddOption(logFileNameOption);
                conformanceCommand.AddOption(logFilePathOption);
                conformanceCommand.AddOption(resultsFileOption);
                conformanceCommand.AddOption(debugDiscoveryOption);
                conformanceCommand.AddOption(debugStartUpOption);

                // ALPCA COMMAND - add arguments and options
                alpacaProtocolCommand.AddArgument(alpacaDeviceArgument);
                alpacaProtocolCommand.AddOption(settingsFileOption);
                alpacaProtocolCommand.AddOption(logFileNameOption);
                alpacaProtocolCommand.AddOption(logFilePathOption);
                alpacaProtocolCommand.AddOption(resultsFileOption);
                alpacaProtocolCommand.AddOption(debugDiscoveryOption);
                alpacaProtocolCommand.AddOption(debugStartUpOption);

                // ALPCA USING SETTINGS COMMAND - add options
                alpacaUsingSettingsCommand.AddOption(settingsFileOption);
                alpacaUsingSettingsCommand.AddOption(logFileNameOption);
                alpacaUsingSettingsCommand.AddOption(logFilePathOption);
                alpacaUsingSettingsCommand.AddOption(resultsFileOption);
                alpacaUsingSettingsCommand.AddOption(debugDiscoveryOption);
                alpacaUsingSettingsCommand.AddOption(debugStartUpOption);

                // CONFORMANCE USING SETTINGS COMMAND - add options
                conformanceUsingSettingsCommand.AddOption(settingsFileOption);
                conformanceUsingSettingsCommand.AddOption(logFileNameOption);
                conformanceUsingSettingsCommand.AddOption(logFilePathOption);
                conformanceUsingSettingsCommand.AddOption(resultsFileOption);
                conformanceUsingSettingsCommand.AddOption(debugDiscoveryOption);
                conformanceUsingSettingsCommand.AddOption(debugStartUpOption);

                // START AS GUI COMMAND  - add options
                startAsGuiCommand.AddOption(logFileNameOption);
                startAsGuiCommand.AddOption(logFilePathOption);
                startAsGuiCommand.AddOption(debugDiscoveryOption);
                startAsGuiCommand.AddOption(debugStartUpOption);
                startAsGuiCommand.AddOption(resultsFileOption);
                startAsGuiCommand.AddOption(settingsFileOption);

                #endregion

                #region Command handlers

                // ROOT COPMMAND handler - Starts the GUI or displays version information
                rootCommand.SetHandler((bool version) =>
                {
                    // Return a task with the required return code
                    return Task.FromResult(RootCommandHandler(version));
                }, versionOption);

                // CONFORMANCE USING SETTINGS COPMMAND handler
                conformanceUsingSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    int returnCode = -8888;

                    // Initialise required variables required by several commands
                    InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);

                    // Run the conformance test
                    returnCode = RunConformanceTest();

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // CONFORMANCE COMMAND handler
                conformanceCommand.SetHandler((device, file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    int returnCode = 0;

                    DeviceTechnology technology = GetDeviceTechnology(device, out ServiceType? serviceType, out string address, out int port, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD *****\r\nArgument: {device}");
                            break;

                        case DeviceTechnology.Alpaca:
                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunConformanceTest();
                            break;

                        case DeviceTechnology.COM:
                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the COM device parameters
                            conformConfiguration.SetComDevice(device, deviceType.Value);

                            // Run the conformance test
                            returnCode = RunConformanceTest();
                            break;

                        default:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE ONLY COM AND ALPACA TECHNOLOGY TYPES ARE SUPPORTED *****\r\nArgument: {device}");
                            break;
                    }

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, deviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // ALPACA PROTOCOL COMMAND handler
                alpacaProtocolCommand.SetHandler((alpacaDevice, file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    int returnCode = 0;

                    DeviceTechnology technology = GetDeviceTechnology(alpacaDevice, out ServiceType? serviceType, out string address, out int port, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD *****\r\nArgument: {alpacaDevice}");
                            break;

                        case DeviceTechnology.Alpaca:
                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunAlpacaProtocolTest();
                            break;

                        case DeviceTechnology.COM:
                            Console.WriteLine($"***** This command can only be used with Alpaca devices *****\r\nArgument: {alpacaDevice}");
                            break;

                        default:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE ONLY COM AND ALPACA TECHNOLOGY TYPES ARE SUPPORTED *****\r\nArgument: {alpacaDevice}");
                            break;
                    }

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, alpacaDeviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // ALPACA USING SETTINGS COMMAND handler
                alpacaUsingSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartUp) =>
                {
                    InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);
                    // Return a task with the required return code
                    return Task.FromResult(RunAlpacaProtocolTest());

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

                // START AS GUI COMMAND handler
                startAsGuiCommand.SetHandler((file, path, results, settings, debugDiscovery, debugStartUp) =>
                {
                    // Return a task with the required return code
                    return Task.FromResult(StartGuiHandler(file, path, debugStartUp, debugDiscovery, results, settings));
                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartUpOption);

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
                Console.WriteLine($"Conform internal error - Exception reading command line or stating application:\r\n{ex}");
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
        static int RootCommandHandler(bool version)
        {
            if (version)
            {
                Console.WriteLine($"\r\nConform Universal {Update.ConformuVersionDisplayString}");
            }
            else
            {
                // Initialise required variables required by several commands
                InitialiseVariables(null, null, false, false, null, null);

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

                    Console.WriteLine($"Starting web server.");
                    Task t = CreateHostBuilder(conformLogger, conformStateManager, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept the / prefix
                         .Build()
                         .RunAsync();

                    t.Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception stating application: {ex.Message}");
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
        static int StartGuiHandler(FileInfo file, string path, bool debugStartUp, bool debugDiscovery, FileInfo resultsFile, FileInfo settingsFile)
        {
            // Initialise required variables required by several commands
            InitialiseVariables(file, path, debugStartUp, debugDiscovery, resultsFile, settingsFile);

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

                Console.WriteLine($"Starting web server.");
                Task t = CreateHostBuilder(conformLogger, conformStateManager, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept the / prefix
                     .Build()
                     .RunAsync();

                t.Wait();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception stating application: {ex.Message}");
                return ex.HResult;
            }
            return 0;
        }

        public static IHostBuilder CreateHostBuilder(ConformLogger conformLogger, ConformStateManager conformStateManager, ConformConfiguration conformConfiguration, string[] args)
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

        static int RunConformanceTest()
        {
            int returnCode;

            // Validate the supplied configuration and only start if there are no settings issues
            string validationMessage = conformConfiguration.Validate();
            if (!string.IsNullOrEmpty(validationMessage)) // There is a configuration issue so present an error message
            {
                Console.WriteLine($"Cannot start test:\r\n{validationMessage}");
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
                    returnCode = tester.TestDevice();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"RunConformanceTest - Exception: \r\n {ex}");
                    returnCode = ex.HResult;
                }
            }

            GC.Collect();

            return returnCode;
        }

        static int RunAlpacaProtocolTest()
        {
            int returnCode;

            // Validate the supplied configuration and only start if there are no settings issues
            string validationMessage = conformConfiguration.Validate();
            if (!string.IsNullOrEmpty(validationMessage)) // There is a configuration issue so present an error message
            {
                Console.WriteLine($"Cannot start test:\r\n{validationMessage}");
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
                    Console.WriteLine($"Error running the Alpaca protocol test:\r\n{ex}");
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

        private static void InitialiseVariables(FileInfo logFileInfo, string logFilePath, bool debugStartUp, bool debugDiscovery, FileInfo resultsFileInfo, FileInfo settingsFileInfo)
        {
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
            if (string.IsNullOrEmpty(logFilePath))
                logFilePath = "";

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

                                Console.WriteLine($"Starting 32bit process");
                                Process.Start(processStartInfo);
                                Console.WriteLine($"32bit Process started");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Exception: {ex}");
                            }

                            return;
                        }
                    }
                }
            }

            #endregion
        }

        private static DeviceTechnology GetDeviceTechnology(string progIdOrUri, out ServiceType? serviceType, out string address, out int port, out DeviceTypes? deviceType, out int deviceNumber)
        {
            // Initialise return values
            DeviceTechnology returnValue = DeviceTechnology.NotSelected;

            serviceType = null;
            address = null;
            port = 0;
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
            string alpacaPattern = @"^" + // Start of URI
                                          // Extract http or https prefix
                @"(?i)(?<Protocol>https?)" +
                // Ignore fixed :// text
                @":\/\/" +
                // Extract IPv4 address
                @"(?<Address>(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)|" +
                @"(?:\[(?:(([0-9A-Fa-f]{1,4}:){7}([0-9A-Fa-f]{1,4}|:))|(([0-9A-Fa-f]{1,4}:){6}(:[0-9A-Fa-f]{1,4}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))" +
                // OR
                @"|" +
                // Extract IPv6 address
                @"(([0-9A-Fa-f]{1,4}:){5}(((:[0-9A-Fa-f]{1,4}){1,2})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){4}(((:" +
                @"[0-9A-Fa-f]{1,4}){1,3})|((:[0-9A-Fa-f]{1,4})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){3}(((:[0-9A-Fa-f]" +
                @"{1,4}){1,4})|((:[0-9A-Fa-f]{1,4}){0,2}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(?:(?:[0-9A-Fa-f]{1,4}:){2}(?:(?:(?::[0-9A-Fa-f]" +
                @"{1,4}){1,5})|((:[0-9A-Fa-f]{1,4}){0,3}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){1}(((:[0-9A-Fa-f]" +
                @"{1,4}){1,6})|((:[0-9A-Fa-f]{1,4}){0,4}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(:(((:[0-9A-Fa-f]{1,4}){1,7})|((:[0-9A-Fa-f]" +
                @"{1,4}){0,5}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:)))(?:%.+)?)\])" +
                // Ignore fixed colon port number separator if present
                @":?" +
                // Extract port number if present
                @"(?<Port>[0-9]{0,5})(?-i)" +
                // Ignore fixed /api/v1/ text
                @"\/api\/v1\/" +
                // Extract the device type
                @"(?<DeviceType>" + deviceList + @")" +
                // Ignore fixed / text
                @"\/" +
                // Extract the device number
                @"(?<DeviceNumber>[0-9]{1,3})" +
                // Ignore fixed / text if present
                @"\/?" +
                // Ignore any remaining text at end of the device identifier
                @"(?<Remainder>[0-9a-zA-Z]*)";

            Match alpacaMatch = Regex.Match(progIdOrUri, alpacaPattern, RegexOptions.CultureInvariant);
            if (alpacaMatch.Success) // This is an Alpaca URI
            {
                // Extract the match groups to convenience variables
                string protocolParameter = alpacaMatch.Groups["Protocol"].Value;
                string addressParameter = alpacaMatch.Groups["Address"].Value;
                string portParameter = alpacaMatch.Groups["Port"].Value;
                string deviceNumberParameter = alpacaMatch.Groups["DeviceNumber"].Value;
                string deviceTypeParameter = alpacaMatch.Groups["DeviceType"].Value;

                Console.WriteLine($"Protocol: {protocolParameter}, Address: {addressParameter}, Port: {portParameter}, Device number: {deviceNumberParameter}, Device type: {deviceTypeParameter}");

                serviceType = protocolParameter.ToLowerInvariant() == "http" ? ServiceType.Http : ServiceType.Https;

                // Test whether the Alpaca host address is empty
                if (string.IsNullOrEmpty(addressParameter))
                {
                    Console.WriteLine($"\r\nThe device identifier appears to be an Alpaca URI but the address cannot be parsed from it\r\n");
                    return returnValue;
                }
                address = addressParameter;
                if (string.IsNullOrEmpty(portParameter))
                {
                    Console.WriteLine($"\r\nNo Alpaca port number was supplied, assuming port 80.\r\n");
                    portParameter = "80";
                }
                port = Convert.ToInt32(portParameter);
                deviceNumber = Convert.ToInt32(deviceNumberParameter);

                deviceType = Devices.StringToDeviceType(deviceTypeParameter);

                returnValue = DeviceTechnology.Alpaca;
            }
            else // Not an Alpaca device so test for a COM ProgID
            {
                string progIdPattern = @"^(?<DeviceName>[a-z0-9.]*)\.(?<DeviceType>" + deviceList + @")(?<Tail>.*)";
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
                        Console.WriteLine($"\r\nThe device type given in the COM ProgID is not a valid device type.\r\n");
                    }
                }
                else
                {
                    Console.WriteLine($"\r\nUnable to identify the device identifier as either a COM ProgID or an Alpaca URI. Is there a typo in the device identifier?\r\n");
                }
            }

            return returnValue;
        }

        #endregion
    }
}