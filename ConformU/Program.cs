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

namespace ConformU
{
    public class Program
    {
        private static string[] commandLineArguments;
        private static List<string> argList;
        private static ConformLogger conformLogger;
        private static ConformStateManager conformStateManager;
        private static ConformConfiguration conformConfiguration;
        private static int returnCode = 0;

#if WINDOWS
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOWMINIMIZED = 2;
        const int SW_SHOW = 5;
#endif

        public static async Task<int> Main(string[] args)
        {
            try
            {
                // Save the command line arguments so they can be reused if a 32bit application is required
                commandLineArguments = args;

                #region Command definitions

                // Define the root command
                var rootCommand = new RootCommand($"Conform Universal {Update.ConformuVersionDisplayString}\r\nCopyright (c) 2021-{DateTime.Now.Year} Peter Simpson\r\n\r\n" +
                    $"Use conformu [command] -h for information on options available in each command.");

                // Define other commands
                var conformanceCommand = new Command("conformance", "Check the specified device for ASCOM device interface conformance");

                var alpacaCommand = new Command("alpaca", "Check the specified device for Alpaca protocol conformance");

                var conformanceSettingsCommand = new Command("conformance-settings", "Check the device configured in the settings file for ASCOM device interface conformance");

                var alpacaSettingsCommand = new Command("alpaca-settings", "Check the device configured in the settings file for Alpaca protocol conformance");

                var guiCommand = new Command("gui", "Start Conform Universal as an interactive GUI application with options to change the log file, results file and settings file locations");

                #endregion

                #region Argument definitions

                // Define arguments that can be applied to commands
                Argument<string> deviceArgument = new Argument<string>(
                    name: "COM_ProgID_or_Alpaca_URI",
                    description: "The device's COM ProgID (e.g. ASCOM.Simulator.Telescope) or its Alpaca URI (e.g. http://[host]:[port]/api/v1/[DeviceType]/[DeviceNumber]).\r\n" +
                    "The device technology (COM or Alpaca) and the ASCOM device type will be inferred from the supplied ProgID or URI.");

                deviceArgument.AddValidator((commandResult) =>
                {
                    Console.WriteLine($"Validator called - Children.Count: {commandResult.Children.Count}, Symbol: {commandResult.GetValueForArgument(deviceArgument)}");

                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(deviceArgument), out _, out _, out _, out _, out _);

                    if (technology == DeviceTechnology.NotSelected)
                    {
                        commandResult.ErrorMessage = $"The COM ProgID or Alpaca URI argument is invalid: {commandResult.GetValueForArgument(deviceArgument)}";
                    }
                    else
                    {
                        Console.WriteLine($"Validated {technology} device successfully!");
                    }
                });

                // Define arguments that can be applied to commands
                Argument<string> alpacaDeviceArgument = new Argument<string>(
                    name: "Alpaca_URI",
                    description: "The device's Alpaca URI (e.g. http://host:port/api/v1/[DeviceType]/[DeviceNumber]).\r\n" +
                    "The ASCOM device type will be inferred from the supplied URI.");

                alpacaDeviceArgument.AddValidator((commandResult) =>
                {
                    Console.WriteLine($"Validator called - Children.Count: {commandResult.Children.Count}, Symbol: {commandResult.GetValueForArgument(alpacaDeviceArgument)}");

                    DeviceTechnology technology = GetDeviceTechnology(commandResult.GetValueForArgument(alpacaDeviceArgument), out _, out _, out _, out _, out _);

                    if (technology == DeviceTechnology.NotSelected)
                    {
                        commandResult.ErrorMessage = $"The COM ProgID or Alpaca URI argument is invalid: {commandResult.GetValueForArgument(alpacaDeviceArgument)}";
                    }
                    else
                    {
                        Console.WriteLine($"Validated {technology} device successfully!");
                    }
                });

                #endregion

                #region Option definitions

                // Define options that can be added to commands
                Option<bool> versionOption = new Option<bool>(
                    aliases: new string[] { "-v", "--version" },
                    description: "Show the Conform Universal version number");

                Option<bool> debugDiscoveryOption = new Option<bool>(
                    aliases: new string[] { "-d", "--debugdiscovery" },
                    description: "Write discovery debug information to the log.");
                debugDiscoveryOption.IsHidden = true;

                Option<bool> debugStartupOption = new Option<bool>(
                    aliases: new string[] { "-t", "--debugstratup" },
                    description: "Write start-up debug information to the log.");
                debugStartupOption.IsHidden = true;

                Option<string> logFilePathOption = new Option<string>(
                    aliases: new string[] { "-p", "--logfilepath" },
                    description: "Fully qualified path to the log file folder.\r\n" +
                    "Overrides the default log file path used by the GUI application, but is ignored when a fully qualified name is specified in the --logfile option.");
                logFilePathOption.ArgumentHelpName = "PATH";

                Option<FileInfo> logFileNameOption = new Option<FileInfo>(
                    aliases: new string[] { "-n", "--logfile" },
                    description: "Filename of the log file (fully qualified or relative to the current directory).\r\n" +
                    "The default GUI log filename and location will be used if this option is omitted.");
                logFileNameOption.ArgumentHelpName = "FILENAME";

                Option<FileInfo> resultsFileOption = new Option<FileInfo>(
                    aliases: new string[] { "-r", "--resultsfile" },
                    description: "Filename of the machine readable results file (fully qualified or relative to the current directory).\r\n" +
                    "The default GUI filename and location will be used if this option is omitted.");
                resultsFileOption.ArgumentHelpName = "FILENAME";

                Option<FileInfo> settingsFileOption = new Option<FileInfo>(
                    aliases: new string[] { "-s", "--settingsfile" },
                    description: "Filename of the settings file to use (fully qualified or relative to the current directory).\r\n" +
                    "The GUI application settings file will be used if this option is omitted.");
                settingsFileOption.ArgumentHelpName = "FILENAME";

                #endregion

                #region Associate arguments and options with commands

                // Add commands and options to the root command
                rootCommand.AddCommand(conformanceCommand);
                rootCommand.AddCommand(alpacaCommand);
                rootCommand.AddCommand(conformanceSettingsCommand);
                rootCommand.AddCommand(alpacaSettingsCommand);
                rootCommand.AddCommand(guiCommand);
                rootCommand.AddOption(versionOption);

                // Add options to the start ConformU with an interactive GUI command
                guiCommand.AddOption(logFileNameOption);
                guiCommand.AddOption(logFilePathOption);
                guiCommand.AddOption(debugDiscoveryOption);
                guiCommand.AddOption(debugStartupOption);
                guiCommand.AddOption(resultsFileOption);
                guiCommand.AddOption(settingsFileOption);

                // Add options to the check Alpaca protocol settings command
                alpacaSettingsCommand.AddOption(settingsFileOption);
                alpacaSettingsCommand.AddOption(logFileNameOption);
                alpacaSettingsCommand.AddOption(logFilePathOption);
                alpacaSettingsCommand.AddOption(resultsFileOption);
                alpacaSettingsCommand.AddOption(debugDiscoveryOption);
                alpacaSettingsCommand.AddOption(debugStartupOption);

                // Add options to the check conformance settings command
                conformanceSettingsCommand.AddOption(settingsFileOption);
                conformanceSettingsCommand.AddOption(logFileNameOption);
                conformanceSettingsCommand.AddOption(logFilePathOption);
                conformanceSettingsCommand.AddOption(resultsFileOption);
                conformanceSettingsCommand.AddOption(debugDiscoveryOption);
                conformanceSettingsCommand.AddOption(debugStartupOption);

                // Add arguments and options to the conformance command
                conformanceCommand.AddArgument(deviceArgument);
                conformanceCommand.AddOption(settingsFileOption);
                conformanceCommand.AddOption(logFileNameOption);
                conformanceCommand.AddOption(logFilePathOption);
                conformanceCommand.AddOption(resultsFileOption);
                conformanceCommand.AddOption(debugDiscoveryOption);
                conformanceCommand.AddOption(debugStartupOption);

                // Add arguments and options to the alpaca command
                alpacaCommand.AddArgument(alpacaDeviceArgument);
                alpacaCommand.AddOption(settingsFileOption);
                alpacaCommand.AddOption(logFileNameOption);
                alpacaCommand.AddOption(logFilePathOption);
                alpacaCommand.AddOption(resultsFileOption);
                alpacaCommand.AddOption(debugDiscoveryOption);
                alpacaCommand.AddOption(debugStartupOption);

                #endregion

                // Create a custom parser what does not present the default version description
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

                // Add the root command handler to start the GUI and display version information
                rootCommand.SetHandler((bool version) =>
                {
                    // Return a task with the required return code
                    return Task.FromResult(RunRootCommand(version));
                }, versionOption);

                // Add the conformance settings command handler
                conformanceSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartup) =>
                {
                    int returnCode = -8888;

                    // Initialise required variables required by several commands
                    InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

                    // Run the conformance test
                    returnCode = RunConformanceSettings();

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartupOption);

                // Add the conformance command handler
                conformanceCommand.SetHandler((device, file, path, resultsFile, settingsFile, debugDiscovery, debugStartup) =>
                {
                    int returnCode = -8888;
                    Console.WriteLine($"SetHandler called with argument {device}");

                    DeviceTechnology technology = GetDeviceTechnology(device, out ServiceType? serviceType, out string address, out int port, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD *****\r\nArgument: {device}");
                            break;

                        case DeviceTechnology.Alpaca:
                            Console.WriteLine($"Alpaca Device: {device}, Device type: {deviceType.Value}");

                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunConformanceSettings();
                            break;

                        case DeviceTechnology.COM:
                            Console.WriteLine($"COM Device: {device}, Device type: {deviceType.Value}");

                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the COM device parameters
                            conformConfiguration.SetComDevice(device, deviceType.Value);

                            // Run the conformance test
                            returnCode = RunConformanceSettings();
                            break;

                        default:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE ONLY COM AND ALPACA TECHNOLOGY TYPES ARE SUPPORTED *****\r\nArgument: {device}");
                            break;
                    }

                    // Return the return code
                    return Task.FromResult(returnCode);

                }, deviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartupOption);

                // Add the alpaca command handler
                alpacaCommand.SetHandler((alpacaDevice, file, path, resultsFile, settingsFile, debugDiscovery, debugStartup) =>
                {
                    int returnCode = -8888;
                    Console.WriteLine($"SetHandler called with argument: '{alpacaDevice}'");

                    DeviceTechnology technology = GetDeviceTechnology(alpacaDevice, out ServiceType? serviceType, out string address, out int port, out DeviceTypes? deviceType, out int deviceNumber);

                    switch (technology)
                    {
                        case DeviceTechnology.NotSelected:
                            Console.WriteLine($"***** SHOULD NEVER SEE THIS MESSAGE BECAUSE THE VALIDATOR SHOLD HAVE STOPPED EXECUTION IF THE DEVICE PROGID/URI IS BAD *****\r\nArgument: {alpacaDevice}");
                            break;

                        case DeviceTechnology.Alpaca:
                            Console.WriteLine($"Alpaca Device: {alpacaDevice}, Device type: {deviceType.Value}");

                            // Initialise required variables required by several commands
                            InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

                            // Set options to conduct a full conformance test
                            conformConfiguration.SetFullTest();

                            // Set the Alpaca device parameters
                            conformConfiguration.SetAlpacaDevice(serviceType.Value, address, port, deviceType.Value, deviceNumber);

                            // Run the conformance test
                            returnCode = RunAlpacaSettings();
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

                }, alpacaDeviceArgument, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartupOption);

                // Add the command line Alpaca protocol check handler
                alpacaSettingsCommand.SetHandler((file, path, resultsFile, settingsFile, debugDiscovery, debugStartup) =>
                {
                    InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);
                    // Return a task with the required return code
                    return Task.FromResult(RunAlpacaSettings());

                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartupOption);

                // Add the GUI command handler
                guiCommand.SetHandler((file, path, results, settings, debugDiscovery, debugStartup) =>
                {
                    // Return a task with the required return code
                    return Task.FromResult(RunGuiCommand(file, path, debugStartup, debugDiscovery, results, settings));
                }, logFileNameOption, logFilePathOption, resultsFileOption, settingsFileOption, debugDiscoveryOption, debugStartupOption);

                // Parse the command line and invoke actions as determined by supplied parameters
                return await parser.InvokeAsync(args);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception reading command line or stating application:\r\n{ex}");
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
        static int RunRootCommand(bool version)
        {
            if (version)
            {
                Console.WriteLine($"\r\nConform Universal {Update.ConformuVersionDisplayString}");
            }
            else
            {
                // Initialise required variables required by several commands
                InitialiseVariables(null, null, false, false, null, null);

                Console.WriteLine($"Running Conform Universal as a GUI application.");
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
        /// <param name="debugStartup"></param>
        /// <param name="debugDiscovery"></param>
        /// <param name="resultsFile"></param>
        /// <param name="settingsFile"></param>
        static int RunGuiCommand(FileInfo file, string path, bool debugStartup, bool debugDiscovery, FileInfo resultsFile, FileInfo settingsFile)
        {
            // Initialise required variables required by several commands
            InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

            Console.WriteLine($"Running Conform Universal as a GUI application.");
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

        static int RunConformanceSettings()
        {
            Console.WriteLine($"Running Conform Universal as a Console application.");

            foreach (string s in argList)
            { Console.WriteLine($"ARG = '{s}'"); }

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
                    Console.WriteLine($"StartTest - Exception: \r\n {ex}");
                    returnCode = ex.HResult;
                }
            }

            GC.Collect();

            return returnCode;
        }

        static int RunAlpacaSettings()
        {
            // Initialise required variables required by several commands
            //InitialiseVariables(file, path, debugStartup, debugDiscovery, resultsFile, settingsFile);

            Console.WriteLine($"Running Conform Universal as a Console application.");

            foreach (string s in argList)
            { Console.WriteLine($"ARG = '{s}'"); }

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

            AlpacaTestManager alpacaTestmanager = new AlpacaTestManager(conformConfiguration, conformLogger, cancellationTokenSource, cancelConformToken);

            int returnCode = -99997;
            try
            {
                returnCode = alpacaTestmanager.TestAlpacaProtocol().Result;
            }
            catch (Exception ex)
            {
                returnCode = -99998;
                Console.WriteLine($"Error running the Alpaca protocol test:\r\n{ex}");
            }

            GC.Collect();

            return returnCode;
        }
        /// <summary>
        /// Initialises configuration variables required to run the application
        /// </summary>
        /// <param name="logFileName"></param>
        /// <param name="path"></param>
        /// <param name="debugStartup"></param>
        /// <param name="debugDiscovery"></param>
        /// <param name="resultsFileInfo"></param>
        /// <param name="settingsFileInfo"></param>

        private static void InitialiseVariables(FileInfo logFileInfo, string logFilePath, bool debugStartup, bool debugDiscovery, FileInfo resultsFileInfo, FileInfo settingsFileInfo)
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
            if (debugStartup) conformLogger.Debug = true;

            conformConfiguration = new(conformLogger, conformStateManager, settingsFileInfo?.FullName);
            conformLogger.Debug = conformConfiguration.Settings.Debug;

            // Enable logging of Alpaca discovery if a command line option requires this
            if (debugDiscovery) conformConfiguration.Settings.TraceDiscovery = true;

            // Set the results filename if supplied on the command line
            if (!string.IsNullOrEmpty(resultsFileInfo?.FullName)) conformConfiguration.Settings.ResultsFileName = resultsFileInfo.FullName;

            // Pass in the blazor server connection timeout
            Console.WriteLine($"Server connection timeout: {conformConfiguration.Settings.ConnectionTimeout}");
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
                        string executable32 = Path.Join(baseFolder64.Replace("64", "32"), "conformu.exe");
                        Console.WriteLine($"Base directory: {AppContext.BaseDirectory}, EXE32: {executable32}");

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
            string alpacaPattern = @"^(?i)(?<Protocol>https?):\/\/(?<Address>[a-zA-Z0-9.]*|\[?[0-9A-F:]*\]?):?(?<Port>[0-9]{0,5})(?-i)\/api\/v1\/(?<DeviceType>" + deviceList + @")\/(?<DeviceNumber>[0-9])\/?(?<Remainder>[0-9a-zA-Z]*)";
            Console.WriteLine($"Alpaca pattern:\r\n{alpacaPattern}");

            Match alpacaMatch = Regex.Match(progIdOrUri, alpacaPattern, RegexOptions.CultureInvariant);
            if (alpacaMatch.Success) // This is an Alpaca URI
            {
                Console.WriteLine($"Alpaca match succeeded for '{progIdOrUri}'");
                serviceType = alpacaMatch.Groups["Protocol"].Value.ToLowerInvariant() == "http" ? ServiceType.Http : ServiceType.Https;
                address = alpacaMatch.Groups["Address"].Value;
                port = Convert.ToInt32(alpacaMatch.Groups["Port"].Value);
                deviceNumber = Convert.ToInt32(alpacaMatch.Groups["DeviceNumber"].Value);

                deviceType = Devices.StringToDeviceType(alpacaMatch.Groups["DeviceType"].Value);

                Console.WriteLine($"  Protocol: {serviceType}, Address: {address}, Port: {port}, Device type: {deviceType}, Device number: {deviceNumber}");
                returnValue = DeviceTechnology.Alpaca;
            }
            else // Not an Alpaca device so test for a COM ProgID
            {
                Console.WriteLine($"Alpaca device match failed for '{progIdOrUri}', testing for COM ProgID");
                string progIdPattern = @"^(?i)(?<DeviceName>[a-z0-9.]*)\.(?<DeviceType>" + deviceList + @")";
                Match comMatch = Regex.Match(progIdOrUri, progIdPattern, RegexOptions.CultureInvariant);
                if (comMatch.Success) // This is an COM ProgID
                {
                    Console.WriteLine($"ProgID match succeeded for '{progIdOrUri}'");
                    deviceType = Devices.StringToDeviceType(comMatch.Groups["DeviceType"].Value);
                    Console.WriteLine($"  Device type: {deviceType}");
                    returnValue = DeviceTechnology.COM;
                }
                else
                {
                    Console.WriteLine($"ProgID match failed for '{progIdOrUri}'");

                }

            }



            return returnValue;
        }


    }
}