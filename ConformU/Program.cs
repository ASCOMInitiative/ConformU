using CommandLine;
using CommandLine.Text;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static ConformU.Globals;

namespace ConformU
{
    public class Program
    {
        private static string[] commandLineArguments;

#if WINDOWS
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOWMINIMIZED = 2;
        const int SW_SHOW = 5;
#endif

        public static void Main(string[] args)
        {
            try
            {
                // Save the command line arguments so they can be reused if a 32bit application is required
                commandLineArguments = args;

                // Parse the command line, options are specified in the CommandLine Options class
                var parser = new Parser(with =>
                {
                    with.HelpWriter = null;
                    with.AutoVersion = false;
                });
                var parserResult = parser.ParseArguments<CommandLineOptions>(args);

                parserResult.MapResult(
                    (CommandLineOptions commandLineOptions) => Run(commandLineOptions),
                    errs => DisplayHelp<CommandLineOptions>(parserResult));

                return;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception reading command line or stating application:\r\n{ex}");
            }
        }

        static int DisplayHelp<T>(ParserResult<T> result)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                //configure help
                h.AdditionalNewLineAfterOption = true;
                h.Heading = $"Conform Universal {Update.ConformuVersionDisplayString}"; //change header
                h.Copyright = $"Copyright (c) 2021-{DateTime.Now.Year} Peter Simpson\r\n"; //change copyright text
                //h.MaximumDisplayWidth = 10000;
                h.AddPreOptionsText("ASCOM Universal Conformance Checker - Tests an Alpaca or COM device to ensure that it conforms to the relevant ASCOM interface specification.");
                return h;
            });
            Console.WriteLine(helpText);
            return 1;
        }

        static int Run(CommandLineOptions commandLineOptions)
        {
            List<string> argList = new();

            int returnCode = 0;

            // Display the applicati0on version
            if (commandLineOptions.Run)
            {
                Console.WriteLine($"Conform Universal {Update.ConformuVersionDisplayString}");
                Console.WriteLine($"Copyright (c) 2023-{DateTime.Now.Year} Peter Simpson");
                Console.WriteLine($"\r\nThe '--commandline' parameter has been withdrawn, please use '--conformancecheck' instead.");
                return -98;
            }

            // Display the applicati0on version
            if (commandLineOptions.Version)
            {
                Console.WriteLine($"Conform Universal {Update.ConformuVersionDisplayString}");
                Console.WriteLine($"Copyright (c) 2023-{DateTime.Now.Year} Peter Simpson");
                return 0;
            }

            // Extract the settings file location if provided
            if (!string.IsNullOrEmpty(commandLineOptions.SettingsFileLocation))
            {
                Console.WriteLine($"Settings file location: {commandLineOptions.SettingsFileLocation}");
                argList.Add($"--{COMMAND_OPTION_SETTINGS}");
                argList.Add(commandLineOptions.SettingsFileLocation);
            }

            // Extract the log file location if provided
            if (!string.IsNullOrEmpty(commandLineOptions.LogFileName))
            {
                Console.WriteLine($"Log file location: {commandLineOptions.LogFileName}");
                argList.Add($"--{COMMAND_OPTION_LOGFILENAME}");
                argList.Add(commandLineOptions.LogFileName);
            }

            // Extract the log file path if provided
            if (!string.IsNullOrEmpty(commandLineOptions.LogFilePath))
            {
                Console.WriteLine($"Log file path: {commandLineOptions.LogFilePath}");
                argList.Add($"--{COMMAND_OPTION_LOGFILEPATH}");
                argList.Add(commandLineOptions.LogFilePath);
            }

            // Flag if discovery debug information should be included in the log file
            if (commandLineOptions.DebugDiscovery.HasValue)
            {
                argList.Add($"--{COMMAND_OPTION_DEBUG_DISCOVERY}");
                argList.Add(commandLineOptions.DebugDiscovery.ToString());
            }

            // Flag if start-up debug information should be included in the log file
            if (commandLineOptions.DebugStartup)
            {
                argList.Add($"--{COMMAND_OPTION_DEBUG_STARTUP}");
                argList.Add(commandLineOptions.DebugStartup.ToString());
            }

            // Set the results filename if supplied 
            if (!string.IsNullOrEmpty(commandLineOptions.ResultsFileName))
            {
                argList.Add($"--{COMMAND_OPTION_RESULTS_FILENAME}");
                argList.Add(commandLineOptions.ResultsFileName);
            }

            #region Create and register logger and configuration services

            // Create logger and configuration objects
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

            string logFileName = commandLineOptions.LogFileName ?? "";
            string logFilePath = commandLineOptions.LogFilePath ?? "";

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
            ConformLogger conformLogger = new(logFileName, logFilePath, loggerName, true);

            // Create a state manager component
            ConformStateManager conformStateManager = new(conformLogger);

            // Debug reading the configuration file if a command line option requires this
            if (commandLineOptions.DebugStartup) conformLogger.Debug = true;

            ConformConfiguration conformConfiguration = new(conformLogger, conformStateManager, commandLineOptions.SettingsFileLocation);
            conformLogger.Debug = conformConfiguration.Settings.Debug;

            // Enable logging of Alpaca discovery if a command line option requires this
            if (commandLineOptions.DebugDiscovery.HasValue) conformConfiguration.Settings.TraceDiscovery = commandLineOptions.DebugDiscovery.Value;

            // Set the results filename if supplied on the command line
            if (!string.IsNullOrEmpty(commandLineOptions.ResultsFileName)) conformConfiguration.Settings.ResultsFileName = commandLineOptions.ResultsFileName;

            // Pass in the blazor server connection timeout
            Console.WriteLine($"Server connection timeout: {conformConfiguration.Settings.ConnectionTimeout}");
            argList.Add($"--{COMMAND_OPTION_CONNECTION_TIMEOUT}");
            argList.Add(conformConfiguration.Settings.ConnectionTimeout.ToString());

            #endregion

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

                            return 0;
                        }
                    }
                }
            }

            Console.WriteLine($"Running as a {(Environment.Is64BitProcess ? "64bit" : "32bit")} application");
            Console.WriteLine($"Base directory: {AppContext.BaseDirectory}");

            #endregion

            // Run from command line if requested
            if (commandLineOptions.ConformancCheck)
            {
                foreach (string s in argList)
                { Console.WriteLine($"ARG = '{s}'"); }

                // Set the report file location if required
                if (!string.IsNullOrEmpty(commandLineOptions.ResultsFileName)) conformConfiguration.Settings.ResultsFileName = commandLineOptions.ResultsFileName;

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
                ConformanceTestManager tester = new(conformConfiguration, conformLogger, cancellationTokenSource, cancelConformToken);
                try
                {
                    returnCode = tester.TestDevice();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"StartTest - Exception: \r\n {ex}");
                }
                tester.Dispose(); // Dispose of the tester

                GC.Collect();
            }
            else // Run as a web operation
            {
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
                    Task t = CreateHostBuilder(conformLogger, conformStateManager, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept / 
                         .Build()
                         .RunAsync();

                    t.Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception stating application: {ex.Message}");
                }
                return 0;
            }

            return returnCode; ;
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

    }
}