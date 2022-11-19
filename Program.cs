using ASCOM.Common;
using ASCOM.Tools;
using CommandLine;
using CommandLine.Text;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static ConformU.Globals;

namespace ConformU
{
    public class Program
    {
        private static ConformConfiguration conformConfiguration;

        public static void Main(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    Console.WriteLine($"CONSOLEARG[{i}] = {args[i]}");
                }
                Console.WriteLine();

                // Parse the command line, options are specified in the CommandLine Options class
                var parser = new CommandLine.Parser(with => with.HelpWriter = null);
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
                h.Heading = $"ASCOM ConforumU {Assembly.GetExecutingAssembly().GetName().Version}"; //change header
                h.Copyright = "Copyright (c) 2021 Peter Simpson\r\n"; //change copyright text
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
            Console.WriteLine($"Starting parsed");

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
            if (commandLineOptions.DebugDiscovery)
            {
                argList.Add($"--{COMMAND_OPTION_DEBUG_DISCOVERY}");
                argList.Add("true");
            }

            // Flag if start-up debug information should be included in the log file
            if (commandLineOptions.DebugStartup)
            {
                argList.Add($"--{COMMAND_OPTION_DEBUG_STARTUP}");
                argList.Add("true");
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

            ConformLogger conformLogger = new(logFileName, logFilePath, loggerName, true);  // Create a logger component
            conformConfiguration = new(conformLogger, commandLineOptions.SettingsFileLocation);

            // Enable logging of Alpaca discovery if a command line option requires this
            conformConfiguration.Settings.TraceDiscovery = commandLineOptions.DebugDiscovery;

            // Set the results filename if supplied on the command line
            conformConfiguration.Settings.ResultsFileName = commandLineOptions.ResultsFileName;

            // Pass in the blazor server connection timeout
            Console.WriteLine($"Server connection timeout: {conformConfiguration.Settings.ConnectionTimeout}");
            argList.Add($"--{COMMAND_OPTION_CONNECTION_TIMEOUT}");
            argList.Add(conformConfiguration.Settings.ConnectionTimeout.ToString());

            #endregion

            // Run from command line if requested
            if (commandLineOptions.Run)
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
                    return 99;
                }

                // Setting have validated OK so start the test
                CancellationTokenSource cancellationTokenSource = new(); // Create a task cancellation token source and cancellation token
                CancellationToken cancelConformToken = cancellationTokenSource.Token;

                // Create a test manager instance to oversee the test
                ConformanceTestManager tester = new(conformConfiguration, conformLogger, cancelConformToken);
                try
                {
                    tester.TestDevice();
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
                // Start a task to check whether any updates are available. Started here to give the maximum time to get a result before the UI is first displayed
                Task.Run(async () =>
                {
                    try
                    {
                        await Update.CheckForUpdates();
                    }
                    catch { } // Ignore exceptions here
                });

                // In production set the working directory to the application directory
                if (!Debugger.IsAttached)
                {
                    Directory.SetCurrentDirectory(Path.GetDirectoryName(Environment.ProcessPath));
                }

                try
                {
                    Console.WriteLine($"Starting web server.");
                    Task t = CreateHostBuilder(conformLogger, conformConfiguration, argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept / 
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

            return 0;
        }

        public static IHostBuilder CreateHostBuilder(ConformLogger conformLogger, ConformConfiguration conformConfiguration, string[] args)
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