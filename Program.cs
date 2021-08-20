using CommandLine;
using CommandLine.Text;
using ConformU;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ConformU
{
    public class Program
    {
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
                    (CommandLineOptions options) => Run(options),
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

        static int Run(CommandLineOptions o)
        {
            List<string> argList = new();
            Console.WriteLine($"Starting parsed");

            // Extract the settings file location if provided
            if (!string.IsNullOrEmpty(o.SettingsFileLocation))
            {
                Console.WriteLine($"Settings file location: {o.SettingsFileLocation}");
                argList.Add($"--{ConformConstants.COMMAND_OPTION_SETTINGS}");
                argList.Add(o.SettingsFileLocation);
            }

            // Extract the log file location if provided
            if (!string.IsNullOrEmpty(o.LogFileName))
            {
                Console.WriteLine($"Log file location: {o.LogFileName}");
                argList.Add($"--{ConformConstants.COMMAND_OPTION_LOGFILENAME}");
                argList.Add(o.LogFileName);
            }

            // Extract the log file path if provided
            if (!string.IsNullOrEmpty(o.LogFilePath))
            {
                Console.WriteLine($"Log file path: {o.LogFilePath}");
                argList.Add($"--{ConformConstants.COMMAND_OPTION_LOGFILEPATH}");
                argList.Add(o.LogFilePath);
            }

            // Run from command line if requested
            if (o.Run)
            {
                foreach (string s in argList)
                { Console.WriteLine($"ARG = '{s}'"); }


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

                string logFileName = o.LogFileName ?? "";
                string logFilePath = o.LogFilePath ?? "";

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
                conformLogger.Debug = true;
                ConformConfiguration settings = new(conformLogger, o.SettingsFileLocation);

                // Validate the supplied configuration and only start if there are no settings issues
                string validationMessage = settings.Validate();
                if (!string.IsNullOrEmpty(validationMessage)) // There is a configuration issue so present an error message
                {
                    Console.WriteLine($"Cannot start test:\r\n{validationMessage}");
                    return 99;
                }

                // Setting have validated OK so start the test
                CancellationTokenSource cancellationTokenSource = new(); // Create a task cancellation token source and cancellation token
                CancellationToken cancelConformToken = cancellationTokenSource.Token;

                // Create a test manager instance to oversee the test
                ConformanceTestManager tester = new(settings, conformLogger, cancelConformToken);
                try
                {
                    tester.TestDevice();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"StartTest - Exception: \r\n {ex}");
                }
                tester.Dispose(); // Dispose of the tester
                tester = null;

                GC.Collect();
            }
            else // Run as a web operation
            {
                if (!Debugger.IsAttached)
                {
                    // Set the working directory to the application directory
                    Directory.SetCurrentDirectory(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));
                }
                CancellationTokenSource tokenSource= new();
                CancellationToken applicationCancellationtoken = tokenSource.Token;

                argList.Add("--environment Development");
                
                Console.WriteLine($"Staring web server.");
                Task t = CreateHostBuilder(argList.ToArray()) // Use the revised argument list because the command line parser is fussy about prefixes and won't accept / 
                     .Build()
                     .RunAsync(applicationCancellationtoken);

                if (!Debugger.IsAttached)
                {
                    OpenBrowser("http://localhost:5000/");
                }
                t.Wait();
                return 0;
            }





            return 0;
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)

                .ConfigureLogging(logging =>
                {
                    logging.ClearProviders();
                    logging.AddConsole();
                    logging.AddDebug();
                })
                .ConfigureWebHostDefaults(webBuilder =>
                    {
                        webBuilder.UseStartup<Startup>();
                    });

        public static void OpenBrowser(string url)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                ProcessStartInfo startInfo = new();
                startInfo.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                startInfo.FileName = "xdg-open";
                startInfo.Arguments = url;
                Process.Start(startInfo);
                //Process.Start("xdg-open", url);
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                Process.Start("open", url);
            }
            else
            {
                // throw 
            }
        }
    }
}