﻿using ASCOM;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using static ConformU.Globals;

namespace ConformU
{
    public class ConformanceTestManager : IDisposable
    {
        private string l_Message;
        private readonly ConformConfiguration configuration;
        private readonly CancellationToken cancellationToken;
        private bool disposedValue;
        private readonly ConformLogger TL;
        private readonly Settings settings;
        private DeviceTesterBaseClass testDevice = null; // Variable to hold the device tester class
        internal static CancellationTokenSource ConformCancellationTokenSource;
        public ConformanceTestManager(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationTokenSource conformCancellationTokenSource, CancellationToken conformCancellationToken)
        {
            configuration = conformConfiguration;
            cancellationToken = conformCancellationToken;
            ConformCancellationTokenSource = conformCancellationTokenSource;
            TL = logger;
            settings = conformConfiguration.Settings;
        }

        private void AssignTestDevice()
        {
            testDevice = null;

            // Assign the required device tester class dependent on which type of device is being tested

            switch (settings.DeviceType) // Set current progID and device test class
            {
                case DeviceTypes.Telescope:
                    testDevice = new TelescopeTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Dome:
                    testDevice = new DomeTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Camera:
                    testDevice = new CameraTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Video:
                    testDevice = new VideoTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Rotator:
                    testDevice = new RotatorTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Focuser:
                    testDevice = new FocuserTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.ObservingConditions:
                    testDevice = new ObservingConditionsTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.FilterWheel:
                    testDevice = new FilterWheelTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.Switch:
                    testDevice = new SwitchTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.SafetyMonitor:
                    testDevice = new SafetyMonitorTester(configuration, TL, cancellationToken);
                    break;

                case DeviceTypes.CoverCalibrator:
                    testDevice = new CoverCalibratorTester(configuration, TL, cancellationToken);
                    break;

                default:
                    TL.LogMessage("Conform:ConformanceCheck", MessageLevel.Error, $"Unknown device type: {settings.DeviceType}. You need to add it to the ConformanceCheck subroutine");
                    throw new InvalidValueException($"Conform:ConformanceCheck - Unknown device type: {settings.DeviceType}. You need to add it to the ConformanceCheck subroutine");
            }
        }

        public void SetupDialog()
        {
            // Assign a tester relevant to the device type being tested
            AssignTestDevice();

            try
            {
                // Create the test device
                testDevice.CreateDevice();

                // Run the SetupDialog method on a separate thread
                Task setupDialogTask = new(() =>
                {
                    testDevice.SetupDialog();
                }, cancellationToken);

                // Test whether we are being cancelled
                Task waitForStopButtonTask = new(() =>
                {
                    do
                    {
                        Thread.Sleep(10);
                    } while (!cancellationToken.IsCancellationRequested);
                    Thread.Sleep(100);
                }, cancellationToken);

                // Start both tasks and wait for either the setup dialogue task or the STOP button task to finish
                setupDialogTask.Start();
                waitForStopButtonTask.Start();
                Task.WaitAny(setupDialogTask, waitForStopButtonTask);

                TL.LogMessage("TestManager:SetupDialog", MessageLevel.Info, $"Setup dialogue task status: {setupDialogTask.Status}, Stop button task status: {waitForStopButtonTask.Status}");

                // Cancel the STOP button task if it is not already in the cancelled state
                if (waitForStopButtonTask.Status == TaskStatus.Running)
                {
                    TL.LogMessage("TestManager:SetupDialog", MessageLevel.Info, $"Cancelling the STOP button task because the setup dialogue method has completed.");
                    ConformCancellationTokenSource.Cancel();
                }

            }
            catch (Exception ex)
            {
                TL.LogMessage("TestManager:SetupDialog", MessageLevel.Error, $"Exception \r\n{ex}");
            }
            finally
            {
                // Always dispose of the device
                try { testDevice.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="numberOfTestCycles"></param>
        /// <returns></returns>
        /// <remarks>
        /// Test process order:
        /// 
        /// InitialiseTest
        /// CreateDevice
        /// PreConnectChecks
        /// Connect
        /// CheckCommonMethods
        /// ReadCanProperties
        /// PreRunCheck
        /// CheckProperties
        /// CheckMethods
        /// CheckPerformance
        /// PostRunCheck
        /// Disconnect
        /// CheckConfiguration
        /// </remarks>
        public int TestDevice(int numberOfTestCycles)
        {
            int returnCode;

            // Start with a blank line to the console log
            Console.WriteLine("");

            // Initialise error recorder
            conformResults = new();

            // Assign a tester relevant to the device type being tested
            AssignTestDevice();

            // Test the device
            try
            {
                if (testDevice is null)
                    throw new ASCOM.InvalidOperationException("No test device has been selected.");

                // Initialise the device
                string testStage = "Initialise";
                testDevice.InitialiseTest();

                try
                {
                    testStage = "CreateDevice";
                    testDevice.CreateDevice();
                    TL.LogMessage(testStage, MessageLevel.OK, "Driver instance created successfully");
                    TL.LogMessage("", MessageLevel.TestOnly, "");

                    // Run pre-connect checks if required
                    if (!cancellationToken.IsCancellationRequested & testDevice.HasPreConnectCheck)
                    {
                        TL.LogMessage("Pre-connect checks", MessageLevel.TestOnly, "");
                        testDevice.PreConnectChecks();
                        TL.LogMessage("", MessageLevel.TestOnly, "");
                    }

                    // Try to set Connected to True
                    try
                    {
                        // Test setting Connected to True
                        TL.LogMessage("Connect to device", MessageLevel.TestOnly, "");
                        testStage = "ConnectToDevice";

                        testDevice.Connect();

                        try
                        {
                            // Initialise the test cycle counter
                            int testCycleCount = 0;

                            // Repeat the main tests until the required number of test cycles has been completed
                            do
                            {
                                // Test common methods
                                if (!cancellationToken.IsCancellationRequested & settings.TestProperties)
                                {
                                    testStage = "CheckCommonMethods";

                                    testDevice.CheckCommonMethods();
                                }

                                // Test and read Can properties
                                if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasCanProperties)
                                {
                                    TL.LogMessage("Can Properties", MessageLevel.TestOnly, "");
                                    testStage = "ReadCanProperties";
                                    testDevice.ReadCanProperties();
                                    TL.LogMessage("", MessageLevel.TestOnly, "");
                                }

                                // Carry out pre-test tasks
                                if (!cancellationToken.IsCancellationRequested & testDevice.HasPreRunCheck)
                                {
                                    TL.LogMessage("Pre-run Checks", MessageLevel.TestOnly, "");
                                    testStage = "PreRunCheck";
                                    testDevice.PreRunCheck();
                                    TL.LogMessage("", MessageLevel.TestOnly, "");
                                }

                                // Test properties
                                if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasProperties)
                                {
                                    TL.LogMessage("Properties", MessageLevel.TestOnly, "");
                                    testStage = "CheckProperties";
                                    testDevice.CheckProperties();
                                    TL.LogMessage("", MessageLevel.TestOnly, "");
                                }

                                // Test methods
                                if (!cancellationToken.IsCancellationRequested & settings.TestMethods & testDevice.HasMethods)
                                {
                                    TL.LogMessage("Methods", MessageLevel.TestOnly, "");
                                    testStage = "CheckMethods";
                                    testDevice.CheckMethods();
                                    TL.LogMessage("", MessageLevel.TestOnly, ""); // Blank line
                                }

                                // Test performance
                                if (!cancellationToken.IsCancellationRequested & settings.TestPerformance & testDevice.HasPerformanceCheck)
                                {
                                    TL.LogMessage("Performance", MessageLevel.TestOnly, "");
                                    testStage = "CheckPerformance";
                                    testDevice.CheckPerformance();
                                    TL.LogMessage("", MessageLevel.TestOnly, "");
                                }

                                // Increment the test cycle counter
                                testCycleCount++;

                                if (cancellationToken.IsCancellationRequested)
                                    break;

                            } while (testCycleCount < numberOfTestCycles); // Exit when the required number of test cycles has been completed.

                            // Carry out post-test tasks
                            if (!cancellationToken.IsCancellationRequested & testDevice.HasPostRunCheck)
                            {
                                TL.LogMessage("Post-run Checks", MessageLevel.TestOnly, "");
                                testStage = "PostRunCheck";
                                testDevice.PostRunCheck();
                                TL.LogMessage("", MessageLevel.TestOnly, ""); // Blank line
                            }

                            // Disconnect from the device
                            try
                            {
                                TL.LogMessage("Disconnect from device", MessageLevel.TestOnly, "");
                                testStage = "Disconnect";

                                testDevice.Disconnect();
                            }
                            catch (Exception ex)
                            {
                                conformResults.Issues.Add(new KeyValuePair<string, string>("Connected", $"Exception when setting Connected to false: {ex.Message}"));
                                TL.LogMessage("Connected", MessageLevel.Issue, $"Exception when setting Connected to False: {ex.Message}");
                                TL.LogMessage("Connected", MessageLevel.Debug, $"{ex}");
                                TL.LogMessage("", MessageLevel.TestOnly, "");
                            }

                            // Carry out check on whether any tests were omitted due to Conform configuration
                            if (!cancellationToken.IsCancellationRequested)
                            {
                                testStage = "CheckConfiguration";
                                testDevice.CheckConfiguration();
                            }

                            // Display completion or "test cancelled" message
                            if (!cancellationToken.IsCancellationRequested | (numberOfTestCycles > 1)) // Test ran to completion or it is a stress test
                            {
                                if (numberOfTestCycles > 1)
                                    TL.LogMessage("Stress test has finished", MessageLevel.TestOnly, "");
                                else
                                    TL.LogMessage("Conformance test has finished", MessageLevel.TestOnly, "");
                            }
                            else // Basic conformance test that was cancelled by the STOP key.
                            {
                                TL.LogMessage("Conformance test interrupted by STOP button or to protect the device.", MessageLevel.TestOnly, "");

                                // Add an issue if the test was interrupted and is therefore incomplete
                                conformResults.Issues.Add(new KeyValuePair<string, string>("StopKey", "The conformance test is incomplete because it was interrupted by the stop key or to protect the device being tested."));
                            }

                        }
                        catch (Exception ex)
                        {
                            conformResults.Issues.Add(new KeyValuePair<string, string>(testStage, $"Exception when testing device - testing abandoned: {ex.Message}"));
                            TL.LogMessage(testStage, MessageLevel.Issue, $"Exception when testing device: {ex.Message}");
                            TL.LogMessage(testStage, MessageLevel.Debug, $"{ex}");
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                            TL.LogMessage(testStage, MessageLevel.TestAndMessage, "Further tests abandoned.");
                        }
                    }
                    catch (MissingMemberException ex) // The Connecting property is missing
                    {
                        testDevice.LogIssue(testStage, $"{ex.Message} Further testing abandoned.");
                        TL.LogMessage(testStage, MessageLevel.Debug, $"{ex}");
                        testDevice.LogInfo(testStage, $"The {settings.DeviceType} device reported interface version {testDevice.GetInterfaceVersion()}, which indicates that it supports the Connect() and Disconnect() methods and the Connecting property.");
                        testDevice.LogInfo(testStage, $"However, the Connecting property is not present in the device interface.");
                        testDevice.LogInfo(testStage, $"Please check whether the device is reporting the correct interface version. For Platform 6 devices the latest {settings.DeviceType} interface version is {DeviceCapabilities.LatestPlatform6Interface[settings.DeviceType.Value]}.");
                        testDevice.LogNewLine();
                        testDevice.LogTestAndMessage(testStage, "Cannot connect to device, further tests abandoned.");
                    }
                    catch (Exception ex) // Exception when connecting to device
                    {
                        testDevice.LogIssue(testStage, $"Connection exception - further testing abandoned: {ex.Message}");
                        TL.LogMessage(testStage, MessageLevel.Debug, $"{ex}");
                        testDevice.LogNewLine();
                        testDevice.LogTestAndMessage(testStage, "Cannot connect to device, further tests abandoned.");
                    }
                }
                catch (InvalidValueException ex) // Interface version is invalid
                {
                    testDevice.LogIssue(testStage, $"The returned interface version is invalid for a {settings.DeviceType} device: {ex.Message}");
                    testDevice.LogInfo(testStage, $"For a Platform 6 interface device the interface version should be: {DeviceCapabilities.LatestPlatform6Interface[settings.DeviceType.Value]}");
                    testDevice.LogInfo(testStage, $"For a Platform 7 interface device that supports Connect(), Disconnect(), Connecting and DeviceState, the interface version should be: {DeviceCapabilities.LatestInterface[settings.DeviceType.Value]}");
                    testDevice.LogNewLine();
                    testDevice.LogTestAndMessage(testStage, "This device is incompatible with ASCOM clients because its interface version is invalid and the device cannot be tested further.");
                }
                catch (Exception ex) // Exception when creating device
                {
                    testDevice.LogIssue("Initialise", $"Unable to {(settings.DeviceTechnology == DeviceTechnology.Alpaca ? "access" : "create")} the device: {ex.Message}");
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                    TL.LogMessage(testStage, MessageLevel.TestAndMessage, "Further tests abandoned as Conform cannot create the driver");
                }

                // Report the success or failure of conformance checking
                TL.LogMessage("", MessageLevel.TestOnly, "");
                if (conformResults.ErrorCount == 0 & conformResults.IssueCount == 0 & conformResults.ConfigurationAlertCount == 0 & !cancellationToken.IsCancellationRequested) // No issues - device conforms as expected
                {
                    TL.LogMessage("Congratulations, no errors, warnings or issues found: your driver passes ASCOM validation!!", MessageLevel.TestOnly, "");
                } // No issues found - success
                else // Some issues found, the device fails the conformance check
                {
                    l_Message = $"Your device had {conformResults.IssueCount} issue{(conformResults.IssueCount == 1 ? "" : "s")}, {conformResults.ErrorCount} error{(conformResults.ErrorCount == 1 ? "" : "s")} and " +
                        $"{conformResults.ConfigurationAlertCount} configuration alert{(conformResults.ConfigurationAlertCount == 1 ? "" : "s")}";

                    TL.LogMessage(l_Message, MessageLevel.TestOnly, "");
                }// Some issues found, the device fails the conformance check

                // List issues, errors and configuration alerts
                if (conformResults.ErrorCount > 0)
                {
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                    TL.LogMessage("Error Summary", MessageLevel.TestOnly, "");
                    foreach (KeyValuePair<string, string> kvp in conformResults.Errors)
                    {
                        TL.LogMessage(kvp.Key, MessageLevel.Error, kvp.Value);
                    }
                }

                if (conformResults.IssueCount > 0)
                {
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                    TL.LogMessage("Issue Summary", MessageLevel.TestOnly, "");
                    foreach (KeyValuePair<string, string> kvp in conformResults.Issues)
                    {
                        TL.LogMessage(kvp.Key, MessageLevel.Issue, kvp.Value);
                    }
                }

                if (conformResults.ConfigurationAlertCount > 0)
                {
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                    TL.LogMessage("Configuration Alert Summary", MessageLevel.TestOnly, "");
                    foreach (KeyValuePair<string, string> kvp in conformResults.ConfigurationAlerts)
                    {
                        TL.LogMessage(kvp.Key, MessageLevel.TestAndMessage, kvp.Value);
                    }
                }

                // Report the timing results if configured to do so
                if (settings.ReportGoodTimings | settings.ReportBadTimings)
                {
                    // Display timing header
                    TL.LogMessage("", $"");
                    TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"See Help for further information.");
                    TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"FAST target response time: {Globals.FAST_TARGET_RESPONSE_TIME:0.0} second{(Globals.FAST_TARGET_RESPONSE_TIME == 1.0 ? "" : "s")}, (configuration and state reporting members).");
                    TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"STANDARD target response time: {Globals.STANDARD_TARGET_RESPONSE_TIME:0.0} second{(Globals.STANDARD_TARGET_RESPONSE_TIME == 1.0 ? "" : "s")}, (property write and asynchronous initiators).");
                    TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"EXTENDED target response time: {Globals.EXTENDED_TARGET_RESPONSE_TIME:0.0} second{(Globals.EXTENDED_TARGET_RESPONSE_TIME == 1.0 ? "" : "s")}, (synchronous methods, ImageArray and ImageArrayVariant).");

                    // Report the timing configuration
                    if (settings.ReportGoodTimings) // Configured to report good outcomes
                    {
                        if (settings.ReportBadTimings) // Configured to report bad outcomes
                            TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"Conform is configured to report both good and bad timing outcomes.");
                        else  // Configured NOT to report bad timing outcomes
                            TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"Conform is configured to report only good timing outcomes and to suppress bad timing outcomes.");
                    }
                    else // Configured NOT to report good timing outcomes
                    {
                        if (settings.ReportBadTimings) // Configured to report bad outcomes
                            TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"Conform is configured to report only bad timing outcomes.");
                        else // Configured NOT to report bad timing outcomes
                            TL.LogMessage("Timing Summary", MessageLevel.TestAndMessage, $"Conform is configured to suppress all timing outcomes.");
                    }

                    // Check whether any timing results were recorded
                    if (conformResults.TimingCount > 0) // Some timings were recorded so report them
                    {
                        // List the timing outcomes
                        TL.LogMessage("", MessageLevel.TestOnly, "");

                        // Provide a wait message if stress testing
                        if (numberOfTestCycles > 1) // Stress test
                        {
                            TL.LogMessage($"Writing {(settings.ReportGoodTimings ? (settings.ReportBadTimings ? "good and bad" : "good") : (settings.ReportBadTimings ? "bad" : "no"))} timings to the log file, please wait...", MessageLevel.TestOnly, "");
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                        }

                        // List the timing results
                        foreach (KeyValuePair<string, string> kvp in conformResults.Timings)
                        {
                            TL.LogMessage(kvp.Key, MessageLevel.TestAndMessage, kvp.Value, logToScreen: numberOfTestCycles == 1);
                        }

                        // Add a new line to the log
                        if (conformResults.Timings.Count > 0)
                            TL.LogMessage($"", MessageLevel.TestOnly, $"", logToScreen: numberOfTestCycles == 1);

                        // Report the overall timing outcome
                        if (conformResults.TimingIssuesCount == 0) // No timing issues - success
                            TL.LogMessage("Congratulations, all members returned within their target response times!!", MessageLevel.TestOnly, "");
                        else // There were timing issues
                            TL.LogMessage($"{conformResults.TimingIssuesCount} member{(conformResults.TimingIssuesCount == 1 ? "" : "s")} " +
                                $"took longer than {(conformResults.TimingIssuesCount == 1 ? "its" : "their")} target response " +
                                $"time{(conformResults.TimingIssuesCount == 1 ? "" : "s")}.", MessageLevel.TestOnly, "", true);
                    }
                    else // No member timings were recorded
                    {
                        TL.LogMessage("", MessageLevel.TestOnly, "", logToScreen: numberOfTestCycles == 1);
                        TL.LogMessage("No member timings were recorded.", MessageLevel.TestOnly, "", true);
                    }
                }

                // Add a blank line to the console output
                Console.WriteLine("");

                try
                {
                    TL.LogMessage("WriteResultsFile", MessageLevel.Debug, $"TraceLogger Log file path: {TL.LogFilePath}, Log file name: {TL.LogFileName}");
                    JsonSerializerOptions options = new()
                    {
                        WriteIndented = true
                    };
                    string json = JsonSerializer.Serialize<ConformResults>(conformResults, options);
                    TL.LogMessage("WriteResultsFile", MessageLevel.Debug, json);

                    // Set the results file filename
                    string reportFileName;
                    if (string.IsNullOrEmpty(settings.ResultsFileName)) // No command line argument has been supplied, so use the log file folder as default
                    {
                        reportFileName = Path.Combine(TL.LogFilePath, "conform.report.txt");
                    }
                    else // A results filename has been supplied on the command line so use it
                    {
                        reportFileName = settings.ResultsFileName;
                    }

                    TL.LogMessage("WriteResultsFile", MessageLevel.Debug, $"Log file path: {TL.LogFilePath}, Report file name: {reportFileName}");
                    File.WriteAllText(reportFileName, json); // Write the file to disk

                    // Create a return code equal to the number of errors + issues
                    returnCode = conformResults.ErrorCount + conformResults.IssueCount + conformResults.ConfigurationAlertCount;

                }
                catch (Exception ex)
                {

                    TL.LogMessage("WriteResultsFile", MessageLevel.Error, ex.ToString());
                    returnCode = -99998;
                }

                TL.SetStatusMessage($"Conformance test has finished.      (Log file: {Path.Combine(TL.LogFilePath, TL.LogFileName)})");
            }
            catch (Exception ex) // An error occurred when initialising the device
            {
                TL.LogMessage("ConformanceTestManager", MessageLevel.Error, $"Exception initialising device: {ex.Message}\r\n{ex}");
                returnCode = -99997;
            }

            // Dispose of the test device
            try
            {
                if (testDevice is not null)
                {
                    testDevice.Dispose();
                    // Clear down the test device and release memory
                    testDevice = null;
                    GC.Collect();
                }
            }
            catch
            {
            }

            TL.LogMessage("TestDevice", MessageLevel.Debug, $"Return code: {returnCode}");
            return returnCode;
        }

        #region Dispose Support

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (testDevice is not null)
                    {
                        testDevice.Dispose();
                        testDevice = null;
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
