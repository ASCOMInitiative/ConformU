using System;
using System.Threading.Tasks;
using System.Threading;
using static ConformU.Globals;
using System.Text.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Permissions;

namespace ConformU
{
    public class ConformanceTestManager : IDisposable
    {
        string l_Message;
        readonly ConformConfiguration configuration;
        readonly CancellationToken cancellationToken;
        private bool disposedValue;
        private readonly ConformLogger TL;
        private readonly Settings settings;
        DeviceTesterBaseClass testDevice = null; // Variable to hold the device tester class

        public ConformanceTestManager(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken)
        {
            configuration = conformConfiguration;
            cancellationToken = conformCancellationToken;
            TL = logger;
            settings = conformConfiguration.Settings;
        }

        private void AssignTestDevice()
        {
            testDevice = null;

            // Assign the required device tester class dependent on which type of device is being tested

            switch (settings.DeviceType) // Set current progID and device test class
            {
                case DeviceType.Telescope:
                    {
                        testDevice = new TelescopeTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Dome:
                    {
                        testDevice = new DomeTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Camera:
                    {
                        testDevice = new CameraTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Video:
                    {
                        testDevice = new VideoTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Rotator:
                    {
                        testDevice = new RotatorTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Focuser:
                    {
                        testDevice = new FocuserTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.ObservingConditions:
                    {
                        testDevice = new ObservingConditionsTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.FilterWheel:
                    {
                        testDevice = new FilterWheelTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Switch:
                    {
                        testDevice = new SwitchTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.SafetyMonitor:
                    {
                        testDevice = new SafetyMonitorTester(configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.CoverCalibrator:
                    {
                        testDevice = new CoverCalibratorTester(configuration, TL, cancellationToken);
                        break;
                    }

                default:
                    {
                        TL.LogMessage("Conform:ConformanceCheck", MessageLevel.Error, $"Unknown device type: {settings.DeviceType}. You need to add it to the ConformanceCheck subroutine");
                        throw new ASCOM.InvalidValueException($"Conform:ConformanceCheck - Unknown device type: {settings.DeviceType}. You need to add it to the ConformanceCheck subroutine");
                    }
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
                setupDialogTask.Start();

                // Test whether we are being cancelled
                Task waitForCancelTask = new(() => WaitForCancel(), cancellationToken);
                waitForCancelTask.Start();
                //Wait for either task finish or cancel
                Task.WaitAny(setupDialogTask, waitForCancelTask);

                TL.LogMessage("TestManager:SetupDialog", MessageLevel.Error, $"setup dialog status: {setupDialogTask.Status}, cancel task status: {waitForCancelTask.Status}");
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

        private void WaitForCancel()
        {
            do
            {
                Thread.Sleep(10);
            } while (!cancellationToken.IsCancellationRequested);
        }

        public void TestDevice()
        {
            // Initialise error recorder
            conformResults = new();

            // Assign a tester relevant to the device type being tested
            AssignTestDevice();

            // Test the device
            try
            {
                if (testDevice is null) throw new ASCOM.InvalidOperationException("No test device has been selected.");
                testDevice.CheckInitialise();

                try
                {
                    testDevice.CreateDevice();
                    TL.LogMessage("ConformanceCheck", MessageLevel.OK, "Driver instance created successfully");

                    // Run pre-connect checks if required
                    if (!cancellationToken.IsCancellationRequested & testDevice.HasPreConnectCheck)
                    {
                        TL.LogMessage("", MessageLevel.TestOnly, "");
                        TL.LogMessage("Pre-connect checks", MessageLevel.TestOnly, ""); testDevice.PreConnectChecks();
                        TL.LogMessage("", MessageLevel.TestOnly, "");
                        TL.LogMessage("Connect", MessageLevel.TestOnly, "");
                    }

                    // Try to set Connected to True
                    try
                    {
                        // Test setting Connected to True
                        if (settings.DisplayMethodCalls) TL.LogMessage("ConformanceCheck", MessageLevel.TestAndMessage, "About to set Connected property");
                        testDevice.Connected = true;
                        TL.LogMessage("ConformanceCheck", MessageLevel.OK, "Connected OK");
                        TL.LogMessage("", MessageLevel.TestOnly, "");

                        // Test common methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties)
                        {
                            testDevice.CheckCommonMethods();
                        }

                        // Test and read Can properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasCanProperties)
                        {
                            TL.LogMessage("Can Properties", MessageLevel.TestOnly, ""); testDevice.ReadCanProperties();
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                        }

                        // Carry out pre-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPreRunCheck)
                        {
                            TL.LogMessage("Pre-run Checks", MessageLevel.TestOnly, ""); testDevice.PreRunCheck();
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                        }

                        // Test properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasProperties)
                        {
                            TL.LogMessage("Properties", MessageLevel.TestOnly, ""); testDevice.CheckProperties();
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                        }

                        // Test methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestMethods & testDevice.HasMethods)
                        {
                            TL.LogMessage("Methods", MessageLevel.TestOnly, ""); testDevice.CheckMethods();
                            TL.LogMessage("", MessageLevel.TestOnly, ""); // Blank line
                        }

                        // Test performance
                        if (!cancellationToken.IsCancellationRequested & settings.TestPerformance & testDevice.HasPerformanceCheck)
                        {
                            TL.LogMessage("Performance", MessageLevel.TestOnly, ""); testDevice.CheckPerformance();
                            TL.LogMessage("", MessageLevel.TestOnly, "");
                        }

                        // Carry out post-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPostRunCheck)
                        {
                            TL.LogMessage("Post-run Checks", MessageLevel.TestOnly, ""); testDevice.PostRunCheck();
                            TL.LogMessage("", MessageLevel.TestOnly, ""); // Blank line
                        }

                        // Display completion or "test cancelled" message
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            TL.LogMessage("Conformance test complete", MessageLevel.TestOnly, "");
                        }
                        else
                        {
                            TL.LogMessage("Conformance test interrupted by STOP button or to protect the device.", MessageLevel.TestOnly, "");
                        }
                    }
                    catch (Exception ex) // Exception when setting Connected = True
                    {
                        TL.LogMessage("Connected", MessageLevel.Error, $"Exception when testing driver: {ex.Message}");
                        TL.LogMessage("Connected", MessageLevel.Debug, $"{ex}");
                        TL.LogMessage("", MessageLevel.TestOnly, "");
                        TL.LogMessage("ConformanceCheck", MessageLevel.TestAndMessage, "Further tests abandoned.");
                    }

                }
                catch (Exception ex) // Exception when creating device
                {
                    TL.LogMessage("Initialise", MessageLevel.Error, $"Unable to {(settings.DeviceTechnology == DeviceTechnology.Alpaca ? "access" : "create")} the device: {ex.Message}");
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                    TL.LogMessage("ConformanceCheck", MessageLevel.TestAndMessage, "Further tests abandoned as Conform cannot create the driver");
                }

                // Report the success or failure of conformance checking
                TL.LogMessage("", MessageLevel.TestOnly, "");
                if (conformResults.ErrorCount == 0 & conformResults.IssueCount == 0 & !cancellationToken.IsCancellationRequested) // No issues - device conforms as expected
                {
                    TL.LogMessage("No errors, warnings or issues found: your driver passes ASCOM validation!!", MessageLevel.TestOnly, "");
                    TL.LogMessage("", MessageLevel.TestOnly, "");
                }
                else // Some issues found, the device fails the conformance check
                {
                    l_Message = "Your driver had " + conformResults.ErrorCount + " error";
                    if (conformResults.ErrorCount != 1)
                        l_Message += "s";
                    l_Message = l_Message + " and " + conformResults.IssueCount + " issue";
                    if (conformResults.IssueCount != 1)
                        l_Message += "s";
                    TL.LogMessage(l_Message, MessageLevel.TestOnly, "");
                }

                // List issues and errors
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

                //JsonSerializerOptions options = new();
                //options.WriteIndented = true;
                //string json = JsonSerializer.Serialize<ConformResults>(conformResults, options);
                //TL.LogMessage(json); LogNewLine();

            }
            catch (Exception ex)
            {
                //LogMsgError("Conform:ConformanceCheck Exception: ", ex.ToString());
                TL.LogMessage("ConformanceTestManager", MessageLevel.Error, ex.ToString());
                //OnLogMessageChanged("ConformanceTestManager", $"{DateTime.Now:HH:mm:ss.fff}  ERROR  ConformanceTestManager - {ex}");
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

        }

        #region Dispose Support

        protected virtual void Dispose(bool disposing)
        {
            Console.WriteLine($"ConformanceTestManager.Dispose() {disposing}");
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

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
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
