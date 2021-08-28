using System;
using System.Threading.Tasks;
using System.Threading;
using static ConformU.Globals;
using System.Text.Json;
using System.Collections.Generic;

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

        public event EventHandler<MessageEventArgs> OutputChanged;
        public event EventHandler<MessageEventArgs> StatusChanged;

        private void AssignTestDevice()
        {
            testDevice = null;

            // Assign the required device tester class dependent on which type of device is being tested

            switch (settings.DeviceType) // Set current progID and device test class
            {
                case DeviceType.Telescope:
                    {
                        testDevice = new TelescopeTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Dome:
                    {
                        testDevice = new DomeTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Camera:
                    {
                        testDevice = new CameraTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Video:
                    {
                        testDevice = new VideoTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Rotator:
                    {
                        testDevice = new RotatorTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Focuser:
                    {
                        testDevice = new FocuserTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.ObservingConditions:
                    {
                        testDevice = new ObservingConditionsTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.FilterWheel:
                    {
                        testDevice = new FilterWheelTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Switch:
                    {
                        testDevice = new SwitchTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.SafetyMonitor:
                    {
                        testDevice = new SafetyMonitorTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.CoverCalibrator:
                    {
                        testDevice = new CoverCalibratorTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                default:
                    {
                        TL.LogMessage("Conform:ConformanceCheck", MessageLevel.Error, "Unknown device type: " + settings.DeviceType + ". You need to add it to the ConformanceCheck subroutine");
                        break;
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
                Task setupDialog = Task.Factory.StartNew(() => testDevice.SetupDialog());

                // Wait for the setup dialogue to be closed
                setupDialog.Wait();
            }
            catch (Exception ex)
            {
                TL.LogMessage("TestManager:SetupDialog", $"Exception \r\n{ex}");
            }
            finally
            {
                // Always dispose of the device
                try { testDevice.Dispose(); } catch { }
            }
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
                    LogOK("ConformanceCheck", "Driver instance created successfully", settings.Debug, this, TL);

                    // Run pre-connect checks if required
                    if (!cancellationToken.IsCancellationRequested & testDevice.HasPreConnectCheck)
                    {
                        LogNewLine(settings.Debug, this, TL);
                        LogTestOnly("Pre-connect checks", settings.Debug, this, TL); testDevice.PreConnectChecks();
                        LogNewLine(settings.Debug, this, TL);
                        LogTestOnly("Connect", settings.Debug, this, TL);
                    }

                    // Try to set Connected to True
                    try
                    {
                        // Test setting Connected to True
                        if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to set Connected property", settings.Debug, this, TL);
                        testDevice.Connected = true;
                        LogOK("ConformanceCheck", "Connected OK", settings.Debug, this, TL);
                        LogNewLine(settings.Debug, this, TL);

                        // Test common methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties)
                        {
                            testDevice.CheckCommonMethods();
                        }

                        // Test and read Can properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasCanProperties)
                        {
                            LogTestOnly("Can Properties", settings.Debug, this, TL); testDevice.ReadCanProperties();
                            LogNewLine(settings.Debug, this, TL);
                        }

                        // Carry out pre-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPreRunCheck)
                        {
                            LogTestOnly("Pre-run Checks", settings.Debug, this, TL); testDevice.PreRunCheck();
                            LogNewLine(settings.Debug, this, TL);
                        }

                        // Test properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasProperties)
                        {
                            LogTestOnly("Properties", settings.Debug, this, TL); testDevice.CheckProperties();
                            LogNewLine(settings.Debug, this, TL);
                        }

                        // Test methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestMethods & testDevice.HasMethods)
                        {
                            LogTestOnly("Methods", settings.Debug, this, TL); testDevice.CheckMethods();
                            LogNewLine(settings.Debug, this, TL); // Blank line
                        }

                        // Test performance
                        if (!cancellationToken.IsCancellationRequested & settings.TestPerformance & testDevice.HasPerformanceCheck)
                        {
                            LogTestOnly("Performance", settings.Debug, this, TL); testDevice.CheckPerformance();
                            LogNewLine(settings.Debug, this, TL);
                        }

                        // Carry out post-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPostRunCheck)
                        {
                            LogTestOnly("Post-run Checks", settings.Debug, this, TL); testDevice.PostRunCheck();
                            LogNewLine(settings.Debug, this, TL); // Blank line
                        }

                        // Display completion or "test cancelled" message
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            LogTestOnly("Conformance test complete", settings.Debug, this, TL);
                        }
                        else
                        {
                            LogTestOnly("Conformance test interrupted by STOP button or to protect the device.", settings.Debug, this, TL);
                        }
                    }
                    catch (Exception ex) // Exception when setting Connected = True
                    {
                        LogError("Connected", $"Exception when testing driver: {ex.Message}", settings.Debug, this, TL);
                        LogDebug("Connected", $"{ex}", settings.Debug, this, TL);
                        LogNewLine(settings.Debug, this, TL);
                        LogTestAndMessage("ConformanceCheck", "Further tests abandoned.", settings.Debug, this, TL);
                    }

                }
                catch (Exception ex) // Exception when creating device
                {
                    LogError("Initialise", $"Unable to {(settings.DeviceTechnology == DeviceTechnology.Alpaca ? "access" : "create")} the device: {ex.Message}", settings.Debug, this, TL);
                    LogNewLine(settings.Debug, this, TL);
                    LogTestAndMessage("ConformanceCheck", "Further tests abandoned as Conform cannot create the driver", settings.Debug, this, TL);
                }

                // Report the success or failure of conformance checking
                LogNewLine(settings.Debug, this, TL);
                if (conformResults.ErrorCount == 0 & conformResults.IssueCount == 0 & !cancellationToken.IsCancellationRequested) // No issues - device conforms as expected
                {
                    LogTestOnly("No errors, warnings or issues found: your driver passes ASCOM validation!!", settings.Debug, this, TL); LogNewLine(settings.Debug, this, TL);
                }
                else // Some issues found, the device fails the conformance check
                {
                    l_Message = "Your driver had " + conformResults.ErrorCount + " error";
                    if (conformResults.ErrorCount != 1)
                        l_Message += "s";
                    l_Message = l_Message + " and " + conformResults.IssueCount + " issue";
                    if (conformResults.IssueCount != 1)
                        l_Message += "s";
                    LogTestOnly(l_Message, settings.Debug, this, TL);
                }

                // List issues and errors
                if (conformResults.ErrorCount > 0)
                {
                    LogNewLine(settings.Debug, this, TL);
                    LogTestOnly("Error Summary", settings.Debug, this, TL);
                    foreach (KeyValuePair<string, string> kvp in conformResults.Errors)
                    {
                        LogMessage(kvp.Key, MessageLevel.Error, kvp.Value, settings.Debug, this, TL);
                    }
                }

                if (conformResults.IssueCount > 0)
                {
                    LogNewLine(settings.Debug, this, TL);
                    LogTestOnly("Issue Summary", settings.Debug, this, TL);
                    foreach (KeyValuePair<string, string> kvp in conformResults.Issues)
                    {
                        LogMessage(kvp.Key, MessageLevel.Issue, kvp.Value, settings.Debug, this, TL);
                    }
                }

                //JsonSerializerOptions options = new();
                //options.WriteIndented = true;
                //string json = JsonSerializer.Serialize<ConformResults>(conformResults, options);
                //LogTestOnly(json); LogNewLine();

            }
            catch (Exception ex)
            {
                //LogMsgError("Conform:ConformanceCheck Exception: ", ex.ToString());
                TL.LogMessage("ConformanceTestManager", ex.ToString());
                OnLogMessageChanged("ConformanceTestManager", $"{DateTime.Now:HH:mm:ss.fff}  ERROR  ConformanceTestManager - {ex}");
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

        internal void OnLogMessageChanged(string id, string message)
        {
            MessageEventArgs e = new()
            {
                Id = id,
                Message = message
            };

            EventHandler<MessageEventArgs> messageEventHandler = OutputChanged;

            if (messageEventHandler is not null)
            {
                messageEventHandler(this, e);
            }
        }

        internal void OnStatusChanged(string status)
        {
            MessageEventArgs e = new()
            {
                Id = "Status",
                Message = status
            };

            EventHandler<MessageEventArgs> messageEventHandler = StatusChanged;

            if (messageEventHandler is not null)
            {
                messageEventHandler(this, e);
            }
        }

        #region Dispose Support

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~DeviceConformanceTester()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
