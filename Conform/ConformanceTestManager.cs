using System;
using System.Threading.Tasks;
using System.Threading;
using static ConformU.Globals;

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
            // Initialise error counters
            g_CountError = 0;
            g_CountIssue = 0;

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
                    testDevice.LogOK("ConformanceCheck", "Driver instance created successfully");

                    // Run pre-connect checks if required
                    if (!cancellationToken.IsCancellationRequested & testDevice.HasPreConnectCheck)
                    {
                        testDevice.LogNewLine();
                        testDevice.LogTestOnly("Pre-connect checks");                        testDevice.PreConnectChecks();
                        testDevice.LogNewLine();
                        testDevice.LogTestOnly("Connect");                    }

                    // Try to set Connected to True
                    try
                    {
                        // Test setting Connected to True
                        if (settings.DisplayMethodCalls) testDevice.LogComment("ConformanceCheck", "About to set Connected property");
                        testDevice.Connected = true;
                        testDevice.LogOK("ConformanceCheck", "Connected OK");
                        testDevice.LogNewLine();

                        // Test common methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties)
                        {
                            testDevice.CheckCommonMethods();
                        }

                        // Test and read Can properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasCanProperties)
                        {
                            testDevice.LogTestOnly("Can Properties");                            testDevice.ReadCanProperties();
                            testDevice.LogNewLine();
                        }

                        // Carry out pre-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPreRunCheck)
                        {
                            testDevice.LogTestOnly("Pre-run Checks");                            testDevice.PreRunCheck();
                            testDevice.LogNewLine();
                        }

                        // Test properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasProperties)
                        {
                            testDevice.LogTestOnly("Properties");                            testDevice.CheckProperties();
                            testDevice.LogNewLine();
                        }

                        // Test methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestMethods & testDevice.HasMethods)
                        {
                            testDevice.LogTestOnly("Methods");                            testDevice.CheckMethods();
                            testDevice.LogNewLine(); // Blank line
                        }

                        // Test performance
                        if (!cancellationToken.IsCancellationRequested & settings.TestPerformance & testDevice.HasPerformanceCheck)
                        {
                            testDevice.LogTestOnly("Performance");                            testDevice.CheckPerformance();
                            testDevice.LogNewLine();
                        }

                        // Carry out post-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPostRunCheck)
                        {
                            testDevice.LogTestOnly("Post-run Checks");                            testDevice.PostRunCheck();
                            testDevice.LogNewLine(); // Blank line
                        }

                        // Display completion or "test cancelled" message
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            testDevice.LogTestOnly("Conformance test complete");                        }
                        else
                        {
                            testDevice.LogTestOnly("Conformance test interrupted by STOP button or to protect the device.");                        }
                    }
                    catch (Exception ex) // Exception when setting Connected = True
                    {
                        testDevice.LogError("Connected", $"Exception when testing driver: {ex.Message}");
                        testDevice.LogDebug("Connected", $"{ex}");
                        testDevice.LogNewLine();
                        testDevice.LogMsg("ConformanceCheck", MessageLevel.TestAndMessage, "Further tests abandoned.");
                    }

                }
                catch (Exception ex) // Exception when creating device
                {
                    testDevice.LogError("Initialise", $"Unable to {(settings.DeviceTechnology == DeviceTechnology.Alpaca ? "access" : "create")} the device: {ex.Message}");
                    testDevice.LogNewLine();
                    testDevice.LogMsg("ConformanceCheck", MessageLevel.TestAndMessage, "Further tests abandoned as Conform cannot create the driver");
                }

                // Report the success or failure of conformance checking
                testDevice.LogNewLine();
                if (g_CountError == 0 & g_CountIssue == 0 & !cancellationToken.IsCancellationRequested) // No issues - device conforms as expected
                {
                    testDevice.LogTestOnly("No errors, warnings or issues found: your driver passes ASCOM validation!!");                    testDevice.LogNewLine();
                }
                else // Some issues found, the device fails the conformance check
                {
                    l_Message = "Your driver had " + g_CountError + " error";
                    if (g_CountError != 1)
                        l_Message += "s";
                    l_Message = l_Message + " and " + g_CountIssue + " issue";
                    if (g_CountIssue != 1)
                        l_Message += "s";
                    testDevice.LogTestOnly(l_Message);                    testDevice.LogNewLine();
                }
            }
            catch (Exception ex)
            {
                //testDevice.LogMsgError("Conform:ConformanceCheck Exception: ", ex.ToString());
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
