using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using System.Threading.Tasks;
using System.Threading;
using ASCOM;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime;
using System.Text;
using Conform;
using static Conform.Globals;
using static ConformU.ConformConstants;

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

        public ConformanceTestManager(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken)
        {
            configuration = conformConfiguration;
            cancellationToken = conformCancellationToken;
            TL = logger;
            settings = conformConfiguration.Settings;


        }

        public event EventHandler<MessageEventArgs> OutputChanged;
        public event EventHandler<MessageEventArgs> StatusChanged;

        public void TestDevice()
        {
            DeviceTesterBaseClass testDevice = null; // Variable to hold the device tester class
            
            // Initialise error counters
            g_CountError = 0;
            g_CountWarning = 0;
            g_CountIssue = 0;

            // Create the required device tester class
            switch (settings.DeviceType) // Set current progID and device test class
            {
                case DeviceType.Telescope:
                    {
                        testDevice = new TelescopeTester(this, configuration, TL, cancellationToken);
                        break;
                    }

                case DeviceType.Dome:
                    {
                        //testDevice = new DomeTester();
                        break;
                    }

                case DeviceType.Camera:
                    {
                        //testDevice = new CameraTester();
                        break;
                    }

                case DeviceType.Video:
                    {
                        //testDevice = new VideoTester();
                        break;
                    }

                case DeviceType.Rotator:
                    {
                        //testDevice = new RotatorTester();
                        break;
                    }

                case DeviceType.Focuser:
                    {
                        //testDevice = new FocuserTester();
                        break;
                    }

                case DeviceType.ObservingConditions:
                    {
                        //testDevice = new ObservingConditionsTester();
                        break;
                    }

                case DeviceType.FilterWheel:
                    {
                        //testDevice = new FilterWheelTester();
                        break;
                    }

                case DeviceType.Switch:
                    {
                        //testDevice = new SwitchTester();
                        break;
                    }

                case DeviceType.SafetyMonitor:
                    {
                        //testDevice = new SafetyMonitorTester();
                        break;
                    }

                case DeviceType.CoverCalibrator:
                    {
                        //testDevice = new CoverCalibratorTester();
                        break;
                    }

                default:
                    {
                        //LogMsg("Conform:ConformanceCheck", MessageLevel.Error, "Unknown device type: " + m_CurrentDeviceType.ToString() + ". You need to add it to the ConformanceCheck subroutine");
                        break;
                    }
            }

            // Test the device
            try
            {
                if (testDevice is null) throw new ASCOM.InvalidOperationException("No test device has been selected.");
                testDevice.CheckInitialise();

                //l_TestDevice.CheckAccessibility();
                testDevice.LogMsg("", MessageLevel.Always, "");
                try
                {
                    testDevice.CreateDevice();
                    testDevice.LogMsg("ConformanceCheck", MessageLevel.OK, "Driver instance created successfully");

                    // Run pre-connect checks if required
                    if (!cancellationToken.IsCancellationRequested & testDevice.HasPreConnectCheck)
                    {
                        testDevice.LogMsg("", MessageLevel.Always, "");
                        testDevice.LogMsg("Pre-connect checks", MessageLevel.Always, "");
                        testDevice.PreConnectChecks();
                        testDevice.LogMsg("", MessageLevel.Always, "");
                        testDevice.LogMsg("Connect", MessageLevel.Always, "");
                    }

                    // Try to set Connected to True
                    try
                    {
                        // Test setting Connected to True
                        if (settings.DisplayMethodCalls) testDevice.LogMsg("ConformanceCheck", MessageLevel.Comment, "About to set Connected property");
                        testDevice.Connected = true;
                        testDevice.LogMsg("ConformanceCheck", MessageLevel.OK, "Connected OK");
                        testDevice.LogMsg("", MessageLevel.Always, "");

                        // Test common methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties)
                        {
                            testDevice.CheckCommonMethods();
                        }

                        // Test and read Can properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasCanProperties)
                        {
                            testDevice.LogMsg("Can Properties", MessageLevel.Always, "");
                            testDevice.ReadCanProperties();
                            testDevice.LogMsg("", MessageLevel.Always, "");
                        }

                        // Carry out pre-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPreRunCheck)
                        {
                            testDevice.LogMsg("Pre-run Checks", MessageLevel.Always, "");
                            testDevice.PreRunCheck();
                            testDevice.LogMsg("", MessageLevel.Always, "");
                        }

                        // Test properties
                        if (!cancellationToken.IsCancellationRequested & settings.TestProperties & testDevice.HasProperties)
                        {
                            testDevice.LogMsg("Properties", MessageLevel.Always, "");
                            testDevice.CheckProperties();
                            testDevice.LogMsg("", MessageLevel.Always, "");
                        }

                        // Test methods
                        if (!cancellationToken.IsCancellationRequested & settings.TestMethods & testDevice.HasMethods)
                        {
                            testDevice.LogMsg("Methods", MessageLevel.Always, "");
                            testDevice.CheckMethods();
                            testDevice.LogMsg("", MessageLevel.Always, ""); // Blank line
                        }

                        // Test performance
                        if (!cancellationToken.IsCancellationRequested & settings.TestPerformance & testDevice.HasPerformanceCheck)
                        {
                            testDevice.LogMsg("Performance", MessageLevel.Always, "");
                            testDevice.CheckPerformance();
                            testDevice.LogMsg("", MessageLevel.Always, "");
                        }

                        // Carry out post-test tasks
                        if (!cancellationToken.IsCancellationRequested & testDevice.HasPostRunCheck)
                        {
                            testDevice.LogMsg("Post-run Checks", MessageLevel.Always, "");
                            testDevice.PostRunCheck();
                            testDevice.LogMsg("", MessageLevel.Always, ""); // Blank line
                        }

                        // Display completion or "test cancelled" message
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            testDevice.LogMsg("Conformance test complete", MessageLevel.Always, "");
                        }
                        else
                        {
                            testDevice.LogMsg("Conformance test interrupted by STOP button or to protect the device.", MessageLevel.Always, "");
                        }
                    }
                    catch (Exception ex) // Exception when setting Connected = True
                    {
                        testDevice.LogMsg("Connected", MessageLevel.Error, $"Exception when setting Connected = True: {ex.Message}");
                        testDevice.LogMsg("", MessageLevel.Always, "");
                        testDevice.LogMsg("ConformanceCheck", MessageLevel.Always, "Further tests abandoned as Conform cannot connect to the driver");
                    }

                }
                catch (Exception ex) // Exception when creating device
                {
                    testDevice.LogMsg("Initialise", MessageLevel.Error, $"Unable to {(settings.DeviceTechnology == DeviceTechnology.Alpaca? "access" : "create")} the device: {ex.Message}");
                    testDevice.LogMsg("", MessageLevel.Always, "");
                    testDevice.LogMsg("ConformanceCheck", MessageLevel.Always, "Further tests abandoned as Conform cannot create the driver");
                }

                // Report the success or failure of conformance checking
                testDevice.LogMsg("", MessageLevel.Always, "");
                if (g_CountError == 0 & g_CountWarning == 0 & g_CountIssue == 0 & !cancellationToken.IsCancellationRequested) // No issues - device conforms as expected
                {
                    testDevice.LogMsg("No errors, warnings or issues found: your driver passes ASCOM validation!!", MessageLevel.Always, "");
                    testDevice.LogMsg("", MessageLevel.Always, "");
                }
                else // Some issues found, the device fails the conformance check
                {
                    l_Message = "Your driver had " + g_CountError + " error";
                    if (g_CountError != 1)
                        l_Message += "s";
                    l_Message = l_Message + ", " + g_CountWarning + " warning";
                    if (g_CountWarning != 1)
                        l_Message += "s";
                    l_Message = l_Message + " and " + g_CountIssue + " issue";
                    if (g_CountWarning != 1)
                        l_Message += "s";
                    testDevice.LogMsg(l_Message, MessageLevel.Always, "");
                    testDevice.LogMsg("", MessageLevel.Always, "");
                }
            }
            catch (Exception ex)
            {
                //testDevice.LogMsg("Conform:ConformanceCheck Exception: ", MessageLevel.Error, ex.ToString());
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
