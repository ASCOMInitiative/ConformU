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
using static Conform.GlobalVarsAndCode;

namespace ConformU
{
    public class DeviceConformanceTester : IDisposable
    {
        ConformConfiguration configuration;
        CancellationToken cancellationToken;
        private bool disposedValue;
        private bool disposedValue1;
        private ConformLogger TL;
        public DeviceConformanceTester(ConformConfiguration conformConfiguration, CancellationToken conforCancellationToken, ConformLogger logger)
        {
            configuration = conformConfiguration;
            cancellationToken = conforCancellationToken;
            TL = logger;
        }

        public event EventHandler<MessageEventArgs> OutputChanged;

        public void TestDevice()
        {
            DeviceTesterBaseClass l_TestDevice = new TelescopeTester(this, TL); // Variable to hold the device being tested
            bool m_TestRunning;

            try
            {
                l_TestDevice.CheckInitialise();
                l_TestDevice.CheckAccessibility();
                l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                try
                {
                    g_Stop = true; // Reset stop flag in case connect fails
                    l_TestDevice.CreateDevice();
                    g_Stop = false; // It worked so allow late steps to run
                    l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgOK, "Driver instance created successfully");
                }
                catch (COMException ex)
                {
                    l_TestDevice.LogMsg("Initialise", MessageLevel.msgError, EX_COM + ex.Message);
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver");
                }
                catch (PropertyNotImplementedException ex)
                {
                    l_TestDevice.LogMsg("Initialise", MessageLevel.msgError, NOT_IMP_NET + ex.Message);
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver");
                }
                catch (DriverException ex)
                {
                    l_TestDevice.LogMsg("Initialise", MessageLevel.msgError, EX_DRV_NET + ex.Message);
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver");
                }
                catch (Exception ex)
                {
                    l_TestDevice.LogMsg("Initialise", MessageLevel.msgError, EX_NET + ex.Message);
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot create the driver");
                }

                if (!TestStop() & l_TestDevice.HasPreConnectCheck)
                {
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("Pre-connect checks", MessageLevel.msgAlways, "");
                    l_TestDevice.PreConnectChecks();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    l_TestDevice.LogMsg("Connect", MessageLevel.msgAlways, "");
                }

                if (!TestStop()) // Only connect if we successfully created the device
                {
                    try
                    {
                        g_Stop = true; // Reset stop flag in case connect fails
                        if (g_Settings.DisplayMethodCalls)
                            l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Connected property");
                        l_TestDevice.Connected = true;
                        g_Stop = false; // It worked so allow late steps to run
                        l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgOK, "Connected OK");
                        l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                    }
                    catch (COMException ex)
                    {
                        l_TestDevice.LogMsg("Connected", MessageLevel.msgError, EX_COM + ex.Message);
                        l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                        l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver");
                    }
                    catch (PropertyNotImplementedException ex)
                    {
                        l_TestDevice.LogMsg("Connected", MessageLevel.msgError, NOT_IMP_NET + ex.Message);
                        l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                        l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver");
                    }
                    catch (DriverException ex)
                    {
                        l_TestDevice.LogMsg("Connected", MessageLevel.msgError, EX_DRV_NET + ex.Message);
                        l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                        l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver");
                    }
                    catch (Exception ex)
                    {
                        l_TestDevice.LogMsg("Connected", MessageLevel.msgError, EX_NET + ex.Message);
                        l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                        l_TestDevice.LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Further tests abandoned as Conform cannot connect to the driver");
                    }
                }

                // Run tests
                if (!TestStop() & g_Settings.TestProperties)
                {
                    l_TestDevice.CheckCommonMethods();
                }

                if (!TestStop() & g_Settings.TestProperties & l_TestDevice.HasCanProperties)
                {
                    l_TestDevice.LogMsg("Can Properties", MessageLevel.msgAlways, "");
                    l_TestDevice.ReadCanProperties();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                }

                if (!TestStop() & l_TestDevice.HasPreRunCheck)
                {
                    l_TestDevice.LogMsg("Pre-run Checks", MessageLevel.msgAlways, "");
                    l_TestDevice.PreRunCheck();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                }

                if (!TestStop() & g_Settings.TestProperties & l_TestDevice.HasProperties)
                {
                    l_TestDevice.LogMsg("Properties", MessageLevel.msgAlways, "");
                    l_TestDevice.CheckProperties();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                }
                if (!TestStop() & g_Settings.TestMethods & l_TestDevice.HasMethods)
                {
                    l_TestDevice.LogMsg("Methods", MessageLevel.msgAlways, "");
                    l_TestDevice.CheckMethods();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, ""); // Blank line
                }

                if (!TestStop() & g_Settings.TestPerformance & l_TestDevice.HasPerformanceCheck)
                {
                    l_TestDevice.LogMsg("Performance", MessageLevel.msgAlways, "");
                    l_TestDevice.CheckPerformance();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, "");
                }

                if (!TestStop() & l_TestDevice.HasPostRunCheck)
                {
                    l_TestDevice.LogMsg("Post-run Checks", MessageLevel.msgAlways, "");
                    l_TestDevice.PostRunCheck();
                    l_TestDevice.LogMsg("", MessageLevel.msgAlways, ""); // Blank line
                }

                if (!TestStop())
                {
                    l_TestDevice.LogMsg("Conformance test complete", MessageLevel.msgAlways, "");
                }
                else
                {
                    l_TestDevice.LogMsg("Conformance test interrupted by STOP button or to protect the device.", MessageLevel.msgAlways, "");
                }

                l_TestDevice.Dispose();
                l_TestDevice = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                l_TestDevice.LogMsg("Conform:ConformanceCheck Exception: ", MessageLevel.msgError, ex.ToString());
            }

            // Extra check just to make sure the device has been disposed if there was an unexpected exception above!
            if (l_TestDevice is object)
            {
                try
                {
                    l_TestDevice.Dispose();
                }
                catch
                {
                }

                l_TestDevice = null;
            }

            m_TestRunning = false;



            // Clear down the test device and release memory
            l_TestDevice = null;
            GC.Collect();
            OnLogMessageChanged("TestDevice", $"Finished processing: Task Cancelled: {cancellationToken.IsCancellationRequested}");
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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue1)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue1 = true;
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
    }


}
