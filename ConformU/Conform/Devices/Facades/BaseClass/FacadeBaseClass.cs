using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Threading;
#if WINDOWS
using System.Windows.Forms;
#endif

namespace ConformU
{
    public class FacadeBaseClass : IDisposable
    {
        const int DRIVER_LOAD_TIMEOUT = 5; // Seconds to wait for the driver to load
        private readonly bool LOG_ENABLED = false; // Enable debug logging of this class
        //private readonly bool LOG_DISABLED = false; // Enable debug logging of this class

        internal dynamic driver; // COM driver object
        internal readonly Settings settings; // Conform configuration settings
        internal readonly ConformLogger logger;
        private bool disposedValue;

#if WINDOWS
        internal DriverHostForm driverHostForm;

        /// <summary>
        /// Commands to start the COM driver hosting environment
        /// </summary>
        /// <param name="arg"></param>
        /// <remarks>These commands need to run on a different thread from the calling thread, hence provision in this form that can be provided as a parameter to the Thread.Start command.</remarks>
        private void DriverOnSeparateThread(object arg)
        {
            // Create the sandbox host form
            if (LOG_ENABLED) logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"About to create driver host form");
            driverHostForm = new DriverHostForm(logger)
            {
                ShowInTaskbar = false
            }; // Create the form
            if (LOG_ENABLED) logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"Created driver host form, starting driver host environment for {settings.ComDevice.ProgId} on thread {Environment.CurrentManagedThreadId}");

            // Start the message loop on this thread to bring the form to life
            Application.Run(driverHostForm);

            // The form has closed
            if (LOG_ENABLED) logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"Environment for driver host {settings.ComDevice.ProgId} shut down on thread {Environment.CurrentManagedThreadId}");
            driverHostForm.Dispose();
        }
#endif

        #region New and Dispose

        public FacadeBaseClass(Settings conformSettings, ConformLogger logger)
        {
            settings = conformSettings;
            this.logger = logger;
            try
            {
#if WINDOWS
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Using COM host form to create driver : {settings.ComDevice.ProgId}. This is thread: {Environment.CurrentManagedThreadId}");

                // Create a new thread to host the sandbox form
                Thread driverThread = new(DriverOnSeparateThread);
                driverThread.SetApartmentState(ApartmentState.STA);
                driverThread.IsBackground = true;

                // Start the sandbox form thread
                driverThread.Start(this);
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Thread {driverThread.ManagedThreadId} started successfully for {settings.ComDevice.ProgId}. This is thread: {Environment.CurrentManagedThreadId}");

                // Wait for the sandbox form to load or timeout
                Stopwatch sw = Stopwatch.StartNew();
                do
                {
                    Thread.Sleep(20);
                    Application.DoEvents();
                } while ((driverHostForm == null) & (sw.ElapsedMilliseconds < DRIVER_LOAD_TIMEOUT * 1000));
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form creation complete.");

                // Test whether the form loaded OK
                if (driverHostForm is null) // Form did not create OK in a timely manner
                {
                    throw new Exception($"It took more than {DRIVER_LOAD_TIMEOUT} seconds to create the sandbox Form for driver {settings.ComDevice.ProgId}.");
                }

                // Wait for the sandbox form to run it's form load event handler
                sw.Reset();
                do
                {
                    Thread.Sleep(20);
                    Application.DoEvents();
                } while ((driverHostForm.FormInitialised == false) & (sw.ElapsedMilliseconds < DRIVER_LOAD_TIMEOUT * 1000));
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form initialisation complete.");

                // Test whether the form initialised OK
                if (driverHostForm.FormInitialised == false) // Form did not initialise in a timely manner
                {
                    throw new Exception($"It took more than {DRIVER_LOAD_TIMEOUT} seconds to create the sandbox Form for driver {settings.ComDevice.ProgId}.");
                }

                // Create the driver on the newly initialised form
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form created OK, about to create driver {settings.ComDevice.ProgId}");
                driverHostForm.CreateDriver(settings.ComDevice.ProgId, ref driver);
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Created driver {settings.ComDevice.ProgId}");
#endif
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Initialisation completed OK");
            }
            catch (Exception ex)
            {
                if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass", MessageLevel.Error, $"Exception creating driver: {ex}");
                throw;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Entered Dispose class, Disposing: {disposing}, Disposed value: {disposedValue}, Driver is null: {driver is null}.");
            if (!disposedValue)
            {
                if (disposing)
                {
                    switch (settings.DeviceTechnology)
                    {
                        case DeviceTechnology.Alpaca:
                            try
                            {
                                if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                driver.Connected = false;
                            }
                            catch { }

                            try
                            {
                                if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                driver.Dispose();
                            }
                            catch { }
                            break;

                        case DeviceTechnology.COM:
                            int remainingObjectCount, loopCount;
                            if (driver is not null)
                            {
                                // Set Connected to false
                                try
                                {
                                    //if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                    Method1Parameter((i) => driver.Connected = i, false);
                                }
                                catch (Exception ex)
                                {
                                    logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception setting Connected false: \r\n{ex}");
                                }

                                // Dispose of the driver
                                try
                                {
                                    //if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                    MethodNoParameters(() => driver.Dispose());
                                }
                                catch (Exception ex)
                                {
                                    logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception disposing driver: \r\n{ex}");
                                }

                                // Fully remove the COM object
                                try
                                {
                                    loopCount = 0;
                                    do
                                    {
                                        loopCount += 1;
                                        remainingObjectCount = Marshal.ReleaseComObject(driver);
                                        logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");
                                    }
                                    while ((remainingObjectCount > 0) & (loopCount <= 20));
                                }
                                catch (Exception ex)
                                {
                                    logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception releasing COM object: \r\n{ex}");
                                }
                            }

                            // Close the sandbox hosting form
                            try
                            {
#if WINDOWS
                                //if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Ending COM facade application.");
                                Application.Exit();
                                //if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"COM facade application ended.");
#endif
                                driver = null;
                                GC.Collect();
                            }
                            catch (Exception ex)
                            {
                                logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception ending application: \r\n{ex}");
                            }

                            break;
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put clean-up code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
        }

        #endregion

        #region Callable methods

        internal void MethodNoParameters(Action action)
        {
            if (LOG_ENABLED) logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"About to call driverHostForm.ActionNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
#if WINDOWS
            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.MethodNoParameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.Invoke(() => { driverHostForm.ActionNoParameters(action); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.MethodNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.ActionNoParameters(action);
            }
#else
            action();
#endif
            if (LOG_ENABLED) logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"Returned from driverHostForm.ActionNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal void Method1Parameter(Action<object> action, object parameter1)
        {
            if (LOG_ENABLED) logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"About to call driverHostForm.Action1Parameter(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Method1Parameter(action) because we are on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.Invoke(() => { driverHostForm.Action1Parameter(action, parameter1); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Method1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.Action1Parameter(action, parameter1);
            }
#else
            action(parameter1);
#endif
            if (LOG_ENABLED) logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"Returned from driverHostForm.Action1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal void Method2Parameters(Action<object, object> action, object parameter1, object parameter2)
        {
            if (LOG_ENABLED) logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"About to call driverHostForm.Action2Parameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Method2Parameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.Invoke(() => { driverHostForm.Action2Parameters(action, parameter1, parameter2); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Method2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
                driverHostForm.Action2Parameters(action, parameter1, parameter2);
            }
#else
            action(parameter1, parameter2);
#endif
            if (LOG_ENABLED) logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"Returned from driverHostForm.Action2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal object FunctionNoParameters(Func<object> action)
        {
            if (LOG_ENABLED) logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"About to call driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            object returnValue;

            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.FuncNoParameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.Invoke(() => { return driverHostForm.FuncNoParameters(action); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.FuncNoParameters(action);
            }
#else
            object returnValue = action();
#endif
            if (LOG_ENABLED) logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"Returned from driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        internal object Function1Parameter(Func<object, object> action, object parameter1)
        {
            if (LOG_ENABLED) logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"About to call driverHostForm.Func1Parameter(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            object returnValue;

            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Function1Parameter(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.Invoke(() => { return driverHostForm.Func1Parameter(action, parameter1); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Function1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.Func1Parameter(action, parameter1);
            }
#else
            object returnValue = action(parameter1);
#endif
            if (LOG_ENABLED) logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"Returned from driverHostForm.Func1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        internal object Function2Parameters(Func<object, object, object> action, object parameter1, object parameter2)
        {
            if (LOG_ENABLED) logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"About to call driverHostForm.Func2Parameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            object returnValue;

            if (driverHostForm.InvokeRequired)
            {
                if (LOG_ENABLED) logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Function2Parameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.Invoke(() => { return driverHostForm.Func2Parameters(action, parameter1, parameter2); });
            }
            else
            {
                if (LOG_ENABLED) logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Function2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = driverHostForm.Func2Parameters(action, parameter1, parameter2);
            }
#else
            object returnValue = action(parameter1, parameter2);
#endif
            if (LOG_ENABLED) logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"Returned from driverHostForm.Func2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        #endregion

        #region Common Members

        public bool Connected
        {
            get
            {
                return (bool)FunctionNoParameters(() => driver.Connected);
            }
            set
            {
                Method1Parameter((i) => driver.Connected = i, value);
            }
        }

        public string Description => (string)FunctionNoParameters(() => driver.Description);

        public string DriverInfo => (string)FunctionNoParameters(() => driver.DriverInfo);

        public string DriverVersion => (string)FunctionNoParameters(() => driver.DriverVersion);

        public short InterfaceVersion => (short)FunctionNoParameters(() => driver.InterfaceVersion);

        public string Name => (string)FunctionNoParameters(() => driver.Name);

        public IList<string> SupportedActions
        {
            get
            {
                List<string> supportedActions = new();
                foreach (string action in (IEnumerable)FunctionNoParameters(() => driver.SupportedActions))
                {
                    supportedActions.Add(action);
                }
                return supportedActions;
            }
        }

        public string Action(string ActionName, string ActionParameters)
        {
            return (string)Function2Parameters((i, j) => driver.Action(i, j), ActionName, ActionParameters);
        }

        public void CommandBlind(string Command, bool Raw = false)
        {
            Method2Parameters((i, j) => driver.CommandBlind(i, j), Command, Raw);
        }

        public bool CommandBool(string Command, bool Raw = false)
        {
            return (bool)Function2Parameters((i, j) => driver.CommandBool(i, j), Command, Raw);
        }

        public string CommandString(string Command, bool Raw = false)
        {
            return (string)Function2Parameters((i, j) => driver.CommandString(i, j), Command, Raw);
        }

        public void SetupDialog()
        {
            MethodNoParameters(() => driver.SetupDialog());
        }

        #endregion

    }
}