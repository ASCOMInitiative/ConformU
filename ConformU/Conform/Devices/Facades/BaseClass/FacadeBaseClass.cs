// Ignore Spelling: Dialog

using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
#if WINDOWS
using System.Windows.Forms;
#endif

namespace ConformU
{
    public class FacadeBaseClass : IAscomDeviceV2, IDisposable
    {
        private const int DRIVER_LOAD_TIMEOUT = 5; // Seconds to wait for the driver to load
        private readonly bool logEnabled = false; // Enable debug logging of this class
        //private readonly bool LOG_DISABLED = false; // Enable debug logging of this class

        internal dynamic Driver; // COM driver object
        internal readonly Settings Settings; // Conform configuration settings
        internal readonly ConformLogger Logger;
        private bool disposedValue;

#if WINDOWS
        internal DriverHostForm DriverHostForm;

        /// <summary>
        /// Commands to start the COM driver hosting environment
        /// </summary>
        /// <param name="arg"></param>
        /// <remarks>These commands need to run on a different thread from the calling thread, hence provision in this form that can be provided as a parameter to the Thread.Start command.</remarks>
        private void DriverOnSeparateThread(object arg)
        {
            // Create the sandbox host form
            if (logEnabled) Logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"About to create driver host form");
            DriverHostForm = new DriverHostForm(Logger)
            {
                ShowInTaskbar = false
            }; // Create the form
            if (logEnabled) Logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"Created driver host form, starting driver host environment for {Settings.ComDevice.ProgId} on thread {Environment.CurrentManagedThreadId}");

            // Start the message loop on this thread to bring the form to life
            Application.Run(DriverHostForm);

            // The form has closed
            if (logEnabled) Logger?.LogMessage("DriverOnSeparateThread", MessageLevel.Debug, $"Environment for driver host {Settings.ComDevice.ProgId} shut down on thread {Environment.CurrentManagedThreadId}");
            DriverHostForm.Dispose();
        }
#endif

        #region New and Dispose

        public FacadeBaseClass(Settings conformSettings, ConformLogger logger)
        {
            Settings = conformSettings;
            this.Logger = logger;
            try
            {
#if WINDOWS
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Using COM host form to create driver : {Settings.ComDevice.ProgId}. This is thread: {Environment.CurrentManagedThreadId}");

                // Create a new thread to host the sandbox form
                Thread driverThread = new(DriverOnSeparateThread);
                driverThread.SetApartmentState(ApartmentState.STA);
                driverThread.IsBackground = true;

                // Start the sandbox form thread
                driverThread.Start(this);
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Thread {driverThread.ManagedThreadId} started successfully for {Settings.ComDevice.ProgId}. This is thread: {Environment.CurrentManagedThreadId}");

                // Wait for the sandbox form to load or timeout
                Stopwatch sw = Stopwatch.StartNew();
                do
                {
                    Thread.Sleep(20);
                    Application.DoEvents();
                } while ((DriverHostForm == null) & (sw.ElapsedMilliseconds < DRIVER_LOAD_TIMEOUT * 1000));
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form creation complete.");

                // Test whether the form loaded OK
                if (DriverHostForm is null) // Form did not create OK in a timely manner
                {
                    throw new Exception($"It took more than {DRIVER_LOAD_TIMEOUT} seconds to create the sandbox Form for driver {Settings.ComDevice.ProgId}.");
                }

                // Wait for the sandbox form to run it's form load event handler
                sw.Reset();
                do
                {
                    Thread.Sleep(20);
                    Application.DoEvents();
                } while ((DriverHostForm.FormInitialised == false) & (sw.ElapsedMilliseconds < DRIVER_LOAD_TIMEOUT * 1000));
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form initialisation complete.");

                // Test whether the form initialised OK
                if (DriverHostForm.FormInitialised == false) // Form did not initialise in a timely manner
                {
                    throw new Exception($"It took more than {DRIVER_LOAD_TIMEOUT} seconds to create the sandbox Form for driver {Settings.ComDevice.ProgId}.");
                }

                // Create the driver on the newly initialised form
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Sandbox form created OK, about to create driver {Settings.ComDevice.ProgId}");
                DriverHostForm.CreateDriver(Settings.ComDevice.ProgId, ref Driver);
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Created driver {Settings.ComDevice.ProgId}");
#endif
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Debug, $"Initialisation completed OK");
            }
            catch (Exception ex)
            {
                if (logEnabled) logger?.LogMessage("FacadeBaseClass", MessageLevel.Error, $"Exception creating driver: {ex}");
                throw;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (logEnabled) Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Entered Dispose class, Disposing: {disposing}, Disposed value: {disposedValue}, Driver is null: {Driver is null}.");
            if (!disposedValue)
            {
                if (disposing)
                {
                    switch (Settings.DeviceTechnology)
                    {
                        case DeviceTechnology.Alpaca:
                            try
                            {
                                if (Settings.DisplayMethodCalls) if (logEnabled) Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                Driver.Connected = false;
                            }
                            catch { }

                            try
                            {
                                if (Settings.DisplayMethodCalls) if (logEnabled) Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                Driver.Dispose();
                            }
                            catch { }
                            break;

                        case DeviceTechnology.COM:
                            int remainingObjectCount, loopCount;
                            if (Driver is not null)
                            {
                                // Set Connected to false
                                try
                                {
                                    //if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                    Method1Parameter((i) => Driver.Connected = i, false);
                                }
                                catch (Exception ex)
                                {
                                    Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception setting Connected false: \r\n{ex}");
                                }

                                // Dispose of the driver
                                try
                                {
                                    //if (settings.DisplayMethodCalls) if (LOG_ENABLED) logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                    MethodNoParameters(() => Driver.Dispose());
                                }
                                catch (Exception ex)
                                {
                                    Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception disposing driver: \r\n{ex}");
                                }

                                // Fully remove the COM object
                                try
                                {
                                    loopCount = 0;
                                    do
                                    {
                                        loopCount += 1;
                                        remainingObjectCount = Marshal.ReleaseComObject(Driver);
                                        Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");
                                    }
                                    while ((remainingObjectCount > 0) & (loopCount <= 20));
                                }
                                catch (Exception ex)
                                {
                                    Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception releasing COM object: \r\n{ex}");
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
                                Driver = null;
                                GC.Collect();
                            }
                            catch (Exception ex)
                            {
                                Logger?.LogMessage("FacadeBaseClass-Dispose", MessageLevel.Debug, $"Exception ending application: \r\n{ex}");
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
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Callable methods

        internal void MethodNoParameters(Action action)
        {
            if (logEnabled) Logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"About to call driverHostForm.ActionNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
#if WINDOWS
            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.MethodNoParameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.Invoke(() => { DriverHostForm.ActionNoParameters(action); });
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.MethodNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.ActionNoParameters(action);
            }
#else
            action();
#endif
            if (logEnabled) Logger?.LogMessage("MethodNoParameters", MessageLevel.Debug, $"Returned from driverHostForm.ActionNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal void Method1Parameter(Action<object> action, object parameter1)
        {
            if (logEnabled) Logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"About to call driverHostForm.Action1Parameter(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Method1Parameter(action) because we are on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.Invoke(() => { DriverHostForm.Action1Parameter(action, parameter1); });
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Method1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.Action1Parameter(action, parameter1);
            }
#else
            action(parameter1);
#endif
            if (logEnabled) Logger?.LogMessage("Method1Parameter", MessageLevel.Debug, $"Returned from driverHostForm.Action1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal void Method2Parameters(Action<object, object> action, object parameter1, object parameter2)
        {
            if (logEnabled) Logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"About to call driverHostForm.Action2Parameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Method2Parameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.Invoke(() => { DriverHostForm.Action2Parameters(action, parameter1, parameter2); });
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Method2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
                DriverHostForm.Action2Parameters(action, parameter1, parameter2);
            }
#else
            action(parameter1, parameter2);
#endif
            if (logEnabled) Logger?.LogMessage("Method2Parameters", MessageLevel.Debug, $"Returned from driverHostForm.Action2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
        }

        internal T FunctionNoParameters<T>(Func<T> action)
        {
            if (logEnabled) Logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"About to call driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            T returnValue;

            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.FuncNoParameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.Invoke(() => DriverHostForm.FuncNoParameters(action));
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.FuncNoParameters(action);
            }
#else
            T returnValue = action();
#endif
            if (logEnabled) Logger?.LogMessage("FunctionNoParameters", MessageLevel.Debug, $"Returned from driverHostForm.FuncNoParameters(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        internal T Function1Parameter<T>(Func<object, T> action, object parameter1)
        {
            if (logEnabled) Logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"About to call driverHostForm.Func1Parameter(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            T returnValue;

            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Function1Parameter(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.Invoke(() => DriverHostForm.Func1Parameter(action, parameter1));
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Function1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.Func1Parameter(action, parameter1);
            }
#else
            T returnValue = action(parameter1);
#endif
            if (logEnabled) Logger?.LogMessage("Function1Parameter", MessageLevel.Debug, $"Returned from driverHostForm.Func1Parameter(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        internal T Function2Parameters<T>(Func<object, object, T> action, object parameter1, object parameter2)
        {
            if (logEnabled) Logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"About to call driverHostForm.Func2Parameters(action) on thread {Environment.CurrentManagedThreadId}");

#if WINDOWS
            T returnValue;

            if (DriverHostForm.InvokeRequired)
            {
                if (logEnabled) Logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"INVOKE REQUIRED to call driverHostForm.Function2Parameters(action) because we are on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.Invoke(() => DriverHostForm.Func2Parameters(action, parameter1, parameter2));
            }
            else
            {
                if (logEnabled) Logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"NO INVOKE REQUIRED to call driverHostForm.Function2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
                returnValue = DriverHostForm.Func2Parameters<T>(action, parameter1, parameter2);
            }
#else
            T returnValue = action(parameter1, parameter2);
#endif
            if (logEnabled) Logger?.LogMessage("Function2Parameters", MessageLevel.Debug, $"Returned from driverHostForm.Func2Parameters(action) on thread {Environment.CurrentManagedThreadId}");
            return returnValue;
        }

        #endregion

        #region IAscomDevice Members

        public bool Connected
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.Connected);
            }
            set
            {
                Method1Parameter((i) => Driver.Connected = i, value);
            }
        }

        public string Description
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.Description);
            }
        }

        public string DriverInfo
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.DriverInfo);
            }
        }

        public string DriverVersion
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.DriverVersion);
            }
        }

        public short InterfaceVersion
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.InterfaceVersion);
            }
        }

        public string Name
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.Name);
            }
        }

        public IList<string> SupportedActions
        {
            get
            {
                return (FunctionNoParameters<IEnumerable>(() => Driver.SupportedActions)).Cast<string>().ToList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            return Function2Parameters<string>((i, j) => Driver.Action(i, j), actionName, actionParameters);
        }

        public void CommandBlind(string command, bool raw = false)
        {
            Method2Parameters((i, j) => Driver.CommandBlind(i, j), command, raw);
        }

        public bool CommandBool(string command, bool raw = false)
        {
            return Function2Parameters<bool>((i, j) => Driver.CommandBool(i, j), command, raw);
        }

        public string CommandString(string command, bool raw = false)
        {
            return Function2Parameters<string>((i, j) => Driver.CommandString(i, j), command, raw);
        }

        public void SetupDialog()
        {
            MethodNoParameters(() => Driver.SetupDialog());
        }

        #endregion

        #region IAscomDeviceV2 members

        public void Connect() => MethodNoParameters(() => Driver.Connect());

        public void Disconnect() => MethodNoParameters(() => Driver.Disconnect());

        public bool Connecting
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.Connecting);
            }
        }

        public List<StateValue> DeviceState
        {
            get
            {
                dynamic deviceStateDynamic = FunctionNoParameters<IEnumerable>(() => Driver.DeviceState);

                List<StateValue> deviceState = new();

                foreach (dynamic device in deviceStateDynamic)
                {
                    deviceState.Add(new StateValue(device.Name, device.Value));
                }

                return deviceState;
            }
        }

        #endregion

    }
}