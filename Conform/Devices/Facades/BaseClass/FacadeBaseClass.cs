using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using static ConformU.ConformConstants;
using System.Threading;
using System.Collections;
#if WINDOWS7_0_OR_GREATER
using System.Windows.Forms;
#endif

namespace ConformU
{
    public class FacadeBaseClass : IDisposable
    {
        internal dynamic driver; // COM driver object
        internal readonly Settings settings; // Conform configuration settings
        internal readonly ConformLogger logger;
        private bool disposedValue;

#if WINDOWS7_0_OR_GREATER
        internal DriverHostForm driverHostForm;
        Form comHostForm;
#endif

        #region New and Dispose

        public FacadeBaseClass(Settings conformSettings, ConformLogger logger)
        {
            settings = conformSettings;
            this.logger = logger;
            try
            {
#if WINDOWS7_0_OR_GREATER
                logger?.LogMessage("FacadeBaseClass", MessageLevel.msgDebug, $"Using COM host form to create ProgID: {settings.ComDevice.ProgId}");
                logger?.LogMessage("FacadeBaseClass", MessageLevel.msgDebug, $"Creating driver {settings.ComDevice.ProgId} on separate thread. This is thread: {Thread.CurrentThread.ManagedThreadId}");
                Thread driverThread = new(DriverOnSeparateThread);
                driverThread.SetApartmentState(ApartmentState.STA);
                driverThread.DisableComObjectEagerCleanup();
                driverThread.IsBackground = true;
                driverThread.Start(this);
                logger?.LogMessage("FacadeBaseClass", MessageLevel.msgDebug, $"Thread started successfully for {settings.ComDevice.ProgId}. This is thread: {Thread.CurrentThread.ManagedThreadId}");

                do
                {
                    Thread.Sleep(50);
                    Application.DoEvents();
                } while (driverHostForm == null);

                logger?.LogMessage("FacadeBaseClass", MessageLevel.msgDebug, $"Completed create driver delegate for {settings.ComDevice.ProgId} on thread {Thread.CurrentThread.ManagedThreadId}");

#else
                logger?.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Using direct variable to create ProgID: {settings.ComDevice.ProgId}");
                Type driverType = Type.GetTypeFromProgID(settings.ComDevice.ProgId);
                logger?.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating Type: {driverType}");
                driver = Activator.CreateInstance(driverType);
#endif
                logger?.LogMessage("FacadeBaseClass", MessageLevel.msgDebug, $"Initialisation completed OK");
            }
            catch (Exception ex)
            {
                logger.LogMessage("CreateDevice", MessageLevel.msgError, $"Exception creating driver: {ex}");
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    switch (settings.DeviceTechnology)
                    {
                        case DeviceTechnology.Alpaca:
                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.msgDebug, $"About to set Connected False.");
                                driver.Connected = false;
                            }
                            catch { }

                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.msgDebug, $"About to call Dispose method.");
                                driver.Dispose();
                            }
                            catch { }
                            break;

                        case DeviceTechnology.COM:
                            int remainingObjectCount, loopCount;
                            if (driver is not null)
                            {
                                try
                                {
                                    if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.msgDebug, $"About to set Connected False.");
                                    driver.Connected = false;
                                }
                                catch { }

                                try
                                {
                                    if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.msgDebug, $"About to call Dispose method.");
                                    driver.Dispose();
                                }
                                catch { }

                                try
                                {
                                    loopCount = 0;
                                    do
                                    {
                                        loopCount += 1;
                                        remainingObjectCount = Marshal.ReleaseComObject(driver);
                                        if (settings.Debug) logger?.LogMessage("Dispose", MessageLevel.msgDebug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");

                                    }
                                    while (remainingObjectCount > 0 & loopCount <= 20);
                                }
                                catch { }

                                try
                                {
                                    driver = null;
                                    GC.Collect();
                                }
                                catch { }
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

        internal void MethodNoParameters(Action action)
        {
            //logger?.LogMessage("MethodNoParameters", $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
#if WINDOWS7_0_OR_GREATER
            driverHostForm.ActionNoParameters(action);
#else
            action();
#endif
            //logger?.LogMessage("MethodNoParameters", $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
        }

        internal void Method1Parameter(Action<object> action, object parameter1)
        {
            //logger?.LogMessage("Method1Parameter", MessageLevel.msgDebug, $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");

#if WINDOWS7_0_OR_GREATER
            driverHostForm.Action1Parameter(action, parameter1);
#else
            action(parameter1);
#endif
            //logger?.LogMessage("Method1Parameter", MessageLevel.msgDebug, $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
        }

        internal void Method2Parameters(Action<object, object> action, object parameter1, object parameter2)
        {
            //logger?.LogMessage("Method2Parameters", MessageLevel.msgDebug, $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");

#if WINDOWS7_0_OR_GREATER
            driverHostForm.Action2Parameters(action, parameter1, parameter2);
#else
            action(parameter1,parameter2);
#endif
            //logger?.LogMessage("Method2Parameters", MessageLevel.msgDebug, $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
        }

        internal object FunctionNoParameters(Func<object> action)
        {
            //logger?.LogMessage("FunctionNoParameters", MessageLevel.msgDebug, $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");

#if WINDOWS7_0_OR_GREATER
            object returnValue = driverHostForm.FuncNoParameters(action);
#else
            object returnValue = action();
#endif
            //logger?.LogMessage("FunctionNoParameters", MessageLevel.msgDebug, $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
            return returnValue;
        }

        internal object Function1Parameter(Func<object, object> action, object parameter1)
        {
            //logger?.LogMessage("Function1Parameter", MessageLevel.msgDebug, $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");

#if WINDOWS7_0_OR_GREATER
            object returnValue = driverHostForm.Func1Parameter(action, parameter1);
#else
            object returnValue = action(parameter1);
#endif
            //logger?.LogMessage("Function1Parameter", MessageLevel.msgDebug, $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
            return returnValue;
        }

        internal object Function2Parameters(Func<object, object, object> action, object parameter1, object parameter2)
        {
            //logger?.LogMessage("Function2Parameters", MessageLevel.msgDebug, $"About to call driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");

#if WINDOWS7_0_OR_GREATER
            object returnValue = driverHostForm.Func2Parameters(action, parameter1, parameter2);
#else
            object returnValue = action(parameter1,parameter2);
#endif
            //logger?.LogMessage("Function2Parameters", MessageLevel.msgDebug, $"Returned from driverHostForm.SendVoid(action) on thread {Thread.CurrentThread.ManagedThreadId}");
            return returnValue;
        }


#if WINDOWS7_0_OR_GREATER
        /// <summary>
        /// Commands to start the COM driver hosting environment
        /// </summary>
        /// <param name="arg"></param>
        /// <remarks>These commands need to run on a different thread from the calling thread, hence provision in this form that can be provided as a parameter to the Thread.Start command.</remarks>
        private void DriverOnSeparateThread(object arg)
        {

            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"About to create driver host form");
            driverHostForm = new DriverHostForm(settings.ComDevice.ProgId, ref driver, logger); // Create the form
            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"Created driver host form");
            driverHostForm.Show(); // Make it come into existence - it doesn't exist until its shown for some reason
            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"Shown driver host form");
            driverHostForm.Hide(); // Hide the form from view
            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"Hidden driver host form");

            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"Starting driver host environment for {settings.ComDevice.ProgId} on thread {Thread.CurrentThread.ManagedThreadId}");
            Application.Run();  // Start the message loop on this thread to bring the form to life

            logger?.LogMessage("DriverOnSeparateThread", MessageLevel.msgDebug, $"Environment for driver host {settings.ComDevice.ProgId} shut down on thread {Thread.CurrentThread.ManagedThreadId}");
            driverHostForm.Dispose();

        }

#endif


        #region Common Members

        public bool Connected { get => driver.Connected; set => driver.Connected = value; }

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
            return driver.CommandBool(Command, Raw);
        }

        public string CommandString(string Command, bool Raw = false)
        {
            return driver.CommandString(Command, Raw);
        }

        public void SetupDialog()
        {
            MethodNoParameters(() => driver.SetupDialog());
        }

        #endregion

    }
}