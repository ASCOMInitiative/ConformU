﻿#if WINDOWS
using System.Windows.Forms;
#endif

using System;
using System.Runtime.InteropServices;

#if WINDOWS
namespace ConformU
{
    public partial class DriverHostForm : Form
    {
        private readonly bool logEnabled = false; // Enable debug logging of this class
        private readonly bool logDisabled = false; // Enable debug logging of this class

        private string progId;
        private readonly ConformLogger logger;
        private dynamic driver;

        private readonly object driverLock = new();

        public bool FormInitialised = false;

        public DriverHostForm(ConformLogger logger)
        {
            InitializeComponent();
            this.logger = logger;

            this.FormClosed += DriverHostForm_FormClosed;
            this.Load += DriverHostForm_Load;
            if (logDisabled) logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Instantiated on thread: {Environment.CurrentManagedThreadId}, Message loop is running: {Application.MessageLoop}");
        }

        /// <summary>
        /// Form load event - Create the driver
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DriverHostForm_Load(object sender, EventArgs e)
        {
            if (logDisabled) logger?.LogMessage("DriverHostForm_Load", MessageLevel.Debug, $"Form load event has been called on thread: {Environment.CurrentManagedThreadId}. Message loop is running: {Application.MessageLoop}");
            this.Show();
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
            FormInitialised = true;
        }

        /// <summary>
        /// When the form is closing stop the windows message loop on this thread so that the thread will end
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DriverHostForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            int remainingObjectCount, loopCount;
            if (driver is not null)
            {
                try
                {
                    //if (LOG_DISABLED) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"About to set Connected False.");
                    driver.Connected = false;
                }
                catch (Exception ex)
                {
                    if (logDisabled) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Exception setting Connected false: \r\n{ex}");
                }

                try
                {
                    //if (LOG_DISABLED) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                    //driver.Dispose();
                }
                catch (Exception ex)
                {
                    if (logDisabled) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Exception disposing driver: \r\n{ex}");
                }

                try
                {
                    loopCount = 0;
                    do
                    {
                        loopCount += 1;
                        remainingObjectCount = Marshal.ReleaseComObject(driver);
                        //if (LOG_DISABLED) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");

                    }
                    while ((remainingObjectCount > 0) & (loopCount <= 20));
                }
                catch (Exception ex)
                {
                    if (logDisabled) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Exception releaseing COM object: \r\n{ex}");
                }

                try
                {
#if WINDOWS
                    //if (LOG_DISABLED) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Ending COM facade application.");
                    Application.Exit();
                    //if (LOG_DISABLED) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"COM facade application ended.");
#endif
                    driver = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    if (logDisabled) logger?.LogMessage("DriverHostForm-Dispose", MessageLevel.Debug, $"Exception ending application: \r\n{ex}");
                }
            }

            Application.ExitThread();
        }

        public delegate void ActionNoParametersDelegate(Action action);
        public delegate void Action1ParameterDelegate(Action<object> action, object parameter1);
        public delegate void Action2ParametersDelegate(Action<object, object> action, object parameter1, object parameter2);
        public delegate object FuncNoParameterDelegate(Func<object> func);
        public delegate object Func1ParameterDelegate(Func<object, object> func, object parameter1);
        public delegate object Func2ParameterDelegate(Func<object, object, object> func, object parameter1, object parameter2);

        internal void CreateDriver(string progId, ref dynamic driver)
        {
            this.progId = progId;

            if (this.InvokeRequired)
            {
                if (logDisabled) logger?.LogMessage("HostForm.CreateDriver", MessageLevel.Debug, $"INVOKE REQUIRED - Message loop is running: {Application.MessageLoop}");
                this.Invoke(() => { CreateDriverInvoked(); });
            }
            else
            {
                if (logDisabled) logger?.LogMessage("HostForm.CreateDriver", MessageLevel.Debug, $"NO INVOKE REQUIRED - Message loop is running: {Application.MessageLoop}");
                CreateDriverInvoked();
            }
            driver = this.driver;
        }

        private void CreateDriverInvoked()
        {
            this.Hide();
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Running on thread: {Environment.CurrentManagedThreadId}");
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Getting type for {progId}");
            Type driverType = Type.GetTypeFromProgID(progId);
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Creating Type: {driverType}");
            driver = Activator.CreateInstance(driverType);
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Driver {progId} has been created on thread: {Environment.CurrentManagedThreadId}");
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Driver is not null: {driver is not null}");
            if (logDisabled) logger?.LogMessage("HostForm.CreateDriverInvoked", MessageLevel.Debug, $"Message loop is running: {Application.MessageLoop}");
        }

        internal void ActionNoParameters(Action action)
        {
            if (this.InvokeRequired)
            {
                if (logDisabled) logger?.LogMessage("HostForm.ActionNoParameters", MessageLevel.Debug, $"Invoke required for {action.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                ActionNoParametersDelegate sendVoid = new(ActionNoParameters);
                this.Invoke(sendVoid, action);
                if (logDisabled) logger?.LogMessage("HostForm.ActionNoParameters", MessageLevel.Debug, $"Returned from invoking Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
            else
            {
                if (logDisabled) logger?.LogMessage("HostForm.ActionNoParameters", MessageLevel.Debug, $"About to run Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                action();
                if (logDisabled) logger?.LogMessage("HostForm.ActionNoParameters", MessageLevel.Debug, $"Returned from Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
        }

        internal void Action1Parameter(Action<object> action, object parameter1)
        {
            if (this.InvokeRequired)
            {
                if (logDisabled) logger?.LogMessage("HostForm.Action1Parameter", MessageLevel.Debug, $"Invoke required for {action.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                Action1ParameterDelegate sendVoid = new(Action1Parameter);
                this.Invoke(sendVoid, action, parameter1);
                if (logDisabled) logger?.LogMessage("HostForm.Action1Parameter", MessageLevel.Debug, $"Returned from invoking Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
            else
            {
                if (logDisabled) logger?.LogMessage("HostForm.Action1Parameter", MessageLevel.Debug, $"About to run Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                action(parameter1);
                if (logDisabled) logger?.LogMessage("HostForm.Action1Parameter", MessageLevel.Debug, $"Returned from Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
        }

        internal void Action2Parameters(Action<object, object> action, object parameter1, object parameter2)
        {
            if (this.InvokeRequired)
            {
                if (logEnabled) logger?.LogMessage("HostForm.Action2Parameters", MessageLevel.Debug, $"Invoke required for {action.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                Action2ParametersDelegate sendVoid = new(Action2Parameters);
                this.Invoke(sendVoid, action, parameter1, parameter2);
                if (logEnabled) logger?.LogMessage("HostForm.Action2Parameters", MessageLevel.Debug, $"Returned from invoking Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
            else
            {
                if (logEnabled) logger?.LogMessage("HostForm.Action2Parameters", MessageLevel.Debug, $"About to run Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                action(parameter1, parameter2);
                if (logEnabled) logger?.LogMessage("HostForm.Action2Parameters", MessageLevel.Debug, $"Returned from Action {action.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
            }
        }

        internal T FuncNoParameters<T>(Func<T> func)
        {
                if (logEnabled) logger?.LogMessage("HostForm.FuncNoParameters", MessageLevel.Debug, $"Method called on thread {Environment.CurrentManagedThreadId}, Message loop running: {Application.MessageLoop}");
                if (this.InvokeRequired)
                {
                    if (logEnabled) logger?.LogMessage("HostForm.FuncNoParameters", MessageLevel.Debug, $"Invoke required for {func.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                return this.Invoke(func);
                }
                else
                {
                    if (logEnabled) logger?.LogMessage("HostForm.FuncNoParameters", MessageLevel.Debug, $"About to run Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                T returnValue = func();
                    if (logEnabled) logger?.LogMessage("HostForm.FuncNoParameters", MessageLevel.Debug, $"Returned from Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                    return returnValue;
                }
        }

        internal T Func1Parameter<T>(Func<object, T> func, object parameter1)
        {
            if (this.InvokeRequired)
            {
                if (logEnabled) logger?.LogMessage("HostForm.Func1Parameter", MessageLevel.Debug, $"Invoke required for {func.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                return this.Invoke(() => func(parameter1));
            }
            else
            {
                if (logEnabled) logger?.LogMessage("HostForm.Func1Parameter", MessageLevel.Debug, $"About to run Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                T returnValue = func(parameter1);
                if (logEnabled) logger?.LogMessage("HostForm.Func1Parameter", MessageLevel.Debug, $"Returned from Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                return returnValue;
            }
        }

        internal T Func2Parameters<T>(Func<object, object, T> func, object parameter1, object parameter2)
        {
            if (this.InvokeRequired)
            {
                if (logEnabled) logger?.LogMessage("HostForm.Func2Parameters", MessageLevel.Debug, $"Invoke required for {func.Method.Name} command on thread: {Environment.CurrentManagedThreadId}");
                return this.Invoke(() => func(parameter1, parameter2));
            }
            else
            {
                if (logEnabled) logger?.LogMessage("HostForm.Func2Parameters", MessageLevel.Debug, $"About to run Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                T returnValue = func(parameter1, parameter2);
                if (logEnabled) logger?.LogMessage("HostForm.Func2Parameters", MessageLevel.Debug, $"Returned from Action {func.Method.Name} on thread: {Environment.CurrentManagedThreadId}");
                return returnValue;
            }
        }
    }
}
#endif
