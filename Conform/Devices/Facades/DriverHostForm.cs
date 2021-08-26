using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

#if WINDOWS7_0_OR_GREATER
using System.Windows.Forms;

namespace ConformU
{

    public partial class DriverHostForm : Form
    {
        private string progId;
        private ConformLogger logger;
        private dynamic driver;

        public DriverHostForm(string progId, ref dynamic driver, ConformLogger logger)
        {
            InitializeComponent();
            this.progId = progId;
            this.logger = logger;

            this.FormClosed += DriverHostForm_FormClosed;
            this.Load += DriverHostForm_Load;
            logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Form has been instantiated on thread: {Thread.CurrentThread.ManagedThreadId}");

            logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Using direct variable to create ProgID: {progId}");
            Type driverType = Type.GetTypeFromProgID(progId);
            logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Creating Type: {driverType}");
            driver = Activator.CreateInstance(driverType);
            logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Driver {progId} has been created on thread: {Thread.CurrentThread.ManagedThreadId}");
            this.driver = driver;
            logger?.LogMessage("DriverHostForm", MessageLevel.Debug, $"Driver is not null: {driver is not null}");
        }

        /// <summary>
        /// Form load event - Create the driver
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DriverHostForm_Load(object sender, EventArgs e)
        {

            logger?.LogMessage("DriverHostForm_Load", MessageLevel.Debug, $"Form load event has been called on thread: {Thread.CurrentThread.ManagedThreadId}");

        }

        /// <summary>
        /// When the form is closing stop the windows message loop on this thread so that the thread will end
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DriverHostForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.ExitThread();
        }

        public delegate void ActionNoParametersDelegate(Action action);
        public delegate void Action1ParameterDelegate(Action<object> action, object parameter1);
        public delegate void Action2ParametersDelegate(Action<object, object> action, object parameter1, object parameter2);
        public delegate object FuncNoParameterDelegate(Func<object> action);
        public delegate object Func1ParameterDelegate(Func<object, object> func, object parameter1);
        public delegate object Func2ParameterDelegate(Func<object, object, object> func, object parameter1, object parameter2);

        internal void ActionNoParameters(Action action)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("ActionNoParameters", $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                ActionNoParametersDelegate sendVoid = new(ActionNoParameters);
                this.Invoke(sendVoid, action);
            }
            else
            {
                //logger?.LogMessage("ActionNoParameters", $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                action();
                //logger?.LogMessage("ActionNoParameters",  $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
            }

        }

        internal void Action1Parameter(Action<object> action, object parameter1)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                Action1ParameterDelegate sendVoid = new(Action1Parameter);
                this.Invoke(sendVoid, action, parameter1);
            }
            else
            {
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                action(parameter1);
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
            }

        }

        internal void Action2Parameters(Action<object, object> action, object parameter1, object parameter2)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                Action2ParametersDelegate sendVoid = new(Action2Parameters);
                this.Invoke(sendVoid, action, parameter1, parameter2);
            }
            else
            {
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                action(parameter1, parameter2);
                //logger?.LogMessage("Action2Parameters", MessageLevel.msgDebug, $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
            }

        }

        internal object FuncNoParameters(Func<object> action)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                FuncNoParameterDelegate getValue = new(FuncNoParameters);
                return this.Invoke(getValue, action);
            }
            else
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                object returnValue = action();
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                return returnValue;
            }
        }

        internal object Func1Parameter(Func<object, object> action, object parameter1)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                Func1ParameterDelegate getValue = new(Func1Parameter);
                return this.Invoke(getValue, action, parameter1);
            }
            else
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                object returnValue = action(parameter1);
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                return returnValue;
            }
        }

        internal object Func2Parameters(Func<object, object, object> action, object parameter1, object parameter2)
        {
            if (this.InvokeRequired)
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Invoke required for {action.Method.Name} command on thread: {Thread.CurrentThread.ManagedThreadId}");
                Func2ParameterDelegate getValue = new(Func2Parameters);
                return this.Invoke(getValue, action, parameter1, parameter2);
            }
            else
            {
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"About to run Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                object returnValue = action(parameter1, parameter2);
                //logger?.LogMessage("FuncNoParameters", MessageLevel.msgDebug, $"Returned from Action {action.Method.Name} on thread: {Thread.CurrentThread.ManagedThreadId}");
                return returnValue;
            }
        }
    }
}
#endif
