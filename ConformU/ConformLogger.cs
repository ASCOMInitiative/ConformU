using ASCOM.Common.Interfaces;
using ASCOM.Tools;
using System;
using static ConformU.Globals;

namespace ConformU
{
    public class ConformLogger : TraceLogger, ITraceLogger, IDisposable
    {
        private bool debug;
        public ConformLogger(string logFileName, string logFilePath, string loggerName, bool enabled) : base(logFileName, logFilePath, loggerName, enabled)
        {
            base.IdentifierWidth = TEST_NAME_WIDTH;
        }

        /// <summary>
        ///  Event fired when the message log changes.
        /// </summary>
        public event EventHandler<MessageEventArgs> MessageLogChanged;

        /// <summary>
        /// Event fired when the status message changes.
        /// </summary>
        public event EventHandler<MessageEventArgs> StatusChanged;

        /// <summary>
        /// Flag indicating whether debug messages should be included in the log.
        /// </summary>
        public bool Debug
        {
            get
            {
                return debug;
            }
            set
            {
                debug = value;
                if (value)
                {
                    base.SetMinimumLoggingLevel(ASCOM.Common.Interfaces.LogLevel.Debug);
                }
                else
                {
                    base.SetMinimumLoggingLevel(ASCOM.Common.Interfaces.LogLevel.Information);
                }
            }
        }

        /// <summary>
        /// Log a message on the screen, console and log file
        /// </summary>
        /// <param name="id"></param>
        /// <param name="logLevel"></param>
        /// <param name="message"></param>
        public void LogMessage(string id, MessageLevel logLevel, string message)
        {
            string screenMessage, messageLevel;

            // Ignore debug messages when not in Debug mode
            if ((logLevel == MessageLevel.Debug) & !Debug) return;


            // Format the message level string
            switch (logLevel)
            {
                case MessageLevel.Debug:
                case MessageLevel.Info:
                case MessageLevel.OK:
                case MessageLevel.Issue:
                case MessageLevel.Error:
                    {
                        messageLevel = logLevel.ToString().ToUpperInvariant().PadRight(MESSAGE_LEVEL_WIDTH);
                        break;
                    }

                case MessageLevel.TestOnly:
                case MessageLevel.TestAndMessage:
                    {
                        messageLevel = "".PadRight(MESSAGE_LEVEL_WIDTH);
                        break;
                    }
                default:
                    {
                        throw new ASCOM.InvalidValueException($"LogMsg - Unknown message level: {logLevel}.");
                    }
            }

            // Format the screen message to be consistent with the messages in the log file 
            switch (logLevel)
            {
                case MessageLevel.TestOnly:
                    {
                        screenMessage = id;
                        break;
                    }
                default:
                    {
                        screenMessage = $"{id,-TEST_NAME_WIDTH} {messageLevel} {message}";
                        break;
                    }
            }

            // Write the message to the console
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {id,-TEST_NAME_WIDTH} {messageLevel,MESSAGE_LEVEL_WIDTH} {message}");

            // Write the message to the log file
            base.LogMessage(id, $"{messageLevel,MESSAGE_LEVEL_WIDTH} {message}");

            // Raise the MessaegLogChanged event to Write the message to the screen
            OnMessageLogChanged(screenMessage);
        }

        private void OnMessageLogChanged(string message)
        {
            MessageEventArgs eventArgs = new()
            {
                Message = $"{DateTime.Now:HH:mm:ss.fff} {message}"
            };

            MessageLogChanged?.Invoke(this, eventArgs);
        }

        /// <summary>
        /// Raises an event notifying that the status message has changed
        /// </summary>
        /// <param name="status">new status message.</param>
        /// <remarks>
        /// This is part of ConformLogger for convenience because the logger is used everywhere in the application and already supports the LogMessageChanged event.
        /// </remarks>
        public void SetStatusMessage(string status)
        {
            MessageEventArgs eventArgs = new()
            {
                Message = status
            };

            EventHandler<MessageEventArgs> messageEventHandler = StatusChanged;

            if (messageEventHandler is not null)
            {
                messageEventHandler(this, eventArgs);
            }
        }

        /// <summary>
        /// Override this method to force consistent use of the LogMessage method.
        /// </summary>
        /// <param name="logLevel"></param>
        /// <param name="message"></param>
        public new void Log(ASCOM.Common.Interfaces.LogLevel logLevel, string message)
        {

            if (logLevel >= base.LoggingLevel)
            {
                // Write the message to the console
                Console.WriteLine($"{logLevel,-TEST_NAME_WIDTH} {message}");

                // Write the message to the log file
                base.Log(logLevel, $"{string.Empty,-MESSAGE_LEVEL_WIDTH} {message}");

                // Raise the MessaegLogChanged event to Write the message to the screen
                OnMessageLogChanged($"{logLevel,-TEST_NAME_WIDTH - MESSAGE_LEVEL_WIDTH} {message}");
            }
        }

        public new void LogMessage(string method, string message)
        {
            // Write the message to the console
            Console.WriteLine($"{method}{(string.IsNullOrEmpty(method)?"":" ")}{message}");

            // Write the message to the log file
            base.LogMessage(method,message);

            // Raise the MessaegLogChanged event to Write the message to the screen
            OnMessageLogChanged($"{method} {message}");

        }


    }
}