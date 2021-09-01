using System;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.Utilities;
using static ConformU.Globals;

namespace ConformU
{
    public class ConformLogger : TraceLogger, ILogger
    {
        public ConformLogger(string logFileName, string logFilePath, string loggerName, bool enabled) : base(logFileName, logFilePath, loggerName, enabled)
        {
            Console.WriteLine($"ConformLogger - Creating log file - Name: {logFileName}, Path: {logFilePath}, Type: {loggerName}");
            base.IdentifierWidth = TEST_NAME_WIDTH;
        }

        public event EventHandler<MessageEventArgs> OutputChanged;
        public event EventHandler<MessageEventArgs> StatusChanged;

        public bool Debug { get; set; }

        public void LogMessage(string id, MessageLevel logLevel, string message)
        {
            string screenMessage,messageLevelFormatted;

            // Ignore debug messages when not in Debug mode
            if ((logLevel == MessageLevel.Debug) & !Debug) return;

            switch (logLevel)
            {
                case MessageLevel.Debug:
                case MessageLevel.Info:
                case MessageLevel.OK:
                case MessageLevel.Issue:
                case MessageLevel.Error:
                    {
                        messageLevelFormatted = logLevel.ToString().ToUpperInvariant();
                        break;
                    }

                case MessageLevel.TestOnly:
                case MessageLevel.TestAndMessage:
                    {
                        messageLevelFormatted = "";
                        break;
                    }
                default:
                    {
                        throw new ASCOM.InvalidValueException($"LogMsg - Unknown message level: {logLevel}.");
                    }
            }
            messageLevelFormatted = messageLevelFormatted.PadRight(MESSAGE_LEVEL_WIDTH);

            switch (logLevel)
            {
                case MessageLevel.TestOnly:
                    {
                        screenMessage = id;
                        break;
                    }
                default:
                    {
                        screenMessage = $"{id.PadRight(TEST_NAME_WIDTH)} {messageLevelFormatted} {message}";
                        break;
                    }
            }
            Console.WriteLine($"{id,-TEST_NAME_WIDTH} {message}");
            base.LogMessage(id, message);

            OnLogMessageChanged($"{DateTime.Now:HH:mm:ss.fff} {screenMessage}");
        }

        public new void Log(LogLevel level, string message)
        {
            base.Log(level, message);
        }

        internal void OnLogMessageChanged(string message)
        {
            MessageEventArgs eventArgs = new()
            {
                Message = message
            };

            EventHandler<MessageEventArgs> messageEventHandler = OutputChanged;

            if (OutputChanged is not null)
            {
                messageEventHandler(this, eventArgs);
            }
        }

        internal void OnStatusChanged(string status)
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

    }
}