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

        public void LogMessage(string id, MessageLevel logLevel, string message)
        {
            string formattedMessage = logLevel.ToString().PadRight(MESSAGE_LEVEL_WIDTH) + message;
            string screenMessage;

            string messageLevelFormatted = "        ";

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
            LogMessage(id, formattedMessage);
        }

        public new void LogMessage(string id, string message)
        {
            Console.WriteLine($"{id,-TEST_NAME_WIDTH} {message}");
            base.LogMessage(id, message);

            OnLogMessageChanged("LogMessage", $"{DateTime.Now:HH:mm:ss.fff} {message}");

        }

        public new void Log(LogLevel level, string message)
        {
            base.Log(level, message);
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

    }
}