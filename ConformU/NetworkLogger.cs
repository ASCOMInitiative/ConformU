using System.Diagnostics.Tracing;
using System.Text;
using System;
using ASCOM.Tools;
using Microsoft.Extensions.Logging;

namespace ConformU
{
    /// <summary>
    /// Class to make .NET internal log messages appear on the console and in the trace logger log.
    /// </summary>
    internal sealed class NetworkLogger : EventListener
    {
        // Constant necessary for attaching ActivityId to the events.
        public const EventKeywords TasksFlowActivityIds = (EventKeywords)0x80;
        private readonly TraceLogger logger;

        internal NetworkLogger(TraceLogger logger)
        {
            this.logger = logger;
        }

        protected override void OnEventSourceCreated(EventSource eventSource)
        {
            // List of event source names provided by networking in .NET 5.
            if (eventSource.Name == "System.Net.Http" ||
                eventSource.Name == "System.Net.Sockets" ||
                eventSource.Name == "System.Net.Security" ||
                eventSource.Name == "System.Net.NameResolution")
            {
                EnableEvents(eventSource, EventLevel.LogAlways);
            }
            // Turn on ActivityId.
            else if (eventSource.Name == "System.Threading.Tasks.TplEventSource")
            {
                // Attach ActivityId to the events.
                EnableEvents(eventSource, EventLevel.Verbose, TasksFlowActivityIds);
            }
        }

        protected override void OnEventWritten(EventWrittenEventArgs eventData)
        {
            var sb = new StringBuilder().Append($"{eventData.EventSource.Name}.{eventData.EventName}(");
            for (int i = 0; i < eventData.Payload?.Count; i++)
            {
                sb.Append(eventData.PayloadNames?[i]).Append(": ").Append(eventData.Payload[i]);
                if (i < eventData.Payload?.Count - 1)
                {
                    sb.Append(", ");
                }
            }

            sb.Append(')');
            Console.WriteLine(sb.ToString());
            logger.LogMessage("NetworkLogger", sb.ToString());
        }
    }
}
