    using System;
namespace ConformU
{
    namespace ObsMan
    {
        /// <summary>
        /// Represents a safety-related event, including its condition, type, source, message, and the UTC time it occurred.
        /// </summary>
        public class SafetyState
        {
            /// <summary>
            /// The condition that triggered the safety event.
            /// </summary>
            public SafetyEventCondition EventCondition { get; set; }

            /// <summary>
            /// The category or type of the safety event.
            /// </summary>
            public SafetyEventType EventType { get; set; }

            /// <summary>
            /// The component that generated the event.
            /// </summary>
            public string EventSource { get; set; }

            /// <summary>
            /// A message providing additional context or details about the safety event.
            /// </summary>
            public string EventMessage { get; set; }

            /// <summary>
            /// The UTC time at which the event message was created.
            /// </summary>
            /// <remarks>Please note: This is not the time at which the event started, it is the time at which the SafetyState class was created and returned to the client.</remarks>
            public DateTime EventTimeUtc { get; set; }

            /// <summary>
            /// Initializes a new instance of the SafetyEvent class with the specified event condition, event type, source and message.
            /// </summary>
            /// <remarks>The UTC timestamp of the event is automatically set to the current time when the event is created.</remarks>
            /// <param name="eventCondition">The condition that triggered the safety event.</param>
            /// <param name="eventType">The category or type of the safety event.</param>
            /// <param name="eventSource">The component that generated the event.</param>
            /// <param name="eventMessage">A message providing additional context or details about the safety event.</param>
            public SafetyState(SafetyEventCondition eventCondition, SafetyEventType eventType, string eventSource, string eventMessage)
            {
                ArgumentNullException.ThrowIfNull(eventSource, nameof(eventSource));
                ArgumentNullException.ThrowIfNull(eventMessage, nameof(eventMessage));

                EventCondition = eventCondition;
                EventType = eventType;
                EventSource = eventSource;
                EventMessage = eventMessage;
                EventTimeUtc = DateTime.UtcNow;
            }
        }
    }
}
