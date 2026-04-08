using ConformU;
using System;
using System.Collections.Generic;
using System.Text;
namespace ConformU
{

    /// <summary>
    /// Represents a safety-related event, including its condition, type, source, message, and the UTC time it occurred.
    /// </summary>
    public class SafetyState
    {
        /// <summary>
        /// The human-readable name of the application, device, or driver that generated the event.
        /// </summary>
        public string EventSource { get; set; } = string.Empty;

        /// <summary>
        /// A human-readable name for the rule that triggered this event e.g. "Voltage Monitor" or "Wind Speed Monitor".
        /// </summary>
        public string RuleName { get; set; } = string.Empty;

        /// <summary>
        /// A unique ID, defined by the event source, to identify the rule that triggered this event e.g. a GUID.
        /// </summary>
        /// <remarks>
        /// This field allows applications to update or remove an event that has already been sent to the safety monitor e.g. if the wind speed changes but the safety rule violation remains.
        /// This avoids the need to manage multiple events for the same condition.
        /// </remarks>
        public string RuleId { get; set; } = string.Empty;

        /// <summary>
        /// The type of safety event that triggered the condition.
        /// </summary>
        public SafetyEventType EventType { get; set; }

        /// <summary>
        /// The rule that triggered the event.
        /// </summary>
        public SafetyEventCondition EventCondition { get; set; }

        /// <summary>
        /// A message providing additional context about the safety event.
        /// </summary>
        public string EventMessage { get; set; } = string.Empty;

        /// <summary>
        /// The UTC time at which the event occurred.
        /// </summary>
        public DateTime EventTimeUtc { get; set; }

        public SafetyState()
        {

        }

        /// <summary>
        /// Initializes a new instance of the SafetyEvent class with the specified event source, rule name, rule ID, type, condition, message, and time.
        /// </summary>
        /// <param name="eventSource">The component that generated the event.</param>
        /// <param name="ruleName">A human-readable name for the rule that triggered this event.</param>
        /// <param name="ruleId">A unique ID, defined by the event source, to identify the rule that triggered this event.</param>
        /// <param name="eventType">The category or type of the safety event.</param>
        /// <param name="eventCondition">The condition that triggered the safety event.</param>
        /// <param name="eventMessage">A message providing additional context about the safety event.</param>
        public SafetyState(string eventSource, string ruleName, string ruleId, SafetyEventType eventType, SafetyEventCondition eventCondition, string eventMessage)
        {
            ArgumentNullException.ThrowIfNull(eventSource, nameof(eventSource));
            ArgumentNullException.ThrowIfNull(ruleName, nameof(ruleName));
            ArgumentNullException.ThrowIfNull(ruleId, nameof(ruleId));
            ArgumentNullException.ThrowIfNull(eventMessage, nameof(eventMessage));

            EventSource = eventSource;
            RuleName = ruleName;
            RuleId = ruleId;
            EventType = eventType;
            EventCondition = eventCondition;
            EventMessage = eventMessage;
            EventTimeUtc = DateTime.UtcNow;
        }
    }
}

