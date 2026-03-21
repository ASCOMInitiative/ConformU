namespace ConformU
{
    /// <summary>
    /// Specifies the nature of the safety rule that triggered this event.
    /// </summary>
    /// <remarks>Use this enumeration to categorize the state of monitored parameters in safety monitoring
    /// systems. It enables clear identification of whether a parameter has breached defined thresholds or requires
    /// special handling under the 'Other' category.</remarks>
    public enum SafetyEventCondition
    {
        /// <summary>
        /// The property has fallen below the safety threshold defined for this property.
        /// </summary>
        /// <remarks>Only for ObservingConditions devices.</remarks>
        BelowLimit = 0,

        /// <summary>
        /// The property has reached the safety threshold defined for this property.
        /// </summary>
        /// <remarks>Only for ObservingConditions devices.</remarks>
        EqualLimit = 1,

        /// <summary>
        /// The property has exceeded the safety threshold defined for this property.
        /// </summary>
        /// <remarks>Only for ObservingConditions devices.</remarks>
        AboveLimit = 2,

        /// <summary>
        /// The property is in an unsafe state.
        /// </summary>
        /// <remarks>Only for SafetyMonitor devices.</remarks>
        Unsafe = 3,

        /// <summary>
        /// The property has been forced to a specific state or value.
        /// </summary>
        /// <remarks>For all devices.</remarks>
        ForcedToState = 4,

        /// <summary>
        /// The device is in an error state.
        /// </summary>
        /// <remarks>For all devices.</remarks>
        DeviceInErrorState = 5,
    }
}