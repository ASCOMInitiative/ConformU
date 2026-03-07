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
        BelowLimit,
        EqualLimit,
        AboveLimit,
        Unsafe,
        Other,
        DeviceInErrorState,
        NotAvailable
    }
}