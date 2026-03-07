namespace ConformU
{
    /// <summary>
    /// Specifies the types of safety-related environmental events that can be monitored or reported.
    /// </summary>
    /// <remarks>Use this enumeration to identify and categorize environmental parameters such as temperature,
    /// humidity, wind speed, and other conditions relevant to safety assessments. Each value represents a distinct
    /// measurement or status that may influence operational safety decisions. This enumeration is typically used in
    /// systems that monitor environmental conditions to determine or report safety status.</remarks>
    public enum SafetyEventType
    {
        CloudCover,
        DewPoint,
        Humidity,
        Pressure,
        RainRate,
        SkyBrightness,
        SkyQuality,
        SkyTemperature,
        StarFWHM,
        Temperature,
        WindDirection,
        WindGust,
        WindSpeed,
        SafetyIssue,
        SecurityIssue,
        PowerIssue,
        Other
    }
}
