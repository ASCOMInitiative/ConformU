namespace ConformU
{
    /// <summary>
    /// Information or issue message shown when running Alpaca discovery diagnostics
    /// </summary>
    public class AlpacaDiscoveryDiagnostic
    {
        // Create a new diagnostic message list entry
        public AlpacaDiscoveryDiagnostic(string deviceIp,string message)
        {
            DeviceIp = deviceIp;
            Message= message;
        }

        /// <summary>
        ///  Device IP address
        /// </summary>
        public string DeviceIp { get; set; }

        // Message shown
        public string Message { get; set; }
    }
}
