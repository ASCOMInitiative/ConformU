using ASCOM.Common.Alpaca;

namespace ConformU
{
    public class AlpacaConfiguration
    {
        public bool DiscoveryEnabled { get; set; } = true;
        public bool StrictCasing { get; set; } = true; // Alpaca JSON parsing configuration
        public int NumberOfDiscoveryPolls { get; set; } = 1;
        public double DiscoveryPollInterval { get; set; } = 1.0;
        public int DiscoveryPort { get; set; } = 32227;
        public double DiscoveryDuration { get; set; } = 1.0;
        public bool DiscoveryResolveName { get; set; } = false;
        public bool DiscoveryUseIpV4 { get; set; } = true;
        public bool DiscoveryUseIpV6 { get; set; } = false;
        public ServiceType AccessServiceType { get; set; } = ServiceType.Http;
        public string AccessUserName { get; set; } = "";
        public string AccessPassword { get; set; } = "";
        public ImageArrayTransferType ImageArrayTransferType{ get; set; } = ImageArrayTransferType.BestAvailable;
        public ImageArrayCompression ImageArrayCompression { get; set; } = ImageArrayCompression.None;
        public int EstablishConnectionTimeout { get; set; } = 5;
        public int StandardResponseTimeout{ get; set; } = 10;
        public int LongResponseTimeout{ get; set; } = 100;
        public bool TrustUserGeneratedSslCertificates { get; set; } = false;

        // ALpaca protocol test configuration

        [MandatoryInFullTest(false)]
        public bool ProtocolShowSuccessResponses { get; set; } = true;
        public bool ProtocolReportNotImplementedErrors { get; set; } = false;
        public bool ProtocolTestPrimaryUrlStructure { get; set; } = false;
        public bool ProtocolStrictChecks { get; set; } = false;
        public bool ProtocolShowOkMessages { get; set; } = false;
    }
}
