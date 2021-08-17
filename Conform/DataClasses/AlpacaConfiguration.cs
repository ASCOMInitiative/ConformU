﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class AlpacaConfiguration
    {
        public int NumberOfDiscoveryPolls { get; set; } = 1;
        public int DiscoveryPollInterval { get; set; } = 1;
        public int DiscoveryPort { get; set; } = 32227;
        public double DiscoveryDuration { get; set; } = 1.0;
        public bool DiscoveryResolveName { get; set; } = false;
        public bool DiscoveryUseIpV4 { get; set; } = true;
        public bool DiscoveryUseIpV6 { get; set; } = false;
        public ServiceType AccessServiceType { get; set; } = ServiceType.Http;
        public string AccessUserName { get; set; } = "";
        public string AccessPassword { get; set; } = "";
    }
}