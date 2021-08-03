// (c) 2019 Daniel Van Noord
// This code is licensed under MIT license (see License.txt for details)

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using AlpacaDiscovery;
using ASCOM.Standard.Utilities;

namespace AlpacaDiscovery
{


    // This namespace dual targets NetStandard2.0 and Net35, thus no async await
    internal class Finder : IDisposable
    {
        private readonly TraceLogger TL;
        private readonly Action<IPEndPoint, AlpacaDiscoveryResponse> callbackFunctionDelegate;
        private readonly UdpClient udpClient;
        private readonly List<UdpClient> ipV6Discoveryclients = new ();

        /// <summary>
        /// A cache of all endpoints found by the server
        /// </summary>
        public List<IPEndPoint> CachedEndpoints { get; private set; } = new ();

        #region Initialisation and Dispose
        /// <summary>
        /// Creates a Alpaca Finder object that sends out a search request for Alpaca devices
        /// The results will be sent to the callback and stored in the cache
        /// Calling search and concatenating the results reduces the chance that a UDP packet is lost
        /// This may require firewall access
        /// </summary>
        /// <param name="callback">A callback function to receive the endpoint result</param>
        internal Finder(Action<IPEndPoint, AlpacaDiscoveryResponse> callback, TraceLogger traceLogger)
        {
            TL = traceLogger; // Save the trace logger object
            LogMessage("Finder", "Starting Initialisation...");
            callbackFunctionDelegate = callback;
            udpClient = new UdpClient
            {
                EnableBroadcast = true,
                MulticastLoopback = false
            };

            // 0 tells OS to give us a free ethereal port
            udpClient.Client.Bind(new IPEndPoint(IPAddress.Any, 0));
            udpClient.BeginReceive(new AsyncCallback(FinderDiscoveryCallback), udpClient);
            LogMessage("Finder", "Initialised");
        }

        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (udpClient is object)
                    {
                        try
                        {
                            udpClient.Close();
                        }
                        catch
                        {
                        }
                    }

                    foreach (UdpClient client in ipV6Discoveryclients)
                    {
                        try
                        {
                            client.Close();
                        }
                        catch
                        {
                        }
                    }

                    // try { udpClient.Dispose(); } catch { }
                }

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put clean-up code in Dispose(bool disposing) above.
            Dispose(true);
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Resends the search request
        /// </summary>
        public void Search(int discoveryport, bool ipV4Enabled, bool ipV6Enabled)
        {
            if (ipV4Enabled)
                SendDiscoveryMessageIpV4(discoveryport);
            if (ipV6Enabled)
                SendDiscoveryMessageIpV6(discoveryport);
        }

        /// <summary>
        /// Clears the cached IP Endpoints in CachedEndpoints
        /// </summary>
        public void ClearCache()
        {
            CachedEndpoints.Clear();
        }

        #endregion

        private void SendDiscoveryMessageIpV4(int discoveryPort)
        {
            var adapters = NetworkInterface.GetAllNetworkInterfaces();
            LogMessage("SendDiscoveryMessageIpV4", $"Sending IPv$ discovery broadcasts");
            foreach (NetworkInterface adapter in adapters)
            {
                // Do not try and use non-operational adapters
                if (adapter.OperationalStatus != OperationalStatus.Up)
                    continue;
                if (adapter.Supports(NetworkInterfaceComponent.IPv4))
                {
                    var adapterProperties = adapter.GetIPProperties();
                    if (adapterProperties is null)
                        continue;
                    var uniCast = adapterProperties.UnicastAddresses;
                    if (uniCast.Count > 0)
                    {
                        foreach (UnicastIPAddressInformation uni in uniCast)
                        {
                            if (uni.Address.AddressFamily != AddressFamily.InterNetwork)
                                continue;

                            // Local host addresses (127.*.*.*) may have a null mask in Net Framework. We do want to search these. The correct mask is 255.0.0.0.
                            udpClient.Send(Encoding.ASCII.GetBytes(Constants.DISCOVERY_MESSAGE), Encoding.ASCII.GetBytes(Constants.DISCOVERY_MESSAGE).Length, new IPEndPoint(GetBroadcastAddress(uni.Address, uni.IPv4Mask ?? IPAddress.Parse("255.0.0.0")), discoveryPort));
                            LogMessage("SendDiscoveryMessageIpV4", $"Sent broadcast to: {uni.Address}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Send out discovery message on the IPv6 multicast group
        /// This dual targets NetStandard 2.0 and NetFX 3.5 so no Async Await
        /// </summary>
        private void SendDiscoveryMessageIpV6(int discoveryPort)
        {
            foreach (UdpClient client in ipV6Discoveryclients)
            {
                try
                {
                    client.Close();
                }
                catch
                {
                }
            }

            ipV6Discoveryclients.Clear();
            LogMessage("SendDiscoveryMessageIpV6", $"Sending IPv6 discovery broadcasts");

            // Windows needs to bind a socket to each adapter explicitly
            foreach (NetworkInterface adapter in NetworkInterface.GetAllNetworkInterfaces())
            {
                LogMessage("SendDiscoveryMessageIpV6", $"Found network interface {adapter.Description}, Interface type: {adapter.NetworkInterfaceType} - supports multicast: {adapter.SupportsMulticast}");
                if (adapter.OperationalStatus != OperationalStatus.Up)
                    continue;
                LogMessage("SendDiscoveryMessageIpV6", $"Interface is up");
                if (adapter.Supports(NetworkInterfaceComponent.IPv6) && adapter.SupportsMulticast)
                {
                    LogMessage("SendDiscoveryMessageIpV6", $"Interface supports IPv6");
                    var adapterProperties = adapter.GetIPProperties();
                    if (adapterProperties is null)
                        continue;
                    var uniCast = adapterProperties.UnicastAddresses;
                    LogMessage("SendDiscoveryMessageIpV6", $"Adapter does have properties. Number of unicast addresses: {uniCast.Count}");
                    if (uniCast.Count > 0)
                    {
                        foreach (UnicastIPAddressInformation uni in uniCast)
                        {
                            if (uni.Address.AddressFamily == AddressFamily.InterNetworkV6)
                            {
                                LogMessage("SendDiscoveryMessageIpV6", $"Interface is IPv6");
                                try
                                {
                                    var client = new UdpClient(AddressFamily.InterNetworkV6);

                                    // 0 tells OS to give us a free ethereal port
                                    client.Client.Bind(new IPEndPoint(uni.Address, 0));
                                    client.BeginReceive(new AsyncCallback(FinderDiscoveryCallback), client);
                                    client.Send(Encoding.ASCII.GetBytes(Constants.DISCOVERY_MESSAGE), Encoding.ASCII.GetBytes(Constants.DISCOVERY_MESSAGE).Length, new IPEndPoint(IPAddress.Parse(Constants.ALPACA_DISCOVERY_IPV6_MULTICAST_ADDRESS), discoveryPort));
                                    LogMessage("SendDiscoveryMessageIpV6", $"Sent multicast IPv6 discovery packet");
                                    ipV6Discoveryclients.Add(client);
                                }
                                catch 
                                {
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This callback is shared between IPv4 and IPv6
        /// </summary>
        /// <param name="ar"></param>
        private void FinderDiscoveryCallback(IAsyncResult ar)
        {
            try
            {
                UdpClient udpClient = (UdpClient)ar.AsyncState;
                var alpacaBroadcastResponseEndPoint = new IPEndPoint(IPAddress.Any, Constants.DEFAULT_DISCOVERY_PORT);

                // Obtain the UDP message body and convert it to a string, with remote IP address attached as well
                string ReceiveString = Encoding.ASCII.GetString(udpClient.EndReceive(ar, ref alpacaBroadcastResponseEndPoint));
                LogMessage($"FinderDiscoveryCallback", $"Received {ReceiveString} from Alpaca device at {alpacaBroadcastResponseEndPoint.Address}");

                // Configure the UdpClient class to accept more messages, if they arrive
                udpClient.BeginReceive(new AsyncCallback(FinderDiscoveryCallback), udpClient);

                // Only process Alpaca device responses
                if (ReceiveString.ToLowerInvariant().Contains(Constants.DISCOVERY_RESPONSE_STRING))
                {
                    // Extract the discovery response parameters from the device's JSON response
                    var discoveryResponse = JsonSerializer.Deserialize<AlpacaDiscoveryResponse>(ReceiveString);
                    var alpacaApiEndpoint = new IPEndPoint(alpacaBroadcastResponseEndPoint.Address, discoveryResponse.AlpacaPort); // Create 
                    if (!CachedEndpoints.Contains(alpacaApiEndpoint))
                    {
                        CachedEndpoints.Add(alpacaApiEndpoint);
                        LogMessage("FinderDiscoveryCallback", $"Received new Alpaca API endpoint: {alpacaApiEndpoint} from broadcast endpoint: {alpacaBroadcastResponseEndPoint}");
                        callbackFunctionDelegate?.Invoke(alpacaApiEndpoint, discoveryResponse); // Moved inside the loop so that the callback is only called once per IP address
                    }
                    else
                    {
                        LogMessage("FinderDiscoveryCallback", $"Ignoring duplicate Alpaca API endpoint: {alpacaApiEndpoint} from broadcast endpoint: {alpacaBroadcastResponseEndPoint}");
                    }
                }
            }

            // Ignore these, they can occur after the Finder is disposed
            catch (ObjectDisposedException)
            {
            }
            catch (Exception ex)
            {
                LogMessage("FinderDiscoveryCallback", $"Exception: " + ex.ToString());
            }
        }


        // This turns the unicast address and the subnet into the broadcast address for that range
        // http://blogs.msdn.com/b/knom/archive/2008/12/31/ip-address-calculations-with-c-subnetmasks-networks.aspx
        private static IPAddress GetBroadcastAddress(IPAddress address, IPAddress subnetMask)
        {
            var ipAdressBytes = address.GetAddressBytes();
            var subnetMaskBytes = subnetMask.GetAddressBytes();
            if (ipAdressBytes.Length != subnetMaskBytes.Length)
                throw new ArgumentException("Lengths of IP address and subnet mask do not match.");
            var broadcastAddress = new byte[ipAdressBytes.Length];
            for (int i = 0, loopTo = broadcastAddress.Length - 1; i <= loopTo; i++)
                broadcastAddress[i] = (byte)(ipAdressBytes[i] | subnetMaskBytes[i] ^ 255);
            return new IPAddress(broadcastAddress);
        }

        private void LogMessage(string method, string message)
        {
            if (TL is object)
            {
                string indentSpaces = new(' ', Thread.CurrentThread.ManagedThreadId * Constants.NUMBER_OF_THREAD_MESSAGE_INDENT_SPACES);
                TL.LogMessage($"Finder - {method}", $"{indentSpaces}{Thread.CurrentThread.ManagedThreadId} - {message}");
            }
        }
    }
}