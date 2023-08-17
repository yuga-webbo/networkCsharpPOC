using Microsoft.WindowsAPICodePack.Net;
using System;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using Misc;

class networkPOC
{
    public static void ShowNetworkInterfaces()
    {
        IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();
        NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
        Console.WriteLine("Interface information for {0}.{1}     ",
                computerProperties.HostName, computerProperties.DomainName);
        if (nics == null || nics.Length < 1)
        {
            Console.WriteLine("  No network interfaces found.");
            return;
        }

        Console.WriteLine("  Number of interfaces .................... : {0}", nics.Length);
        foreach (NetworkInterface adapter in nics)
        {
            IPInterfaceProperties properties = adapter.GetIPProperties();
            Console.WriteLine();
            Console.WriteLine(adapter.Description);
            Console.WriteLine(String.Empty.PadLeft(adapter.Description.Length, '='));
            Console.WriteLine("  Interface type .......................... : {0}", adapter.NetworkInterfaceType);
            Console.WriteLine("  Physical Address ........................ : {0}",
                        adapter.GetPhysicalAddress().ToString());
            Console.WriteLine("  Operational status ...................... : {0}",
                adapter.OperationalStatus);
            string versions = "";

            // Create a display string for the supported IP versions.
            if (adapter.Supports(NetworkInterfaceComponent.IPv4))
            {
                versions = "IPv4";
            }
            if (adapter.Supports(NetworkInterfaceComponent.IPv6))
            {
                if (versions.Length > 0)
                {
                    versions += " ";
                }
                versions += "IPv6";
            }
            Console.WriteLine("  IP version .............................. : {0}", versions);
            // ShowIPAddresses(properties);

            // The following information is not useful for loopback adapters.
            if (adapter.NetworkInterfaceType == NetworkInterfaceType.Loopback)
            {
                continue;
            }
            Console.WriteLine("  DNS suffix .............................. : {0}",
                properties.DnsSuffix);

            string label;
            if (adapter.Supports(NetworkInterfaceComponent.IPv4))
            {
                IPv4InterfaceProperties ipv4 = properties.GetIPv4Properties();
                Console.WriteLine("  MTU...................................... : {0}", ipv4.Mtu);
                if (ipv4.UsesWins)
                {

                    IPAddressCollection winsServers = properties.WinsServersAddresses;
                    if (winsServers.Count > 0)
                    {
                        label = "  WINS Servers ............................ :";
                        // ShowIPAddresses(label, winsServers);
                    }
                }
            }

            Console.WriteLine("  DNS enabled ............................. : {0}",
                properties.IsDnsEnabled);
            Console.WriteLine("  Dynamically configured DNS .............. : {0}",
                properties.IsDynamicDnsEnabled);
            Console.WriteLine("  Receive Only ............................ : {0}",
                adapter.IsReceiveOnly);
            Console.WriteLine("  Multicast ............................... : {0}",
                adapter.SupportsMulticast);
            // ShowInterfaceStatistics(adapter);

            Console.WriteLine();
        }
    }
    public static void ShowNetworkInterfacesPhysicalAddress()
    {
        IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();
        NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
        Console.WriteLine("Interface information for {0}.{1}     ",
                computerProperties.HostName, computerProperties.DomainName);
        if (nics == null || nics.Length < 1)
        {
            Console.WriteLine("  No network interfaces found.");
            return;
        }

        Console.WriteLine("  Number of interfaces .................... : {0}", nics.Length);
        foreach (NetworkInterface adapter in nics)
        {
            IPInterfaceProperties properties = adapter.GetIPProperties(); //  .GetIPInterfaceProperties();
            Console.WriteLine();
            Console.WriteLine(adapter.Description);
            Console.WriteLine(String.Empty.PadLeft(adapter.Description.Length, '='));
            Console.WriteLine("  Interface type .......................... : {0}", adapter.NetworkInterfaceType);
            Console.Write("  Physical address ........................ : ");
            PhysicalAddress address = adapter.GetPhysicalAddress();
            byte[] bytes = address.GetAddressBytes();
            for (int i = 0; i < bytes.Length; i++)
            {
                // Display the physical address in hexadecimal.
                Console.Write("{0}", bytes[i].ToString("X2"));
                // Insert a hyphen after each byte, unless we're at the end of the address.
                if (i != bytes.Length - 1)
                {
                    Console.Write("-");
                }
            }
            Console.WriteLine();
        }
    }
    public static void DisplayIPv4NetworkInterfaces()
    {
        NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
        IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();
        Console.WriteLine("IPv4 interface information for {0}.{1}",
           properties.HostName, properties.DomainName);
        Console.WriteLine();

        foreach (NetworkInterface adapter in nics)
        {
            // Only display informatin for interfaces that support IPv4.
            if (adapter.Supports(NetworkInterfaceComponent.IPv4) == false)
            {
                continue;
            }
            Console.WriteLine(adapter.Description);
            // Underline the description.
            Console.WriteLine(String.Empty.PadLeft(adapter.Description.Length, '='));
            IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
            // Try to get the IPv4 interface properties.
            IPv4InterfaceProperties p = adapterProperties.GetIPv4Properties();

            if (p == null)
            {
                Console.WriteLine("No IPv4 information is available for this interface.");
                Console.WriteLine();
                continue;
            }
            // Display the IPv4 specific data.
            Console.WriteLine("  Index ............................. : {0}", p.Index);
            Console.WriteLine("  MTU ............................... : {0}", p.Mtu);
            Console.WriteLine("  APIPA active....................... : {0}",
                p.IsAutomaticPrivateAddressingActive);
            Console.WriteLine("  APIPA enabled...................... : {0}",
                p.IsAutomaticPrivateAddressingEnabled);
            Console.WriteLine("  Forwarding enabled................. : {0}",
                p.IsForwardingEnabled);
            Console.WriteLine("  Uses WINS ......................... : {0}",
                p.UsesWins);
            Console.WriteLine();
        }
    }
   
    public static void InviteViewersCommand()
    {
        NetworkInterface[] networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();

        string id = networkInterfaces
            .Where(i => i.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 && i.OperationalStatus == OperationalStatus.Up)
            .Select(i => i.Id).FirstOrDefault();

        Guid networkId;
        string ssid = string.Empty;
        bool isWifiConnected = false;

        if (id != null)
        {
            networkId = new Guid(id);

            var wifiNetworks = NetworkListManager.GetNetworks(NetworkConnectivityLevels.Connected)
                .Where(n => n.DomainType == DomainType.NonDomainNetwork);


            foreach (var network in wifiNetworks)
            {
                var adapter = network.Connections.Where(n => n.AdapterId == networkId).FirstOrDefault();
                if (adapter != null)
                {
                    ssid = adapter.Network.Name;
                    isWifiConnected = true;
                    break;
                }
            }
        }

        if (ssid == string.Empty) ssid = "(blank to be filled out by Hal rep on location)";

        Func<NetworkInterface, bool> networkTypeFilter;

        if (isWifiConnected)
            networkTypeFilter = networkInterface =>
                networkInterface.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 &&
                networkInterface.OperationalStatus == OperationalStatus.Up;
        else
            networkTypeFilter = networkInterface =>
                networkInterface.NetworkInterfaceType == NetworkInterfaceType.Ethernet &&
                networkInterface.OperationalStatus == OperationalStatus.Up;

        var ipadresses = networkInterfaces
                .Where(networkTypeFilter)
                .SelectMany(i => i.GetIPProperties().UnicastAddresses)
                .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
                .Select(a => a.Address.ToString())
                .ToList();

        var port = string.Empty;

        string link = string.Empty;

        for (int i = 0; i < ipadresses.Count; i++)
        {
            link += $"http://{ipadresses[i]}";
            if (!string.IsNullOrEmpty(port))
            {
                link += $":{port}";
            }

            if (i != ipadresses.Count - 1 && ipadresses.Count > 1) link += " ";
        }

        string hostName = null;

        IPAddress localhostIpAddress = networkInterfaces
            .SelectMany(i =>
                i.GetIPProperties().UnicastAddresses)
            .Where(a =>
                a.Address.AddressFamily == AddressFamily.InterNetwork && IPAddress.IsLoopback(a.Address))
            .Select(a =>
                a.Address).
            FirstOrDefault();

        if (localhostIpAddress != null)
        {
            string domainName = IPGlobalProperties.GetIPGlobalProperties().DomainName;
            domainName = "." + domainName;

            IPHostEntry ipHostEntry = Dns.GetHostEntry(localhostIpAddress);
            hostName = ipHostEntry.HostName;

            if (!hostName.EndsWith(domainName))  // if + does not already include domain name
            {
                hostName += domainName;   // add the domain name part
            }

            if (!string.IsNullOrEmpty(port))
            {
                hostName += $":{port}";
            }

            hostName = $"http://{hostName}";
        }

        try
        {
            string subject = "Halliburton Cementing Real-Time Data Stream: Connection Instructions";

            Console.WriteLine("  link ............................. : {0}", link);
            Console.WriteLine("  HostName ............................. : {0}", hostName);
            Console.WriteLine("  ssid ............................. : {0}", ssid);

        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            Console.WriteLine("Unable to create email. Please try closing outlook before trying again");
        }
    }
    public static void Main(string[] args)
    {
        //ShowNetworkInterfaces();
        //ShowNetworkInterfacesPhysicalAddress();
        //DisplayIPv4NetworkInterfaces();
        InviteViewersCommand();
    }
}
