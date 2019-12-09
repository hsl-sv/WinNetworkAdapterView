using System;
using System.Data;
using System.Management;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

struct NAC
{
    public string[] DefaultIPGateway;
    public sbyte DefaultTOS;
    public sbyte DefaultTTL;
    public bool DHCPEnabled;
    public DateTime DHCPLeaseExpires;
    public DateTime DHCPLeaseObtained;
    public string DHCPServer;
    public string DNSDomain;
    public string DNSHostName;
    public string[] DNSServerSearchOrder;
    public string[] IPAddress;
    public UInt32 IPConnectionMetric;
    public bool IPEnabled;
    public string[] IPSubnet;
    public bool IPUseZeroBroadcast;
    public UInt32 KeepAliveInterval;
    public UInt32 KeepAliveTime;
    public string MACAddress;
    public UInt32 MTU;
    public string ServiceName;
    public UInt16 TcpWindowSize;
};

struct NA
{
    public string AdapterType;
    public string Caption;
    public string MACAddress;
    public string Manufacturer;
    public UInt64 MaxSpeed;
    public string Name;
    public string NetConnectionID;
};

namespace WinNetworkAdapterView
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable dt = new DataTable();
        ManagementClass MC_NAC = new ManagementClass("Win32_NetworkAdapterConfiguration");
        ManagementClass MC_NA = new ManagementClass("Win32_NetworkAdapter");

        public MainWindow()
        {
            InitializeComponent();

            dt.Columns.Add("Caption", typeof(string));
            dt.Columns.Add("AdapterType", typeof(string));
            dt.Columns.Add("Manufacturer", typeof(string));
            dt.Columns.Add("NetConnectionID", typeof(string));
            dt.Columns.Add("IPv4", typeof(string));
            dt.Columns.Add("IPv6", typeof(string));
            dt.Columns.Add("IPSubnet", typeof(string));
            dt.Columns.Add("DefaultIPGateway", typeof(string));
            dt.Columns.Add("DNSServerSearchOrder", typeof(string));
            dt.Columns.Add("MACAddress", typeof(string));

            GetNetworks();

            RoutedCommand refresh = new RoutedCommand();
            KeyBinding b = new KeyBinding()
            {
                Command = refresh,
                Key = Key.R
            };
            InputBindings.Add(b);
            CommandBindings.Add(new CommandBinding(refresh, Refresh));
        }

        private void Refresh(object sender, ExecutedRoutedEventArgs e)
        {
            GetNetworks();
        }

        private void GetNetworks()
        {
            ManagementObjectCollection NAMOC = MC_NA.GetInstances();
            ManagementObjectCollection NACMOC = MC_NAC.GetInstances();

            int i = NAMOC.Count;

            NAC[] nac = new NAC[i];
            NA[] na = new NA[i];
            string[] ipv4 = new string[i];
            string[] ipv6 = new string[i];
            string[] subnet = new string[i];
            string[] gateway = new string[i];
            string[] dns = new string[i];

            dt.Rows.Clear();

            int idx = 0;
            foreach (ManagementObject MO in NAMOC)
            {
                if (MO["Caption"] != null) na[idx].Caption = MO["Caption"].ToString();
                else na[idx].Caption = "";

                if (MO["AdapterType"] != null) na[idx].AdapterType = MO["AdapterType"].ToString();
                else na[idx].AdapterType = "";

                if (MO["Manufacturer"] != null) na[idx].Manufacturer = MO["Manufacturer"].ToString();
                else na[idx].Manufacturer = "";

                if (MO["NetConnectionID"] != null) na[idx].NetConnectionID = MO["NetConnectionID"].ToString();
                else na[idx].NetConnectionID = "";

                if (MO["MACAddress"] != null) na[idx].MACAddress = MO["MACAddress"].ToString();
                else na[idx].MACAddress = "";

                idx++;
            }

            idx = 0;
            foreach (ManagementObject MO in NACMOC)
            {
                if ((string[])MO["IPAddress"] != null)
                {
                    nac[idx].IPAddress = (string[])MO["IPAddress"];
                    ipv4[idx] = nac[idx].IPAddress[0];
                    if (nac[idx].IPAddress.Length != 1)
                    {
                        ipv6[idx] = nac[idx].IPAddress[1];
                    }
                    else
                    {
                        ipv6[idx] = "";
                    }
                }
                else { ipv4[idx] = ""; ipv6[idx] = ""; }

                if (MO["IPSubnet"] != null)
                {
                    nac[idx].IPSubnet = (string[])MO["IPSubnet"];
                    subnet[idx] = nac[idx].IPSubnet[0];
                }
                else { subnet[idx] = ""; }

                if (MO["DefaultIPGateway"] != null)
                {
                    nac[idx].DefaultIPGateway = (string[])MO["DefaultIPGateway"];
                    gateway[idx] = nac[idx].DefaultIPGateway[0];
                }
                else { gateway[idx] = ""; }

                if (MO["DNSServerSearchOrder"] != null)
                {
                    nac[idx].DNSServerSearchOrder = (string[])MO["DNSServerSearchOrder"];
                    dns[idx] = nac[idx].DNSServerSearchOrder[0];
                }
                else
                {
                    dns[idx] = "";
                }

                idx++;
            }

            for (int k = 0; k < NAMOC.Count; k++)
            {
                dt.Rows.Add(new string[] {
                    na[k].Caption,
                    na[k].AdapterType,
                    na[k].Manufacturer,
                    na[k].NetConnectionID,
                    ipv4[k],
                    ipv6[k],
                    subnet[k],
                    gateway[k],
                    dns[k],
                    na[k].MACAddress
                });
            }

            dg.ItemsSource = dt.DefaultView;
        }

        private void Dg_CellEditEnding(object sender, System.Windows.Controls.DataGridCellEditEndingEventArgs e)
        {
            ManagementObjectCollection NACMOC = MC_NAC.GetInstances();
            var row = e.Row.GetIndex();
            var col = e.Column.DisplayIndex;
            DataRowView drow = (DataRowView)dg.SelectedItems[0];

            foreach (ManagementObject MO in NACMOC)
            {
                if (MO["Caption"].ToString() == drow["Caption"].ToString())
                {
                    switch (col)
                    {
                        case 4: // IPv4
                            if ((((TextBox)e.EditingElement).Text == "auto") || (((TextBox)e.EditingElement).Text == ""))
                            {
                                MO.InvokeMethod("EnableDHCP", null, null);
                            }
                            else
                            {
                                var newIP = MO.GetMethodParameters("EnableStatic");
                                ManagementBaseObject setIP;
                                newIP["IPAddress"] = new string[] { ((TextBox)e.EditingElement).Text };
                                newIP["SubnetMask"] = new string[] { drow["IPSubnet"].ToString() };
                                setIP = MO.InvokeMethod("EnableStatic", newIP, null);
                            }
                            break;
                        case 6: // IPSubnet
                            var newSubnet = MO.GetMethodParameters("EnableStatic");
                            ManagementBaseObject setSubnet;
                            newSubnet["IPAddress"] = new string[] { drow["IPv4"].ToString() };
                            newSubnet["SubnetMask"] = new string[] { ((TextBox)e.EditingElement).Text };
                            setSubnet = MO.InvokeMethod("EnableStatic", newSubnet, null);

                            break;
                        case 7: // DefaultIPGateway
                            var newGateway = MO.GetMethodParameters("SetGateways");
                            ManagementBaseObject setGateway;
                            newGateway["DefaultIPGateway"] = new string[] { ((TextBox)e.EditingElement).Text };
                            newGateway["GatewayCostMetric"] = new int[] { 1 };
                            setGateway = MO.InvokeMethod("SetGateways", newGateway, null);
                            break;
                        case 8: // DNSDomain
                            if ((((TextBox)e.EditingElement).Text == "auto") || (((TextBox)e.EditingElement).Text == ""))
                            {
                                var newDNS = MO.GetMethodParameters("SetDNSServerSearchOrder");
                                ManagementBaseObject setDNS;
                                newDNS["DNSServerSearchOrder"] = null;
                                setDNS = MO.InvokeMethod("SetDNSServerSearchOrder", newDNS, null);
                            }
                            else
                            {
                                var newDNS = MO.GetMethodParameters("SetDNSServerSearchOrder");
                                ManagementBaseObject setDNS;
                                newDNS["DNSServerSearchOrder"] = new string[] { ((TextBox)e.EditingElement).Text };
                                setDNS = MO.InvokeMethod("SetDNSServerSearchOrder", newDNS, null);
                            }
                            break;
                        default:
                            break;
                    };
                }
            }
        }
    }
}
