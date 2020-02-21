using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace D7ComAgentClientDOTNETClient
{
	public partial class Form1 : Form
	{
		/// <summary>
		/// used by StringToXMLAttributeValue()
		/// </summary>
		private static XmlDocument _xmlDoc = new XmlDocument();
		/// <summary>
		/// used by StringToXMLAttributeValue()
		/// </summary>
		private static XmlAttribute _attr = _xmlDoc.CreateAttribute("a");

		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				D7COMAGENTCLIENTLib.INRCOMClient clt = new D7COMAGENTCLIENTLib.NRCOMClient();
				try
				{
					clt.Connect("127.0.0.1", 12345, "netreport");
					NRNELRFXLib.INRNelrfXO2 hlp = new NRNELRFXLib.NRNelrfXO();
					hlp.type = "SomeFilter";
					var now = DateTime.UtcNow;
					hlp.time = now;
					hlp.milliSeconds = (short)now.Millisecond;
					hlp["somexml"] = StringToXMLAttributeValue(@"<doc><info attr1=""value1""></doc>");
					hlp.ip = GetLocalIPv4Address();
					hlp["path"] = StringToXMLAttributeValue(@"c:\temp");
					hlp["name"] = StringToXMLAttributeValue("test.vbs");
					var xml = hlp.xml;
					if (MessageBox.Show(xml, "SEND?", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
						clt.SendNelrfRecord(xml);
				}
				finally
				{
					Marshal.FinalReleaseComObject(clt); // forces disconnection
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private static string GetLocalIPv4Address()
		{
			var host = Dns.GetHostEntry(Dns.GetHostName());
			foreach (var ip in host.AddressList)
			{
				if (ip.AddressFamily == AddressFamily.InterNetwork)
					return ip.ToString();
			}
			return "127.0.0.1";
		}

		private static string StringToXMLAttributeValue(string s)
		{
			_attr.Value = s;
			return _attr.InnerXml;
		}
	}
}
