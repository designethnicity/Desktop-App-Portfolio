using Microsoft.VisualBasic.Devices;
using System;
using System.Management;

namespace Project_Portfolio
{
    public partial class Form1 : Form
    {
        private static ManagementObjectSearcher computerSystemSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_ComputerSystem");
        private static ManagementObjectSearcher baseboardSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");

        public Form1()
        {
            InitializeComponent();
            string pcName = Environment.MachineName;
            label15.Text = pcName;
            try
            {
                label7.Text = this.Model;
            }
            catch { }
            try
            {
                ManagementObject myManagementObject = new ManagementObject("Win32_OperatingSystem=@");
                string MySN = (string)myManagementObject["SerialNumber"];
                label11.Text = MySN;
            }
            catch
            {
                label11.Text = this.SerialNumber;
            }
            try
            {
                SelectQuery SqDrive = new SelectQuery("Win32_DiskDrive");
                ManagementObjectSearcher objOSDetailsDrive = new ManagementObjectSearcher(SqDrive);
                ManagementObjectCollection osDetailsCollectionDrive = objOSDetailsDrive.Get();
                UInt64 total = 0;
                foreach (ManagementObject mo in osDetailsCollectionDrive)
                {
                    total += (UInt64)mo["Size"];
                }
                label10.Text = FormatBytes((UInt64)total);
            }
            catch { }
            try
            {
                string windowsEdition = GetWindowsEdition();
                label13.Text = windowsEdition;
            }
            catch { }
            try
            {
                ManagementScope scope = new ManagementScope(@"\\.\root\microsoft\windows\storage");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM MSFT_PhysicalDisk");
                string type = "";
                scope.Connect();
                searcher.Scope = scope;

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    switch (Convert.ToInt16(queryObj["MediaType"]))
                    {
                        case 1:
                            type = "Unspecified";
                            break;

                        case 3:
                            type = "HDD";
                            break;

                        case 4:
                            type = "SSD";
                            break;

                        case 5:
                            type = "SCM";
                            break;

                        default:
                            type = "Unspecified";
                            break;
                    }
                }
                label10.Text = label10.Text + ' ' + type;
                searcher.Dispose();
            }
            catch { }
            try
            {
                ComputerInfo cminfo = new ComputerInfo();
                label9.Text = FormatBytes(cminfo.TotalPhysicalMemory);
            }
            catch { }
            try
            {
                SelectQuery Sq = new SelectQuery("Win32_Processor");
                ManagementObjectSearcher objOSDetails = new ManagementObjectSearcher(Sq);
                ManagementObjectCollection osDetailsCollection = objOSDetails.Get();
                foreach (ManagementObject mo in osDetailsCollection)
                {
                    label8.Text = (string)mo["Name"];
                }
            }
            catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        public string Model
        {
            get
            {
                try
                {
                    foreach (ManagementObject queryObj in computerSystemSearcher.Get())
                    {
                        return queryObj["Model"].ToString();
                    }
                    return "";
                }
                catch (Exception) { return ""; }
            }
        }
        public string SerialNumber
        {
            get
            {
                try
                {
                    foreach (ManagementObject queryObj in baseboardSearcher.Get())
                    {
                        return queryObj["SerialNumber"].ToString();
                    }
                    return "";
                }
                catch (Exception)
                {
                    return "";
                }
            }
        }
        static string GetWindowsEdition()
        {
            string edition = "";

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    edition = obj["Caption"].ToString();
                    break; // assuming there's only one result
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return edition;
        }
        private static string FormatBytes(ulong bytes)
        {
            string[] Suffix = { "B", "KB", "MB", "GB", "TB" };
            int i;
            double dblSByte = bytes;
            for (i = 0; i < Suffix.Length && bytes >= 1024; i++, bytes /= 1024)
            {
                dblSByte = bytes / 1024.0;
            }

            return String.Format("{0:0.##} {1}", Convert.ToInt64(dblSByte), Suffix[i]);
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }
    }
}
