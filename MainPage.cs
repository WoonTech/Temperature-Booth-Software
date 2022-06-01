//================================================= DO NOT DELETE AND KEEP THIS AT THE BEGINNING OF THE SOURCE CODE ===========================================================//
//=============================================================================================================================================================================//
//=============================================================================================================================================================================//
//
//                                                                  
//
// PROJECT      : TEMPERATURE BOOTH
// ITEM         : TEMPERATURE BOOTH APPLICATION (EXE)
// PROGRAMMER   : Wan Heng Woon
//
// INFORMATION  : THIS APPLICATION SERVES AS A PLATFORM TO KEY-IN TEMPERATURE DATA DURING COVID-19 OUTBREAK. THIS APPLICATION IS MEANT TO BE USED ALONG WITH THE SENSOR 
//                (PROGRAMMED USING ARDUINO CONTROLLER)
//                            
// PERIOD       : PHASE 0 ; 21-25 JUNE   2020  - USER INTERFACCE FOR THE TEMPERATURE DATA KEY-IN ------------------------------------------------------> VER1.0 TempBooth5 
//                PHASE 1 ; 02 SEPTEMBER 2020  - NEW K3 : SENSOR DATA + BADGE ID (API)-----------------------------------------------------------------> VER1.0 TempBooth5 
//                                               POC for the K3 Serial communication and data retrieve
//                                               Merging the K3 Retrieve and adding Badge ID scanning with the previous system
//                          03 SEPTEMBER 2020  - NEW K3 : DATABASE + DATA GRID VIEW + USER INTERFACE --------------------------------------------------> VER1.0 TempBooth5 
//                                               Merging the K3 Retrieve and adding Badge ID scanning with the previous system
//                          04 SEPTEMBER 2020  - NEW K3 : FINAL TOUCH UP; COMPLETED -------------------------------------------------------------------> VER1.0 TempBooth5 
//                                               Finale touch up with new design.              
//=============================================================================================================================================================================//

// if nk key in manual, how   == done, scan card first, then enable the temp box and assign focus, both sensor or manual key in is okay
// if kalau ada problem, how to reset 

using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Data.SqlClient;
using System.Net;
using Newtonsoft.Json;
using System.Drawing.Drawing2D;
using Excel = Microsoft.Office.Interop.Excel;

namespace TempBooth
{
    public partial class MainPage : Form
    {
        SerialPort sp = new SerialPort();
        string desc, name;
        int A, B;
        bool connected = false;
        bool yes = true;
        int C = 0;
        DataTable Dt = new DataTable();

        int total = 0,excel_count=1;
        Timer timer = new Timer();
        Timer timer_2 = new Timer();
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        //Excel.Worksheets sheets = xlApp.Workbooks;
        //xlWorkBook = xlApp.Workbooks.Add(misValue);
        object misValue = System.Reflection.Missing.Value;
        string path="C:\\Users\\Wan Heng Woon\\OneDrive\\Desktop\\csharp-Excel.xls";

        public Form1()
        {
            InitializeComponent();
            timer1.Start();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region define_datagridview
            Dt.Columns.Add("No");
            Dt.Columns.Add("Badge ID");
            Dt.Columns.Add("Name");
            Dt.Columns.Add("Temperature");
            Dt.Columns.Add("Submitted Date Time");
            dataGridView1.DataSource = Dt;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            #endregion

            #region define_blinkingtimer
            timer.Interval = 500;
            timer.Enabled = false;
            timer.Start();
            timer.Tick += new EventHandler(timer_Tick);
            #endregion
            textbox_badgeId.Select();
            label7.Text = DateTime.Now.ToString();
            #region define_wificonnectiontimer
            timer_2.Interval = 500;
            timer_2.Enabled = false;
            timer_2.Start();
            timer_2.Tick += new EventHandler(timer_Tick2);
            #endregion

        }
        //connect to sensor and timerly read the data
        private void timer1_Tick(object sender, EventArgs e)
        {
            #region get_current_comports_and_establish_connection
            try
            {
                Win32DeviceMgmt.GetAllCOMPorts();//var aAdmin = new Win32DeviceMgmt.DeviceInfo();
                var value = Win32DeviceMgmt.GetAllCOMPorts();
                A = value.Count;

                for (int i = 0; i < A; i++)
                {
                    desc = value[i].decsription;
                    name = value[i].name;

                    if (yes == desc.Contains("CH340"))
                    {
                        label14.Text = desc;
                        label6.Text = name;
                        label14.ForeColor = Color.Lime;
                        label6.ForeColor = Color.Lime;
                        label2.ForeColor = Color.Lime;
                        label10.ForeColor = Color.Lime;

                        if (!connected)
                        {
                            sp = new SerialPort(name, 115200, Parity.None, 8, StopBits.One);
                            sp.Open();
                            connected = true;
                            sp.WriteLine("O");
                            timer2.Start();
                            timer1.Stop();
                        }
                        else
                        {
                        }
                    }
                    else
                    {
                        label6.Text = string.Empty;
                        label14.Text = string.Empty;
                    }
                }
            }
            catch
            {
            }
            #endregion
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            #region doublecheck_total_number_of_available_comports
            var value = Win32DeviceMgmt.GetAllCOMPorts();
            B = value.Count;
            if (B != A)
            {
                sp.Dispose();
                connected = false;
                sp.Close();

                label6.Text = string.Empty;
                label14.Text = string.Empty;
                label14.ForeColor = Color.FromArgb(212, 68, 52);
                label6.ForeColor = Color.FromArgb(212, 68, 52);
                label2.ForeColor = Color.FromArgb(212, 68, 52);
                label10.ForeColor = Color.FromArgb(212, 68, 52);
                timer1.Start();
                timer2.Stop();
            }
            #endregion

            #region get_sensor_reading
            try
            {
                var replacements = new[] {
                    new { Find = "T body = ", Replace = "" },
                    new { Find = " C, weak low", Replace = "" },
                    new { Find = " C, weak high", Replace = "" },
                    new { Find = ", weak high", Replace = "" },
                    new { Find = ", weak low", Replace = "" },};

                List<string> temp = new List<string>();
                string a = sp.ReadExisting();
                if (a!=null)
                {
                    string pattern = "\r\n";
                    string[] strNamesArray = Regex.Split(a, pattern);    // Split on hyphens

                    for (int i = 0; i < strNamesArray.Length; i++)
                    {
                        if (strNamesArray[i].Contains("T body"))
                        {
                            temp.Add(strNamesArray[i]);
                        }
                    }

                    if (temp.Count > 0)
                    {
                        if (C == 1)
                        {
                            label17.Visible = true;
                            foreach (string content in temp)
                            {
                                string test = content.Split('=', ',')[2];
                                if (test != " ambience compensate")
                                {
                                    var originalString = content;
                                    foreach (var set in replacements)
                                    {
                                        originalString = originalString.Replace(set.Find, set.Replace);
                                    }
                                    textbox_temperatureValue.Text = originalString.Substring(0, 4);
                                    label17.Text = originalString.Substring(0, 4);
                                }
                                if (test == " ambience compensate")
                                {
                                    textbox_temperatureValue.Text = "Low";
                                    label17.Text = "Low";
                                }
                            }
                        }
                        else
                        {
                            //label9.Visible = true;
                            label17.Visible = false;
                            label3.Visible = false;
                            label1.Visible = false;
                            label16.Text = "";
                            label18.Visible = true;
                            label15.Visible = false;
                        }
                    }
                }
            }
            catch
            {
            }
            #endregion
        }


        //scan badge card and submit data
        private void textbox_badgeId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textbox_badgeId.Text.Count() == 10 || label16.Text.Count() ==10)
                {
                    if (CheckForInternetConnection())
                    {
                        label15.Text = "";
                        label17.Text = "";
                        #region retrieve_details_from_api
                        string jsonsource = @"{""application/json"": [{""SystemName"": ""MySystemName"",""svcmethod"":""GetInfoByCardNo"",""cardNo"":""" + textbox_badgeId.Text + "" + "\"" + "}]}";
                        string jsonresult = CallAPI(jsonsource, @"http://xxxxx", "POST");
                        DataSet data = JsonConvert.DeserializeObject<DataSet>(jsonresult);
                        #endregion

                        #region display_data
                        string badge = data.Tables[0].Rows[0][0].ToString();
                        string name1 = data.Tables[0].Rows[0][1].ToString();
                        label16.Text = badge;
                        label15.Text = name1;
                        label15.Visible = true;
                        label1.Visible = true;
                        label3.Visible = true;
                        //label8.Visible = false;
                        label18.Visible = false;
                        if (label16.Text == "Error : Exception has been thrown by the target of an invocation.")
                        {
                            textbox_badgeId.Clear();
                            label16.Text = "";
                            label15.Text = "";
                            label3.Visible = false;
                            label1.Visible = false;
                            //label8.Visible = true;
                            label18.Visible = true;
                            return;
                        }
                        //textBox2.Focus();
                        //Invoke(new Action(() => textBox2.Text = textBox3.Text));
                        C = 1;
                        //label9.Visible = false;
                        textbox_temperatureValue.Enabled = true;
                        textbox_temperatureValue.Focus();
                        #endregion
                    }
                    else {
                        label15.Text = "";
                        label17.Text = "";
                        label16.Text = textbox_badgeId.Text;
                        label15.Text = "No Internet Connection";
                        label15.Visible = true;
                        label1.Visible = true;
                        label3.Visible = true;
                        //label8.Visible = false;
                        label18.Visible = false;
                        C = 1;
                        textbox_temperatureValue.Enabled = true;
                        textbox_temperatureValue.Focus();
                    }

                }
            }
            catch
            {
                textbox_badgeId.Clear();
                label16.Text = "";
                textbox_temperatureValue.Clear();
                label17.Text = "";
                label18.Visible = true;
                label3.Visible = false;
                label1.Visible = false;
                //label8.Visible = true;
            }
        }

        private static string CallAPI(string json, string posturl, string Method)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(posturl);
            httpWebRequest.Method = Method;

            httpWebRequest.ContentType = "application/json";
            //httpWebRequest.Credentials = CredentialCache.DefaultCredentials;
            httpWebRequest.Credentials = new NetworkCredential("mpdnabila", "MynameisKhan@4", "atmex");
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {

                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }

        private void textbox_temperatureValue_TextChanged(object sender, EventArgs e)
        {
            button_submit.PerformClick();
        }


        private void button_submit_Click(object sender, EventArgs e)
        {           
            if (textbox_temperatureValue.Text.Contains(".") && textbox_temperatureValue.Text.Length == 4)
            {
                total++;
                label17.Visible = true;
                label17.Text = textbox_temperatureValue.Text;
                #region add_into_datgridview
                DataRow dr = Dt.NewRow();
                dr["No"] = total.ToString();
                dr["Badge ID"] = label16.Text;
                dr["Name"] = label15.Text;
                dr["Temperature"] = textbox_temperatureValue.Text;
                dr["Submitted Date Time"] = DateTime.Now.ToString();
                Dt.Rows.Add(dr);
                dataGridView1.DataSource = Dt;
                dataGridView1.Sort(dataGridView1.Columns["Submitted Date Time"], ListSortDirection.Descending);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                if (dataGridView1.Rows.Count > 30)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 1);
                }
                #endregion

                if (CheckForInternetConnection())
                {

                    #region insert_to_database
                    using (SqlConnection con = new SqlConnection(@"Data Source=Server_Name;Initial Catalog=Table_Name; User ID=UserID;Password=Password"))
                    {
                        String query = "INSERT INTO dbo.Temperature_Data (No,[Badge ID],Name,Temperature,[Submitted Date Time] ) VALUES(@No,@ID,@Name,@Temperature,@Time)";
                        using (SqlCommand command = new SqlCommand(query, con))
                        {
                            command.Parameters.AddWithValue("@No", total.ToString());
                            command.Parameters.AddWithValue("@ID", label16.Text);
                            command.Parameters.AddWithValue("@Name", label15.Text);
                            command.Parameters.AddWithValue("@Temperature", textbox_temperatureValue.Text);
                            command.Parameters.AddWithValue("@Time", DateTime.Now.ToString());

                            con.Open();
                            int result = command.ExecuteNonQuery();
                            if (result < 0)
                                Console.WriteLine("Error inserting data into Database! \nDetails information : dbo.SignalTransfer2 at ATMNTS76");
                        }
                    }
                    #endregion

                }
                else
                {
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    if (File.Exists(path))
                    {

                        xlWorkBook = xlApp.Workbooks.Open(@"C:\\Users\\Wan Heng Woon\\OneDrive\\Desktop\\csharp-Excel.xls");
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.Cells[excel_count, 1] = total.ToString();
                        xlWorkSheet.Cells[excel_count, 2] = textbox_badgeId.Text;
                        xlWorkSheet.Cells[excel_count, 3] = textbox_temperatureValue.Text;
                        xlWorkSheet.Cells[excel_count, 4] = DateTime.Now.ToString();
                        //.SaveAs("C:\\Users\\Wan Heng Woon\\OneDrive\\Desktop\\csharp-Excel.xls");
                        xlApp.DisplayAlerts = false;
                        

                    }
                    else
                    {
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[excel_count, 1] = total.ToString();
                        xlWorkSheet.Cells[excel_count, 2] = textbox_badgeId.Text;
                        xlWorkSheet.Cells[excel_count, 3] = textbox_temperatureValue.Text;
                        xlWorkSheet.Cells[excel_count, 4] = DateTime.Now.ToString();
                        xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, misValue, misValue, false);
                        //xlWorkBook.SaveAs(path);

                    }
                    
                }
                #region reset
                //button_reset.PerformClick();
                xlWorkBook.Save();
                Clean();
                excel_count++;
                textbox_badgeId.Clear();
                textbox_temperatureValue.Clear();
                textbox_badgeId.Focus();
                textbox_temperatureValue.Enabled = false;
                C = 0;
                #endregion
            }
            else if (textbox_temperatureValue.Text.Length > 4)
            {
                textbox_temperatureValue.Text = "";
            }
        }



        //reset
        private void button_reset_Click(object sender, EventArgs e)
        {
            textbox_badgeId.Clear();
            label16.Text = "";
            label15.Text = "";
            label17.Text = "";
            label3.Visible = false;
            label1.Visible = false;
            label18.Visible = true;
            textbox_temperatureValue.Clear();
            
            textbox_badgeId.Focus();
            textbox_temperatureValue.Enabled = false;
            C = 0;
        }
        void timer_Tick(object sender, EventArgs e)
        {
            if (label18.ForeColor == Color.FromArgb(247, 222, 220))
                label18.ForeColor = Color.FromArgb(212, 68, 52);
            else
                label18.ForeColor = Color.FromArgb(247, 222, 220);
            
        }

        void timer_Tick2(object sender, EventArgs e)
        {
            if (CheckForInternetConnection())
            {
                //label13.Text = "connected";

                if (File.Exists(path))
                {

                    #region openfile
                    int rw = 0;
                    string[] badgeID = new string[2000];
                    string[] total_ID = new string[2000];
                    string[] temperature = new string[2000];
                    string[] date = new string[2000];
                    Excel.Range range;
                    xlWorkBook = xlApp.Workbooks.Open(@"" + path + "");
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    range = xlWorkSheet.UsedRange;
                    rw = range.Rows.Count;
                    for (int i = 1; i <= rw; i++)
                    {
                        total_ID[i] = range.Cells[i, 1].Value2.ToString();

                    }
                    for (int i = 1; i <= rw; i++)
                    {
                        badgeID[i] = range.Cells[i, 2].Value2.ToString();
                    }
                    for (int i = 1; i <= rw; i++)
                    {
                        temperature[i] = range.Cells[i, 3].Value2.ToString();
                    }
                    for (int i = 1; i <= rw; i++)
                    {
                        date[i] = range.Cells[i, 4].Value2.ToString();
                    }
                    #endregion
                    /*#region retrieve_details_from_api_and_save_it_to_data_base
                    for (int i = 0; i < rw; i++)
                    {
                        string jsonsource = @"{""application/json"": [{""SystemName"": ""MySystemName"",""svcmethod"":""GetInfoByCardNo"",""cardNo"":""" + badgeID[i] + "" + "\"" + "}]}";
                        string jsonresult = CallAPI(jsonsource, @"http://xxxx", "POST");
                        DataSet data = JsonConvert.DeserializeObject<DataSet>(jsonresult);
                        string badge = data.Tables[0].Rows[0][0].ToString();
                        string name1 = data.Tables[0].Rows[0][1].ToString();
                        using (SqlConnection con = new SqlConnection(@"Data Source=ServerName;Initial Catalog=TableName; User ID=userID;Password=password"))
                        {
                            String query = "INSERT INTO dbo.Temperature_Data (No,[Badge ID],Name,Temperature,[Submitted Date Time] ) VALUES(@No,@ID,@Name,@Temperature,@Time)";
                            using (SqlCommand command = new SqlCommand(query, con))
                            {
                                command.Parameters.AddWithValue("@No", total_ID[i].ToString());
                                command.Parameters.AddWithValue("@ID", badge);
                                command.Parameters.AddWithValue("@Name", name1);
                                command.Parameters.AddWithValue("@Temperature", temperature[i]);
                                command.Parameters.AddWithValue("@Time", date[i]);

                                con.Open();
                                int result = command.ExecuteNonQuery();
                                if (result < 0)
                                    Console.WriteLine("Error inserting data into Database! \nDetails information : dbo.SignalTransfer2 at ATMNTS76");
                            }
                        }
                    }

                    #endregion*/
                    Clean();
                    File.Delete(path);
                }
            }


            else
            {
                //label13.Text = "disconnected";
            }
        }

        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        private void Clean()
        {    
            //ReleaseComObject(xlWorkSheet);
            xlApp.Application.Quit();
            xlApp.Quit();
            //xlWorkBook.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    public class Win32DeviceMgmt
    {
        private const UInt32 DIGCF_PRESENT = 0x00000002;
        private const UInt32 DIGCF_DEVICEINTERFACE = 0x00000010;
        private const UInt32 SPDRP_DEVICEDESC = 0x00000000;
        private const UInt32 DICS_FLAG_GLOBAL = 0x00000001;
        private const UInt32 DIREG_DEV = 0x00000001;
        private const UInt32 KEY_QUERY_VALUE = 0x0001;
        private const string GUID_DEVINTERFACE_COMPORT = "86E0D1E0-8089-11D0-9CE4-08003E301F73";

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public Int32 cbSize;
            public Guid ClassGuid;
            public Int32 DevInst;
            public UIntPtr Reserved;
        };

        [DllImport("setupapi.dll")]
        private static extern Int32 SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("setupapi.dll")]
        private static extern bool SetupDiEnumDeviceInfo(IntPtr DeviceInfoSet, Int32 MemberIndex, ref SP_DEVINFO_DATA DeviceInterfaceData);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetupDiGetDeviceRegistryProperty(IntPtr deviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData,
            uint property, out UInt32 propertyRegDataType, StringBuilder propertyBuffer, uint propertyBufferSize, out UInt32 requiredSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, UInt32 iEnumerator, IntPtr hParent, UInt32 nFlags);

        [DllImport("Setupapi", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetupDiOpenDevRegKey(IntPtr hDeviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, uint scope,
            uint hwProfile, uint parameterRegistryValueKind, uint samDesired);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueExW", SetLastError = true)]
        private static extern int RegQueryValueEx(IntPtr hKey, string lpValueName, int lpReserved, out uint lpType,
            StringBuilder lpData, ref uint lpcbData);

        [DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int RegCloseKey(IntPtr hKey);

        [DllImport("kernel32.dll")]
        private static extern Int32 GetLastError();

        public struct DeviceInfo
        {
            public string name;
            public string decsription;
        }

        public static List<DeviceInfo> GetAllCOMPorts()
        {
            Guid guidComPorts = new Guid(GUID_DEVINTERFACE_COMPORT);
            IntPtr hDeviceInfoSet = SetupDiGetClassDevs(
                ref guidComPorts, 0, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
            if (hDeviceInfoSet == IntPtr.Zero)
            {
                throw new Exception("Failed to get device information set for the COM ports");
            }

            try
            {
                List<DeviceInfo> devices = new List<DeviceInfo>();
                Int32 iMemberIndex = 0;
                while (true)
                {
                    SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                    deviceInfoData.cbSize = Marshal.SizeOf(typeof(SP_DEVINFO_DATA));
                    bool success = SetupDiEnumDeviceInfo(hDeviceInfoSet, iMemberIndex, ref deviceInfoData);
                    if (!success)
                    {
                        // No more devices in the device information set
                        break;
                    }

                    DeviceInfo deviceInfo = new DeviceInfo();
                    deviceInfo.name = GetDeviceName(hDeviceInfoSet, deviceInfoData);
                    deviceInfo.decsription = GetDeviceDescription(hDeviceInfoSet, deviceInfoData);
                    devices.Add(deviceInfo);

                    iMemberIndex++;
                }
                return devices;
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(hDeviceInfoSet);
            }
        }

        private static string GetDeviceName(IntPtr pDevInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            IntPtr hDeviceRegistryKey = SetupDiOpenDevRegKey(pDevInfoSet, ref deviceInfoData,
                DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_QUERY_VALUE);
            if (hDeviceRegistryKey == IntPtr.Zero)
            {
                throw new Exception("Failed to open a registry key for device-specific configuration information");
            }

            StringBuilder deviceNameBuf = new StringBuilder(256);
            try
            {
                uint lpRegKeyType;
                uint length = (uint)deviceNameBuf.Capacity;
                int result = RegQueryValueEx(hDeviceRegistryKey, "PortName", 0, out lpRegKeyType, deviceNameBuf, ref length);
                if (result != 0)
                {
                    throw new Exception("Can not read registry value PortName for device " + deviceInfoData.ClassGuid);
                }
            }
            finally
            {
                RegCloseKey(hDeviceRegistryKey);
            }

            string deviceName = deviceNameBuf.ToString();
            return deviceName;
        }

        private static string GetDeviceDescription(IntPtr hDeviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            StringBuilder descriptionBuf = new StringBuilder(256);
            uint propRegDataType;
            uint length = (uint)descriptionBuf.Capacity;
            bool success = SetupDiGetDeviceRegistryProperty(hDeviceInfoSet, ref deviceInfoData, SPDRP_DEVICEDESC,
                out propRegDataType, descriptionBuf, length, out length);
            if (!success)
            {
                throw new Exception("Can not read registry value PortName for device " + deviceInfoData.ClassGuid);
            }
            string deviceDescription = descriptionBuf.ToString();
            return deviceDescription;
        }

    }
    
}
