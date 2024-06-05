using System;
using System.Net.Http;
using System.ServiceProcess;
using System.Timers;
using System.Diagnostics;
using System.Text.Json;
using System.Management;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UserInformationService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        private readonly HttpClient httpClient;
        private const string LogFilePath = @"C:\UserInformationService\log.txt";
        private const string UserLogFilePath = @"C:\UserInformationService\user_log.txt";
        private const int LogNetworkDevicesCommand = 128; // Custom command number

        public Service1()
        {
            InitializeComponent();
            httpClient = new HttpClient();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer(3000); // 3 seconds
            timer.Elapsed += OnElapsedTime;
            timer.AutoReset = true;
            timer.Enabled = true;
            Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
            Directory.CreateDirectory(Path.GetDirectoryName(UserLogFilePath));
        }

        private async void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            string userName = GetLoggedInUserWithSID(out string userSid);

            var data = new
            {
                UserName = userName,
                UserSID = userSid,
                RecordedAt = DateTime.Now
            };

            var json = JsonSerializer.Serialize(data);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            try
            {
                await httpClient.PostAsync("http://10.200.40.25:85/User/AddUserInfo", content);
                LogUserInformation(userName, userSid);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error: {ex.Message}", EventLogEntryType.Error);
                LogError(ex.Message);
            }
        }

        public string GetLoggedInUserWithSID(out string userSid)
        {
            userSid = null;
            try
            {
                string userName = null;
                ManagementScope scope = new ManagementScope("\\\\.\\root\\cimv2");
                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_LogonSession WHERE LogonType = 2");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                ManagementObjectCollection queryCollection = searcher.Get();

                foreach (ManagementObject logon in queryCollection)
                {
                    foreach (ManagementObject account in logon.GetRelated("Win32_Account"))
                    {
                        userName = account["Name"] + "\\" + account["Domain"];
                        userSid = account["SID"].ToString();
                        return userName;
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error retrieving logged in user: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error retrieving logged in user: {ex.Message}");
            }

            return null;
        }

        private void LogUserInformation(string userName, string userSid)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(UserLogFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now}: UserName: {userName}, UserSID: {userSid}");
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing user information to log file: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error writing user information to log file: {ex.Message}");
            }
        }

        private void LogNetworkDevices()
        {
            try
            {
                string networkDevicesLogFilePath = @"C:\UserInformationService\network_devices_log.txt";
                Directory.CreateDirectory(Path.GetDirectoryName(networkDevicesLogFilePath));

                using (StreamWriter sw = new StreamWriter(networkDevicesLogFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now}: Connected Network Devices:");

                    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2");
                    foreach (ManagementObject adapter in searcher.Get())
                    {
                        string name = adapter["Name"]?.ToString();
                        string macAddress = adapter["MACAddress"]?.ToString();

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(macAddress))
                        {
                            sw.WriteLine($"Name: {name}, MAC Address: {macAddress}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing network devices information to log file: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error writing network devices information to log file: {ex.Message}");
            }
        }

        private void LogData(string data)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now}: {data}");
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing to log file: {ex.Message}", EventLogEntryType.Error);
            }
        }

        private void LogError(string error)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now}: ERROR: {error}");
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing to log file: {ex.Message}", EventLogEntryType.Error);
            }
        }

        protected override void OnCustomCommand(int command)
        {
            base.OnCustomCommand(command);
            if (command == LogNetworkDevicesCommand)
            {
                LogNetworkDevices();
            }
        }

        protected override void OnStop()
        {
            timer.Enabled = false;
        }
    }
}
