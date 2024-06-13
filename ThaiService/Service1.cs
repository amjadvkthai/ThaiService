using System;
using System.Net.Http;
using System.ServiceProcess;
using System.Timers;
using System.Diagnostics;
using System.Text.Json;
using System.Management;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Security.Principal;

namespace UserInformationService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        private readonly HttpClient httpClient;
        private const string LogFilePath = @"C:\Program Files\ThaiM\log.txt";
        private const string LastSentFilePath = @"C:\Program Files\ThaiM\last_sent.txt";
        private const int SendAllDataToApi = 128; // Custom command number
        private DateTime lastSent;
        private TimeSpan intervalCheck = TimeSpan.FromHours(1); // 1 hour check interval

        public Service1()
        {
            InitializeComponent();
            httpClient = new HttpClient();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer(intervalCheck.TotalMilliseconds); // Check every 1 hour
            timer.Elapsed += OnElapsedTime;
            timer.AutoReset = true;
            timer.Enabled = true;
            Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));

            lastSent = ReadLastSentTime();

            // Check if it's been 1 hour since the last sent time, send data if needed
            Task.Run(() => CheckAndSendDataImmediately());
        }

        private async void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            await CheckAndSendDataImmediately();
        }

        private async Task CheckAndSendDataImmediately()
        {
            if ((DateTime.Now - lastSent).TotalHours >= 1 || lastSent == DateTime.MinValue)
            {
                await CollectAndSendData();
            }
        }

        private async Task CollectAndSendData()
        {
            var (userName, userSid, hostName) = GetLoggedInUserWithSID();

            var systemInfo = GatherSystemInformation();
            var driveInfos = GatherDriveInformation();
            var installedSoftware = GetInstalledSoftware();
            var networkDevices = GatherNetworkDevices();
            var listeningPorts = GatherListeningPorts();

            var softwareNames = new List<string>();
            var softwareVersions = new List<string>();

            foreach (var software in installedSoftware)
            {
                var parts = software.Split(new[] { " - " }, StringSplitOptions.None);
                if (parts.Length == 2)
                {
                    softwareNames.Add(parts[0]);
                    softwareVersions.Add(parts[1]);
                }
                else
                {
                    softwareNames.Add(parts[0]);
                    softwareVersions.Add("N/A");
                }
            }

            var dataList = new List<object>();

            var data = new
            {
                userID = userSid,
                userName = userName,
                hostName = hostName,
                os = systemInfo.OS,
                version = systemInfo.Version,
                servicePack = systemInfo.ServicePack,
                physicalMemory = systemInfo.PhysicalMemory,
                freeMemory = systemInfo.FreeMemory,
                processor = systemInfo.Processor,
                cores = systemInfo.Cores,
                clockSpeed = systemInfo.ClockSpeed,
                driveName = driveInfos.DriveNames,
                totalSize = driveInfos.TotalSizes,
                usedSpace = driveInfos.UsedSpaces,
                freeSpace = driveInfos.FreeSpaces,
                softwareName = string.Join(",", softwareNames),
                softwareVersion = string.Join(",", softwareVersions),
                networkDeviceName = networkDevices.DeviceNames,
                macAddress = networkDevices.MacAddresses,
                protocol = listeningPorts.Protocols,
                localAddress = listeningPorts.LocalAddresses,
                remoteAddress = listeningPorts.RemoteAddresses,
                state = listeningPorts.States,
                date = DateTime.Now
            };

            dataList.Add(data);

            var json = JsonSerializer.Serialize(dataList);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync("https://itapp.teamthai.org/api/api/UserInfo", content);
                if (response.IsSuccessStatusCode)
                {
                    lastSent = DateTime.Now; // Update last sent time
                    SaveLastSentTime(lastSent);

                    // LogData(json); // Log the data to log.txt
                }
                else
                {
                    throw new HttpRequestException($"Failed to send data. Status code: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error: {ex.Message}", EventLogEntryType.Error);
                LogError(ex.Message);
            }
        }

        private DateTime ReadLastSentTime()
        {
            if (File.Exists(LastSentFilePath))
            {
                string lastSentString = File.ReadAllText(LastSentFilePath);
                if (DateTime.TryParse(lastSentString, out DateTime lastSentTime))
                {
                    return lastSentTime;
                }
            }
            return DateTime.MinValue;
        }

        private void SaveLastSentTime(DateTime lastSentTime)
        {
            try
            {
                File.WriteAllText(LastSentFilePath, lastSentTime.ToString());
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error saving last sent time: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error saving last sent time: {ex.Message}");
            }
        }

        public (string UserName, string UserSid, string HostName) GetLoggedInUserWithSID()
        {
            string userSid = null;
            string hostName = System.Environment.MachineName;
            string userName = null;

            try
            {
                // Get the currently logged in user
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
                ManagementObjectCollection collection = searcher.Get();

                foreach (ManagementObject obj in collection)
                {
                    userName = obj["UserName"]?.ToString();
                    break;
                }

                if (!string.IsNullOrEmpty(userName))
                {
                    string domainName = userName.Split('\\')[0];
                    string accountName = userName.Split('\\')[1];

                    NTAccount account = new NTAccount(domainName, accountName);
                    SecurityIdentifier sid = (SecurityIdentifier)account.Translate(typeof(SecurityIdentifier));
                    userSid = sid.ToString();
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error retrieving logged in user: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error retrieving logged in user: {ex.Message}");
            }

            return (userName, userSid, hostName);
        }

        private (string OS, string Version, string ServicePack, string PhysicalMemory, string FreeMemory, string Processor, string Cores, string ClockSpeed) GatherSystemInformation()
        {
            string os = "N/A", version = "N/A", servicePack = "N/A", physicalMemory = "N/A", freeMemory = "N/A", processor = "N/A", cores = "N/A", clockSpeed = "N/A";

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject osObj in searcher.Get())
                {
                    os = osObj["Caption"]?.ToString() ?? "N/A";
                    version = osObj["Version"]?.ToString() ?? "N/A";
                    servicePack = osObj["ServicePackMajorVersion"]?.ToString() ?? "N/A";
                    physicalMemory = ((Convert.ToInt64(osObj["TotalVisibleMemorySize"]) / 1024.0 / 1024).ToString("F2") + " GB");
                    freeMemory = ((Convert.ToInt64(osObj["FreePhysicalMemory"]) / 1024.0 / 1024).ToString("F2") + " GB");
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (ManagementObject cpu in searcher.Get())
                {
                    processor = cpu["Name"]?.ToString() ?? "N/A";
                    cores = cpu["NumberOfCores"]?.ToString() ?? "N/A";
                    clockSpeed = cpu["MaxClockSpeed"]?.ToString() ?? "N/A" + " MHz";
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error gathering system information: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error gathering system information: {ex.Message}");
            }

            return (os, version, servicePack, physicalMemory, freeMemory, processor, cores, clockSpeed);
        }

        private (string DriveNames, string TotalSizes, string UsedSpaces, string FreeSpaces) GatherDriveInformation()
        {
            var driveNames = new List<string>();
            var totalSizes = new List<string>();
            var usedSpaces = new List<string>();
            var freeSpaces = new List<string>();

            try
            {
                DriveInfo[] allDrives = DriveInfo.GetDrives();

                foreach (DriveInfo drive in allDrives)
                {
                    if (drive.IsReady)
                    {
                        long totalSize = drive.TotalSize;
                        long freeSpace = drive.TotalFreeSpace;
                        long usedSpace = totalSize - freeSpace;

                        driveNames.Add(drive.Name);
                        totalSizes.Add((totalSize / 1024 / 1024).ToString());
                        usedSpaces.Add((usedSpace / 1024 / 1024).ToString());
                        freeSpaces.Add((freeSpace / 1024 / 1024).ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error gathering drive information: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error gathering drive information: {ex.Message}");
            }

            return (string.Join(",", driveNames), string.Join(",", totalSizes), string.Join(",", usedSpaces), string.Join(",", freeSpaces));
        }

        private List<string> GetInstalledSoftware()
        {
            List<string> softwareList = new List<string>();
            try
            {
                GetInstalledSoftwareFromRegistry(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", Microsoft.Win32.RegistryView.Registry64, softwareList);
                GetInstalledSoftwareFromRegistry(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", Microsoft.Win32.RegistryView.Registry32, softwareList);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error retrieving installed software: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error retrieving installed software: {ex.Message}");
            }
            return softwareList;
        }

        private void GetInstalledSoftwareFromRegistry(string registryKey, Microsoft.Win32.RegistryView registryView, List<string> softwareList)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey baseKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, registryView))
                {
                    using (Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(registryKey))
                    {
                        foreach (string subkey_name in key.GetSubKeyNames())
                        {
                            using (Microsoft.Win32.RegistryKey subkey = key.OpenSubKey(subkey_name))
                            {
                                string displayName = subkey.GetValue("DisplayName") as string;
                                string displayVersion = subkey.GetValue("DisplayVersion") as string;
                                string publisherName = subkey.GetValue("Publisher") as string;

                                if (!string.IsNullOrEmpty(displayName) && !string.IsNullOrEmpty(publisherName) && !publisherName.Equals("Microsoft Corporation", StringComparison.OrdinalIgnoreCase))
                                {
                                    softwareList.Add($"{displayName} - {displayVersion} - {publisherName}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error accessing registry key: {registryKey} in view: {registryView}. Error: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error accessing registry key: {registryKey} in view: {registryView}. Error: {ex.Message}");
            }
        }

        private (string DeviceNames, string MacAddresses) GatherNetworkDevices()
        {
            var deviceNames = new List<string>();
            var macAddresses = new List<string>();

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2");
                foreach (ManagementObject adapter in searcher.Get())
                {
                    string name = adapter["Name"]?.ToString();
                    string macAddress = adapter["MACAddress"]?.ToString();

                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(macAddress))
                    {
                        deviceNames.Add(name);
                        macAddresses.Add(macAddress);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing network devices information to log file: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error writing network devices information to log file: {ex.Message}");
            }

            return (string.Join(",", deviceNames), string.Join(",", macAddresses));
        }

        private (string Protocols, string LocalAddresses, string RemoteAddresses, string States) GatherListeningPorts()
        {
            var protocols = new List<string>();
            var localAddresses = new List<string>();
            var remoteAddresses = new List<string>();
            var states = new List<string>();

            try
            {
                Process netStatProcess = new Process();
                netStatProcess.StartInfo.FileName = "netstat.exe";
                netStatProcess.StartInfo.Arguments = "-an";
                netStatProcess.StartInfo.UseShellExecute = false;
                netStatProcess.StartInfo.RedirectStandardOutput = true;
                netStatProcess.Start();

                string output = netStatProcess.StandardOutput.ReadToEnd();
                netStatProcess.WaitForExit();

                var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                var filteredLines = lines.Where(line => line.Contains("LISTENING") || line.Contains("UDP"));

                foreach (var line in filteredLines)
                {
                    var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 4)
                    {
                        protocols.Add(parts[0]);
                        localAddresses.Add(parts[1]);
                        remoteAddresses.Add(parts[2]);
                        states.Add(parts[3]);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UserInformationService", $"Error writing listening ports information to log file: {ex.Message}", EventLogEntryType.Error);
                LogError($"Error writing listening ports information to log file: {ex.Message}");
            }

            return (string.Join(",", protocols), string.Join(",", localAddresses), string.Join(",", remoteAddresses), string.Join(",", states));
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

        protected override void OnCustomCommand(int command)
        {
            base.OnCustomCommand(command);
            if (command == SendAllDataToApi)
            {
                Task.Run(() => CollectAndSendData()); // Call the function to send all data to API
            }
        }

        protected override void OnStop()
        {
            timer.Enabled = false;
        }
    }
}
