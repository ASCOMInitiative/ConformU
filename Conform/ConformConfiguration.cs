using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace ConformU
{
    [DefaultMember("Settings")]
    public class ConformConfiguration
    {
        private const string FOLDER_NAME = "conform"; // Folder name underneath the local application data folder
        private const string SETTINGS_FILENAME = "conform.settings"; // Settings file name

        private readonly ConformLogger TL;
        Settings settings;
        readonly string fileSettingsFileName;

        /// <summary>
        /// Create a Configuration management instance and load the current settings
        /// </summary>
        /// <param name="logger">Data logger instance.</param>
        public ConformConfiguration(ConformLogger logger, string configurationFile)
        {
            TL = logger;

            try
            {
                settings = new();

                if (string.IsNullOrEmpty(configurationFile))
                {
                    string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), FOLDER_NAME);
                    fileSettingsFileName = Path.Combine(folderName, SETTINGS_FILENAME);
                    TL.LogDebug("ConformConfiguration", $"Settings folder: {folderName}, Settings file: {fileSettingsFileName}");
                }
                else
                {
                    fileSettingsFileName = configurationFile;
                    TL.LogDebug("ConformConfiguration", $"Settings file: {fileSettingsFileName}");
                }


                if (File.Exists(fileSettingsFileName))
                {
                    TL.LogDebug("ConformConfiguration", "File exists and read OK");
                    string serialisedSettings = File.ReadAllText(fileSettingsFileName);
                    TL.LogDebug("ConformConfiguration", $"Serialised settings: {serialisedSettings}");

                    settings = JsonSerializer.Deserialize<Settings>(serialisedSettings, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    Status = "Settings read successfully";
                }
                else
                {
                    TL.LogMessage("ConformConfiguration", $"Configuration file does not exist, initialising new file: {fileSettingsFileName}");
                    PersistSettings(settings);
                    Status = "Settings set to defaults on first time use.";
                }
            }
            catch (JsonException ex)
            {
                TL.LogDebug("ConformConfiguration", $"Error parsing Conform settings file: {ex.Message}");
                Status = "Settings file corrupted, please reset to default values";
            }
            catch (Exception ex)
            {
                TL.LogDebug("ConformConfiguration", ex.ToString());
                Status = "Exception reading settings, default values are in use.";
            }
        }

        /// <summary>
        /// Validate the settings values
        /// </summary>
        public string Validate()
        {
            string issueList = ""; // Initialise to an empty string indicating that all is OK

            if (settings.DeviceName == Globals.NO_DEVICE_SELECTED) issueList += "\r\nNo device has been selected.";
            if ((settings.DeviceTechnology != DeviceTechnology.Alpaca) & (settings.DeviceTechnology != DeviceTechnology.COM)) issueList += $"\r\nTechnology type is not Alpaca or COM: '{settings.DeviceTechnology}'";

            switch (settings.DeviceTechnology)
            {
                case DeviceTechnology.Alpaca:
                    if (string.IsNullOrEmpty(settings.AlpacaDevice.AscomDeviceType)) issueList += $"\r\nAlpaca device type is not set";

                    if (string.IsNullOrEmpty(settings.AlpacaDevice.IpAddress)) issueList += $"\r\nAlpaca device IP address is not set";

                    if (settings.AlpacaDevice.IpPort == 0) issueList += $"\r\nAlpaca device IP Port is not set.";
                    if (settings.AlpacaDevice.InterfaceVersion == 0) issueList += $"\r\nAlpaca interface version is not set";
                    break;

                case DeviceTechnology.COM:
                    if (settings.ComDevice.ProgId == "") return "CurrentComDevice.ProgId  is empty.";

                    break;
            }
            issueList = issueList.Trim();

            return issueList;
        }

        /// <summary>
        /// Persist current settings
        /// </summary>
        public void Save()
        {
            TL.LogDebug("Save", "persisting settings to settings file");
            PersistSettings(settings);
            Status = $"Settings saved at {DateTime.Now:HH:mm:ss.f}.";

            RaiseConfigurationChnagedEvent();
        }

        private void RaiseConfigurationChnagedEvent()
        {
            if (ConfigurationChanged is not null)
            {
                EventArgs args = new();
                TL.LogDebug("RaiseConfigurationChnagedEvent", "About to call configuration changed event handler");
                ConfigurationChanged(this, args);
                TL.LogDebug("RaiseConfigurationChnagedEvent", "Returned from configuration changed event handler");
            }
        }

        public void Reset()
        {
            TL.LogDebug("Reset", "Resetting settings file to default values");
            settings = new();
            PersistSettings(settings);
            Status = $"Settings reset at {DateTime.Now:HH:mm:ss.f}.";
            RaiseConfigurationChnagedEvent();

        }

        public delegate void MessageEventHandler(object sender, MessageEventArgs e);

        public event EventHandler ConfigurationChanged;

        public Settings Settings
        {
            get { return settings; }
        }

        /// <summary>
        /// Text message describing any issues found when validating the settings
        /// </summary>
        public string Status { get; private set; }

        /// <summary>
        /// Enables discovery process messages in the Conform log
        /// </summary>
        public bool DebugDiscovery
        {
            set
            {
                settings.DebugAlpacaDiscovery = value;
            }
        }

        /// <summary>
        /// Flag indicating whether de-serialisation of Alpaca JSON responses is sensitive to case of JSON element names
        /// </summary>
        /// <remarks>If Set TRUE JSON element names that are incorrectly cased will be ignored. If FALSE they will be accepted.</remarks>
        public bool JsonDeserialisationIsCaseSensitive { get; private set; }

        private void PersistSettings(Settings settingsToPersist)
        {
            try
            {
                if (settingsToPersist is null)
                {
                    throw new ArgumentNullException(nameof(settingsToPersist));
                }

                TL.LogDebug("PersistSettings", $"Settings file: {fileSettingsFileName}");

                JsonSerializerOptions options = new()
                {
                    WriteIndented = true
                };
                string serialisedSettings = JsonSerializer.Serialize<Settings>(settingsToPersist, options);

                Directory.CreateDirectory(Path.GetDirectoryName(fileSettingsFileName));
                File.WriteAllText(fileSettingsFileName, serialisedSettings);
            }
            catch (Exception ex)
            {
                TL.LogMessage("PersistSettings", ex.ToString());
            }

        }
    }
}