using System;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace ConformU
{
    [DefaultMember("Settings")]
    public class ConformConfiguration : IDisposable
    {
        private const string FOLDER_NAME = "conform"; // Folder name underneath the local application data folder
        private const string SETTINGS_FILENAME = "conform.settings"; // Settings file name

        private ConformLogger TL;
        Settings settings;
        private bool disposedValue;

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
                    SettingsFileName = Path.Combine(folderName, SETTINGS_FILENAME);
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Settings folder: {folderName}, Settings file: {SettingsFileName}");
                }
                else
                {
                    SettingsFileName = configurationFile;
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Settings file: {SettingsFileName}");
                }


                if (File.Exists(SettingsFileName))
                {
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, "File exists and read OK");
                    string serialisedSettings = File.ReadAllText(SettingsFileName);
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Serialised settings: {serialisedSettings}");

                    settings = JsonSerializer.Deserialize<Settings>(serialisedSettings, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    Status = "Settings read OK";
                }
                else
                {
                    TL.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Configuration file does not exist, initialising new file: {SettingsFileName}");
                    PersistSettings(settings);
                    Status = "Settings set to defaults on first time use.";
                }
            }
            catch (JsonException ex)
            {
                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error parsing Conform settings file: {ex.Message}");
                Status = "Settings file corrupted, please reset to default values";
            }
            catch (Exception ex)
            {
                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, ex.ToString());
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
            TL?.LogMessage("Save", MessageLevel.Debug, "Persisting settings to settings file");
            PersistSettings(settings);
            Status = $"Settings saved at {DateTime.Now:HH:mm:ss.f}.";

            // Raise configuration has changed event
            if (ConfigurationChanged is not null)
            {
                EventArgs args = new();
                TL?.LogMessage("Save", MessageLevel.Debug, "About to call configuration changed event handler");
                ConfigurationChanged(this, args);
                TL?.LogMessage("Save", MessageLevel.Debug, "Returned from configuration changed event handler");
            }
        }

        internal void RaiseUiHasChangedEvent()
        {
            if (UiHasChanged is not null)
            {
                EventArgs args = new();
                TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "About to call UI has changed event handler");
                UiHasChanged(this, args);
                TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "Returned from UI has changed event handler");
            }
        }

        public void Reset()
        {
            TL?.LogMessage("Reset", MessageLevel.Debug, "Resetting settings file to default values");
            settings = new();
            PersistSettings(settings);
            Status = $"Settings reset at {DateTime.Now:HH:mm:ss.f}.";
            RaiseUiHasChangedEvent();

        }

        public delegate void MessageEventHandler(object sender, MessageEventArgs e);

        public event EventHandler ConfigurationChanged;

        public event EventHandler UiHasChanged;

        public Settings Settings
        {
            get { return settings; }
        }

        /// <summary>
        /// Text message describing any issues found when validating the settings
        /// </summary>
        public string Status { get; private set; }

        public string SettingsFileName { get; private set; }

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

                TL?.LogMessage("PersistSettings", MessageLevel.Debug, $"Settings file: {SettingsFileName}");

                JsonSerializerOptions options = new()
                {
                    WriteIndented = true
                };
                string serialisedSettings = JsonSerializer.Serialize<Settings>(settingsToPersist, options);

                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFileName));
                File.WriteAllText(SettingsFileName, serialisedSettings);
            }
            catch (Exception ex)
            {
                TL.LogMessage("PersistSettings", MessageLevel.Debug, ex.ToString());
            }

        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Console.WriteLine("ConformConfiguration.Dispose()...");
                    TL = null;
                    settings = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ConformConfiguration()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}