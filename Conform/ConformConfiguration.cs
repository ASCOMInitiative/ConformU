using Blazorise.Utilities;
using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ConformU
{
    [DefaultMember("Settings")]
    public class ConformConfiguration : IDisposable, IConformConfiguration
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
                // Create a new settings file with default values in case the supplied file cannot be used
                settings = new();

                // Get the full settings file name including path
                if (string.IsNullOrEmpty(configurationFile)) // No override settings fie has been specified so use the application default settings file
                {
                    string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), FOLDER_NAME);
                    SettingsFileName = Path.Combine(folderName, SETTINGS_FILENAME);
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Settings folder: {folderName}, Settings file: {SettingsFileName}");
                }
                else // An override settings file has been supplied so use it instead of the default settings file
                {
                    SettingsFileName = configurationFile;
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Settings file: {SettingsFileName}");
                }

                // Load the values in the settings file if it exists
                if (File.Exists(SettingsFileName)) // Settings file exists
                {
                    // Read the file contents into a string
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, "File exists and read OK");
                    string serialisedSettings = File.ReadAllText(SettingsFileName);
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Serialised settings: {serialisedSettings}");

                    // Set de-serialisation options
                    JsonSerializerOptions options = new()
                    {
                        PropertyNameCaseInsensitive = true // Ignore incorrect element name casing
                    };
                    options.Converters.Add(new JsonStringEnumConverter()); // Accept both string member names and integer member values as valid for enum elements.

                    // De-serialise the settings string into a Settings object
                    settings = JsonSerializer.Deserialize<Settings>(serialisedSettings, options);

                    // Test whether the retrieved settings match the requirements of this version of ConformU
                    if (settings.SettingsCompatibilityVersion == Settings.SETTINGS_COMPATIBILTY_VERSION) // Version numbers match so all is well
                    {
                        Status = $"Settings read OK.";
                    }
                    else // Version numbers don't match so reset to defaults
                    {
                        int originalSettingsCompatibilityVersion=0;
                        try
                        {
                            originalSettingsCompatibilityVersion = settings.SettingsCompatibilityVersion;

                            // Rename the current settings file to preserve it
                            string badVersionSettingsFileName = $"{SettingsFileName}.badversion";
                            File.Delete(badVersionSettingsFileName);
                            File.Move(SettingsFileName, $"{badVersionSettingsFileName}");

                            settings = new();
                            PersistSettings(settings);
                            Status = $"The current settings version: {originalSettingsCompatibilityVersion} does not match the required version: {Settings.SETTINGS_COMPATIBILTY_VERSION}. Settings reset to default values and original settings file renamed to {badVersionSettingsFileName}.";
                        }
                        catch (Exception ex2)
                        {
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error persisting new Conform settings file: {ex2}");
                            Status = $"The current settings version:{originalSettingsCompatibilityVersion} does not match the required version: {Settings.SETTINGS_COMPATIBILTY_VERSION} but the new settings could not be persisted: {ex2.Message}.";
                        }
                    }
                }
                else // Settings file does not exist
                {
                    TL.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Configuration file does not exist, initialising new file: {SettingsFileName}");
                    PersistSettings(settings);
                    Status = $"Settings set to defaults on first time use.";
                }
            }
            catch (JsonException ex1)
            {
                // There was an exception when parsing the settings file so report it and persist new default values
                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error parsing Conform settings file: {ex1}");
                try
                {
                    string corruptSettingsFileName = $"{SettingsFileName}.corrupt";
                    File.Delete(corruptSettingsFileName);
                    File.Move(SettingsFileName, $"{corruptSettingsFileName}");
                    Status = $"The settings file is corrupt ({ex1.Message}), settings have been reset to default values. The corrupt file has been renamed to {SettingsFileName}.corrupt";
                }
                catch (Exception ex2)
                {
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error persisting Conform settings file: {ex2}");
                    Status = $"The settings file is corrupt ({ex1.Message}), but default settings could not be saved to the file system: {ex2.Message}. Default values are in effect.";
                }
            }
            catch (Exception ex)
            {
                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, ex.ToString());
                Status = $"Exception reading the settings file, default values are in use.";
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
                    if (settings.AlpacaDevice.AscomDeviceType == null) issueList += $"\r\nAlpaca device type is not set";

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

                // Set the version number of this settings file
                settingsToPersist.SettingsCompatibilityVersion = Settings.SETTINGS_COMPATIBILTY_VERSION;

                TL?.LogMessage("PersistSettings", MessageLevel.Debug, $"Settings file: {SettingsFileName}");

                JsonSerializerOptions options = new()
                {
                    WriteIndented = true
                };
                options.Converters.Add(new JsonStringEnumConverter());

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