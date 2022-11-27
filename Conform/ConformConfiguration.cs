using Blazorise.Utilities;
using System;
using System.Collections.Generic;
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
        readonly int settingsFileVersion;

        #region Initialiser and Dispose

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
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, "File exists, about to read it...");
                    string serialisedSettings = File.ReadAllText(SettingsFileName);
                    TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Serialised settings: \r\n{serialisedSettings}");

                    // Make a basic check to see if this file is a beta / pre-release version that doesn't have a version number. If so replace with a new version
                    if (!serialisedSettings.Contains("\"SettingsCompatibilityVersion\":")) // No compatibility version found so assume that this is a pre-release version
                    {
                        // Persist the default settings values
                        try
                        {
                            // Rename the current settings file to preserve it
                            string badVersionSettingsFileName = $"{SettingsFileName}.prereleaseversion";
                            File.Delete(badVersionSettingsFileName);
                            File.Move(SettingsFileName, $"{badVersionSettingsFileName}");

                            // Persist the default settings values
                            settings = new();
                            PersistSettings(settings);

                            Status = $"A pre-release settings file was found. Settings have been reset to defaults and the original settings file has been renamed to {badVersionSettingsFileName}.";
                        }
                        catch (Exception ex2)
                        {
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"A pre-release settings file found but an error occurred when persisting new Conform settings: {ex2}");
                            Status = $"$\"A pre-release settings file was found but an error occurred when persisting new Conform settings: {ex2.Message}.";
                        }
                    }
                    else // File does have a compatibility version so read in the settings from the file
                    {
                        // Try to read in the settings version number from the settings file
                        try
                        {
                            var appSettings = JsonDocument.Parse(serialisedSettings, new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip });
                            settingsFileVersion = appSettings.RootElement.GetProperty("SettingsCompatibilityVersion").GetInt32();
                        }
                        catch (KeyNotFoundException)
                        {
                            // Ignore key not found exceptions because this indicates a corrupt file or a pre-release 1.0.0 version 
                        }

                        TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Found settings version: {settingsFileVersion}");

                        // Handle different file versions
                        switch (settingsFileVersion)
                        {
                            // File version 1 - first production release
                            case 1:

                                try
                                {
                                    // Set de-serialisation options
                                    JsonSerializerOptions options = new()
                                    {
                                        PropertyNameCaseInsensitive = true // Ignore incorrect element name casing
                                    };
                                    options.Converters.Add(new JsonStringEnumConverter()); // For increased resilience, accept both string member names and integer member values as valid for enum elements.

                                    // De-serialise the settings string into a Settings object
                                    settings = JsonSerializer.Deserialize<Settings>(serialisedSettings, options);

                                    // Test whether the retrieved settings match the requirements of this version of ConformU
                                    if (settings.SettingsCompatibilityVersion == Settings.SETTINGS_COMPATIBILTY_VERSION) // Version numbers match so all is well
                                    {
                                        Status = $"Settings read OK.";
                                    }
                                    else // Version numbers don't match so reset to defaults
                                    {
                                        int originalSettingsCompatibilityVersion = 0;
                                        try
                                        {
                                            originalSettingsCompatibilityVersion = settings.SettingsCompatibilityVersion;

                                            // Rename the current settings file to preserve it
                                            string badVersionSettingsFileName = $"{SettingsFileName}.badversion";
                                            File.Delete(badVersionSettingsFileName);
                                            File.Move(SettingsFileName, $"{badVersionSettingsFileName}");

                                            // Persist the default settings values
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
                                catch (JsonException ex1)
                                {
                                    // There was an exception when parsing the settings file so report it and set default values
                                    TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error parsing Conform settings file: {ex1}");

                                    // Set default values
                                    settings = new();

                                    Status = $"The settings file is corrupt ({ex1.Message}) and application settings have been reset to default values. Please correct the error in the file or use the \"Reset to defaults\" button on the Settings page to persist new values.";
                                }
                                catch (Exception ex1)
                                {
                                    TL?.LogMessage("ConformConfiguration", MessageLevel.Error, ex1.ToString());
                                    Status = $"Exception reading the settings file, default values are in use.";
                                }
                                break;

                            // Handle unknown settings version numbers
                            default:

                                // Persist default settings values because the file version is unknown and the file may be corrupt
                                try
                                {
                                    // Rename the current settings file to preserve it
                                    string badVersionSettingsFileName = $"{SettingsFileName}.unknownversion";
                                    File.Delete(badVersionSettingsFileName);
                                    File.Move(SettingsFileName, $"{badVersionSettingsFileName}");

                                    // Persist the default settings values
                                    settings = new();
                                    PersistSettings(settings);

                                    Status = $"An unsupported settings version was found: {settingsFileVersion}. Settings have been reset to defaults and the original settings file has been renamed to {badVersionSettingsFileName}.";
                                }
                                catch (Exception ex2)
                                {
                                    TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"An unsupported settings version was found: {settingsFileVersion} but an error occurred when persisting new Conform settings: {ex2}");
                                    Status = $"$\"An unsupported settings version was found: {settingsFileVersion} but an error occurred when persisting new Conform settings: {ex2.Message}.";
                                }
                                break;
                        }
                    }
                }
                else // Settings file does not exist
                {
                    TL.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Configuration file does not exist, initialising new file: {SettingsFileName}");
                    PersistSettings(settings);
                    Status = $"First time use - settings set to default values.";
                }
            }
            catch (Exception ex)
            {
                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, ex.ToString());
                Status = $"Unexpected exception reading the settings file, default values are in use.";
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

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put clean-up code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Public methods

        public void Reset()
        {
            TL?.LogMessage("Reset", MessageLevel.Debug, "Resetting settings file to default values");
            settings = new();
            PersistSettings(settings);
            Status = $"Settings reset at {DateTime.Now:HH:mm:ss.f}.";
            RaiseUiHasChangedEvent();
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

        public Settings Settings
        {
            get { return settings; }
        }

        public string SettingsFileName { get; private set; }

        /// <summary>
        /// Text message describing any issues found when validating the settings
        /// </summary>
        public string Status { get; private set; }

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

        #endregion

        #region Event handlers

        public delegate void MessageEventHandler(object sender, MessageEventArgs e);

        public event EventHandler ConfigurationChanged;

        public event EventHandler UiHasChanged;

        #endregion

        #region Support code

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

        #endregion

    }
}