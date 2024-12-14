using ASCOM.Common;
using ASCOM.Common.Alpaca;
using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ConformU
{
    [DefaultMember("Settings")]
    public class ConformConfiguration : IDisposable
    {
        private const string FOLDER_NAME = "conform"; // Folder name underneath the local application data folder
        private const string SETTINGS_FILENAME = "conform.settings"; // Settings file name

        private ConformLogger TL;
        private Settings settings;
        private bool disposedValue;
        private readonly int settingsFileVersion;
        private readonly JsonDocument appSettingsDocument = null;

        private readonly SessionState conformStateManager;

        private static readonly JsonSerializerOptions jsonSerialisationOptions; // JSON De-serialisation options

        #region Initialisers and Dispose

        /// <summary>
        /// Static initialiser so we only need to set JSON de-serialisation options once
        /// </summary>
        static ConformConfiguration()
        {
            // Set JSON de-serialisation options
            jsonSerialisationOptions = new()
            {
                PropertyNameCaseInsensitive = true, // Ignore incorrect element name casing
                WriteIndented = true
            };
            jsonSerialisationOptions.Converters.Add(new JsonStringEnumConverter()); // For increased resilience, accept both string member names and integer member values as valid for enum elements.
        }

        /// <summary>
        /// Create a Configuration management instance and load the current settings
        /// </summary>
        /// <param name="logger">Data logger instance.</param>
        public ConformConfiguration(ConformLogger logger, SessionState conformStateManager, string configurationFile)
        {
            TL = logger;
            this.conformStateManager = conformStateManager;

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

                            Status = $"A pre-release settings file was found.\r\n\r\nApplication settings have been reset to defaults and the original settings file has been renamed to {badVersionSettingsFileName}.";
                        }
                        catch (Exception ex2)
                        {
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"A pre-release settings file found but an error occurred when saving new Conform settings: {ex2}");
                            Status = $"$\"A pre-release settings file was found but an error occurred when saving new Conform settings: {ex2.Message}.";
                        }
                    }
                    else // File does have a compatibility version so read in the settings from the file
                    {
                        TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Found compatibility version element...");
                        // Try to read in the settings version number from the settings file
                        try
                        {
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"About to parse settings string");
                            appSettingsDocument = JsonDocument.Parse(serialisedSettings, new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip });
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"About to get settings version");
                            settingsFileVersion = appSettingsDocument.RootElement.GetProperty("SettingsCompatibilityVersion").GetInt32();
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Found settings version: {settingsFileVersion}");

                            // Handle different file versions
                            switch (settingsFileVersion)
                            {
                                // File version 1 - first production release
                                case 1:

                                    try
                                    {
                                        // De-serialise the settings string into a Settings object
                                        settings = JsonSerializer.Deserialize<Settings>(serialisedSettings, jsonSerialisationOptions);

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

                                                Status = $"The current settings version: {originalSettingsCompatibilityVersion} does not match the required version: {Settings.SETTINGS_COMPATIBILTY_VERSION}. Application settings have been reset to default values and the original settings file renamed to {badVersionSettingsFileName}.";
                                            }
                                            catch (Exception ex2)
                                            {
                                                TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error persisting new Conform settings file: {ex2}");
                                                Status = $"The current settings version:{originalSettingsCompatibilityVersion} does not match the required version: {Settings.SETTINGS_COMPATIBILTY_VERSION} but the new settings could not be saved: {ex2.Message}.";
                                            }
                                        }
                                    }
                                    catch (JsonException ex1)
                                    {
                                        // There was an exception when parsing the settings file so report it and set default values
                                        TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error de-serialising Conform settings file: {ex1}");
                                        Status = $"There was an error de-serialising the settings file and application default settings are in effect.\r\n\r\nPlease correct the error in the file or use the \"Reset to Defaults\" button on the Settings page to save new values.\r\n\r\nJSON parser error message:\r\n{ex1.Message}";
                                    }
                                    catch (Exception ex1)
                                    {
                                        TL?.LogMessage("ConformConfiguration", MessageLevel.Error, ex1.ToString());
                                        Status = $"Exception reading the settings file, default values are in effect.";
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
                                        TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"An unsupported settings version was found: {settingsFileVersion} but an error occurred when saving new Conform settings: {ex2}");
                                        Status = $"$\"An unsupported settings version was found: {settingsFileVersion} but an error occurred when saving new Conform settings: {ex2.Message}.";
                                    }
                                    break;
                            }
                        }
                        catch (JsonException ex)
                        {
                            // There was an exception when parsing the settings file so report it and use default values
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Error getting settings file version from settings file: {ex}");
                            Status = $"An error occurred when reading the settings file version and application default settings are in effect.\r\n\r\nPlease correct the error in the file or use the \"Reset to Defaults\" button on the Settings page to create a new settings file.\r\n\r\nJSON parser error message:\r\n{ex.Message}";
                        }
                        catch (Exception ex)
                        {
                            TL?.LogMessage("ConformConfiguration", MessageLevel.Error, $"Exception parsing the settings file: {ex}");
                            Status = $"Exception parsing the settings file: {ex.Message}";
                        }
                        finally
                        {
                            appSettingsDocument?.Dispose();
                        }
                    }
                }
                else // Settings file does not exist
                {
                    TL.LogMessage("ConformConfiguration", MessageLevel.Debug, $"Configuration file does not exist, initialising new file: {SettingsFileName}");
                    PersistSettings(settings);
                    Status = $"First time use - configuration set to default values.";
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
            try
            {
                settings = new();
                PersistSettings(settings);
                Status = $"Settings reset at {DateTime.Now:HH:mm:ss}.";
                conformStateManager.RaiseUiHasChangedEvent();
            }
            catch (Exception ex)
            {
                TL?.LogMessage("Reset", MessageLevel.Error, $"Exception during Reset: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Persist current settings
        /// </summary>
        public void Save()
        {
            TL?.LogMessage("Save", MessageLevel.Debug, "Saving settings to settings file");
            PersistSettings(settings);
            Status = $"Settings saved at {DateTime.Now:HH:mm:ss}.";

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
                    if (settings.AlpacaDevice.InterfaceVersion != 1) issueList += $"\r\nAlpaca interface version must be 1 but is actually: {settings.AlpacaDevice.InterfaceVersion}";
                    break;

                case DeviceTechnology.COM:
                    if (settings.ComDevice.ProgId == "") return "CurrentComDevice.ProgId  is empty.";

                    break;
            }
            issueList = issueList.Trim();

            return issueList;
        }

        /// <summary>
        /// Sets data values to test a COM device
        /// </summary>
        /// <param name="progId"></param>
        /// <param name="deviceType"></param>
        public void SetComDevice(string progId, DeviceTypes deviceType)
        {
            settings.DeviceTechnology = DeviceTechnology.COM;

            settings.ComDevice.ProgId = progId;
            settings.DeviceType = deviceType;
        }

        /// <summary>
        /// Sets data values to test an Alpaca device
        /// </summary>
        /// <param name="serviceType"></param>
        /// <param name="address"></param>
        /// <param name="port"></param>
        /// <param name="deviceType"></param>
        /// <param name="deviceNumber"></param>
        public void SetAlpacaDevice(ServiceType serviceType, string address, int port, int alpacaInterfaceVersion, DeviceTypes deviceType, int deviceNumber)
        {
            settings.DeviceTechnology = DeviceTechnology.Alpaca;

            settings.AlpacaDevice.ServiceType = serviceType;
            settings.AlpacaDevice.ServerName = address;
            settings.AlpacaDevice.IpAddress = address;
            settings.AlpacaDevice.IpPort = port;
            settings.AlpacaDevice.InterfaceVersion = alpacaInterfaceVersion;
            settings.AlpacaDevice.AscomDeviceType = deviceType;
            settings.AlpacaDevice.AlpacaDeviceNumber = deviceNumber;
            settings.DeviceType = deviceType;
        }

        /// <summary>
        /// Sets configuration to run a complete test without debug or trace
        /// </summary>
        /// <remarks>Uses reflection to iterate over settings properties looking for IncludeInFullTest attributes. The default values supplied in the attributes are then applied to the properties.</remarks>
        public void SetFullTest()
        {
            // Get all Settings class properties
            var settingsProperties = typeof(Settings).GetProperties();

            // Iterate over the Settings class properties
            foreach (var property in settingsProperties)
            {
                // Extract the IncludeInFullTest attribute if present
                MandatoryInFullTestAttribute attribute = property.GetCustomAttribute<MandatoryInFullTestAttribute>();

                // If present set the property to its "full settings" value contained in the attribute
                if (attribute != null)
                {
                    TL.LogMessage("SetFullTest", MessageLevel.Debug, $"Found property Settings.{property.Name}, Default value: {attribute.FullSettingsValue}");

                    // Get a PropertyInfor for the property
                    PropertyInfo propertyInfo = Settings.GetType().GetProperty(property.Name);

                    // Set the property's value to the configured full settings value 
                    propertyInfo.SetValue(Settings, attribute.FullSettingsValue, null);
                }
            }

            // Get all AlpacaConfiguration class properties
            var alpacaConfigurationProperties = typeof(AlpacaConfiguration).GetProperties();

            // Iterate over the AlpacaConfiguration class properties
            foreach (var property in alpacaConfigurationProperties)
            {
                // Extract the IncludeInFullTest attribute if present
                MandatoryInFullTestAttribute attribute = property.GetCustomAttribute<MandatoryInFullTestAttribute>();

                // If present set the property to its "full settings" value contained in the attribute
                if (attribute != null)
                {
                    TL.LogMessage("SetFullTest", MessageLevel.Debug, $"Found property Settings.AlpacaConfiguration.{property.Name}, Default value: {attribute.FullSettingsValue}");

                    // Get a PropertyInfor for the property
                    PropertyInfo propertyInfo = Settings.AlpacaConfiguration.GetType().GetProperty(property.Name);

                    // Set the property's value to the configured full settings value 
                    propertyInfo.SetValue(Settings.AlpacaConfiguration, attribute.FullSettingsValue, null);
                }
            }

            // Enable all Telescope method tests
            settings.TelescopeTests = new()
            {
                { "CanMoveAxis", true },
                { "Park/Unpark", true },
                { "AbortSlew", true },
                { "AxisRate", true },
                { "FindHome", true },
                { "MoveAxis", true },
                { "PulseGuide", true },
                { "SlewToCoordinates", true },
                { "SlewToCoordinatesAsync", true },
                { "SlewToTarget", true },
                { "SlewToTargetAsync", true },
                { "DestinationSideOfPier", true },
                { "SlewToAltAz", true },
                { "SlewToAltAzAsync", true },
                { "SyncToCoordinates", true },
                { "SyncToTarget", true },
                { "SyncToAltAz", true }
            };
        }

        #endregion

        #region Event handlers

        public delegate void MessageEventHandler(object sender, MessageEventArgs e);

        public event EventHandler ConfigurationChanged;

        //public event EventHandler UiHasChanged;

        #endregion

        #region Support code

        private void PersistSettings(Settings settingsToPersist)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(settingsToPersist, nameof(settingsToPersist));

                // Set the version number of this settings file
                settingsToPersist.SettingsCompatibilityVersion = Settings.SETTINGS_COMPATIBILTY_VERSION;

                TL?.LogMessage("PersistSettings", MessageLevel.Debug, $"Settings file: {SettingsFileName}");

                string serialisedSettings = JsonSerializer.Serialize<Settings>(settingsToPersist, jsonSerialisationOptions);

                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFileName));
                File.WriteAllText(SettingsFileName, serialisedSettings);
            }
            catch (Exception ex)
            {
                TL.LogMessage("PersistSettings", MessageLevel.Debug, ex.ToString());
            }

        }

        //internal void RaiseUiHasChangedEvent()
        //{
        //    if (UiHasChanged is not null)
        //    {
        //        EventArgs args = new();
        //        TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "About to call UI has changed event handler");
        //        UiHasChanged(this, args);
        //        TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "Returned from UI has changed event handler");
        //    }
        //}

        #endregion

    }
}