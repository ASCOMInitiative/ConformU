using System;

namespace ConformU
{
    public interface IConformConfiguration
    {
        Settings Settings { get; }
        string SettingsFileName { get; }
        string Status { get; }

        event EventHandler ConfigurationChanged;
        //event EventHandler UiHasChanged;

        void Dispose();
        void Reset();
        void Save();
        string Validate();
    }
}