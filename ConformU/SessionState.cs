using System;
using System.Text;

namespace ConformU
{
    public class SessionState : IDisposable
    {
        private bool disposedValue;
        private readonly ConformLogger TL;

        // Max characters to retain in each log pane before trimming
        private const int LOG_MAX_LENGTH = 1_000_000;
        private const int LOG_TRIM_LENGTH = 900_000;

        private readonly StringBuilder _conformLogBuilder = new();
        private readonly StringBuilder _protocolLogBuilder = new();
        private readonly object _conformLogLock = new();
        private readonly object _protocolLogLock = new();

        public SessionState(ConformLogger logger)
        {
            TL = logger;
            TL?.LogMessage("ConformStateManager", MessageLevel.Debug, "Service started");
        }

        public event EventHandler UiHasChanged;

        public void RaiseUiHasChangedEvent()
        {
            if (UiHasChanged is not null)
            {
                EventArgs args = new();
                TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "About to call UI has changed event handler");
                UiHasChanged(this, args);
                TL?.LogMessage("RaiseUiHasChangedEvent", MessageLevel.Debug, "Returned from UI has changed event handler");
            }
        }

        /// <summary>
        /// The current content of the main conformance log pane.
        /// Setting this property replaces the entire buffer (use for clear/initialise).
        /// For incremental appends from the hot path use <see cref="AppendToConformLog"/>.
        /// </summary>
        public string ConformLog
        {
            get { lock (_conformLogLock) { return _conformLogBuilder.ToString(); } }
            set { lock (_conformLogLock) { _conformLogBuilder.Clear(); if (value is not null) _conformLogBuilder.Append(value); } }
        }

        /// <summary>
        /// Append a message to the conformance log, trimming the oldest content when the
        /// buffer exceeds <see cref="LOG_MAX_LENGTH"/> characters.
        /// </summary>
        public void AppendToConformLog(string message)
        {
            lock (_conformLogLock)
            {
                _conformLogBuilder.Append(message);
                if (_conformLogBuilder.Length > LOG_MAX_LENGTH)
                    _conformLogBuilder.Remove(0, _conformLogBuilder.Length - LOG_TRIM_LENGTH);
            }
        }

        /// <summary>
        /// Current value of the main ConformLog window scroll position
        /// </summary>
        /// <remarks>Used to restore the scroll position when the user does a browser refresh or returns to the home page.</remarks>
        public double ConformLogScrollTop { get; set; } = 0;

        public bool SafetyWarningDisplayed { get; set; } = false;

        /// <summary>
        /// The current content of the Alpaca protocol log pane.
        /// Setting this property replaces the entire buffer (use for clear/initialise).
        /// For incremental appends from the hot path use <see cref="AppendToProtocolLog"/>.
        /// </summary>
        public string ProtocolLog
        {
            get { lock (_protocolLogLock) { return _protocolLogBuilder.ToString(); } }
            set { lock (_protocolLogLock) { _protocolLogBuilder.Clear(); if (value is not null) _protocolLogBuilder.Append(value); } }
        }

        /// <summary>
        /// Append a message to the protocol log.
        /// </summary>
        public void AppendToProtocolLog(string message)
        {
            lock (_protocolLogLock)
            {
                _protocolLogBuilder.Append(message);
            }
        }

        #region IDisposable support

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ConformStateManager()
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

        #endregion
    }
}
