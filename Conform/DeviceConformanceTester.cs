using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using System.Threading.Tasks;
using System.Threading;

namespace ConformU
{
    public class DeviceConformanceTester : IDisposable
    {
        ConformConfiguration configuration;
        CancellationToken cancellationToken;
        public DeviceConformanceTester(ConformConfiguration conformConfiguration, CancellationToken conforCancellationToken)
        {
            configuration = conformConfiguration;
            cancellationToken = conforCancellationToken;
        }

        public void TestDevice()
        {
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine($"OutputChanged is null {OutputChanged is null} Loop {i}");
                OnLogMessageChanged("TestDevice", $"Loop {i} {cancellationToken.IsCancellationRequested} {configuration.Settings.CurrentDeviceType}");
                if (cancellationToken.IsCancellationRequested) break;
                Thread.Sleep(1000);
            }
            Console.WriteLine($"Finished processing: Task Cancelled: {cancellationToken.IsCancellationRequested}");
            OnLogMessageChanged("TestDevice", $"Finished processing: Task Cancelled: {cancellationToken.IsCancellationRequested}");

        }

        public event EventHandler<MessageEventArgs> OutputChanged;

        protected virtual void OnLogMessageChanged(string id, string message)
        {
            MessageEventArgs e = new()
            {
                Id = id,
                Message = message
            };

            EventHandler<MessageEventArgs> messageEventHandler = OutputChanged;

            if (messageEventHandler is not null)
            {
                messageEventHandler(this, e);
            }
        }


        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}