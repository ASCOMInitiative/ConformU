using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace ConformU
{
    public class CircuitHandlerService : CircuitHandler
    {
        private readonly IHostApplicationLifetime lifetime = null; // This is required if the StopApplication method is used.
        private readonly ILogger<Startup> logger = null;

        //public CircuitHandlerService(IHostApplicationLifetime lifetime, ILogger<Startup> logger)
        public CircuitHandlerService(IHostApplicationLifetime lifetime, ILogger<Startup> logger)
        {
            this.lifetime = lifetime; // This is required if the StopApplication method is used.
            this.logger = logger;
        }

        public override Task OnConnectionUpAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** OnConnectionUpAsync *****");
            return base.OnCircuitClosedAsync(circuit, cancellationToken);
        }

        public override Task OnConnectionDownAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** OnConnectionDownAsync *****");

            return base.OnCircuitClosedAsync(circuit, cancellationToken);
        }

        public override Task OnCircuitOpenedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** OnCircuitOpenedAsync {circuit.Id} *****");
            return base.OnCircuitOpenedAsync(circuit, cancellationToken);
        }

        public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            Console.WriteLine("OnCircuitClosedAsync has been called...");
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** OnCircuitClosedAsync {circuit.Id} *****");
            if (!Debugger.IsAttached) // Only use this mechanic outside of a dev environment
            {
                // At the time of writing any attempt to use Environment.Exit or lifetime.StopApplication() results in an undefined wait time before the 
                // console hosting application closes. For this reason the more brutal Kill option is used to terminate the process immediately.
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} About to kill process...");
                Process.GetCurrentProcess().Kill();
                try
                {
                    lifetime.StopApplication();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"OnCircuitClosedAsync exception \r\n{ex}");
                }
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} Killed process.");
            }
            else
            {
                logger.LogInformation($"***** OnCircuitClosedAsync {circuit.Id} *****");
            }
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** END OF OnCircuitClosedAsync {circuit.Id} *****");
            return base.OnCircuitClosedAsync(circuit, cancellationToken);
        }
    }

}
