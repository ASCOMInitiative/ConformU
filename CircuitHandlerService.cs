using System;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Diagnostics;

namespace ConformU
{
    public class CircuitHandlerService : CircuitHandler
    {
        private IHostApplicationLifetime lifetime = null;
        private ILogger<Startup> logger = null;

        public CircuitHandlerService(IHostApplicationLifetime lifetime, ILogger<Startup> logger)
        {
            this.lifetime = lifetime;
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
            logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} ***** OnCircuitClosedAsync {circuit.Id} *****");
            if (!Debugger.IsAttached) // Only use this mechanic outside of a dev environment
            {
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} About to call StopApplication()");

                lifetime.StopApplication();
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} Called StopApplication(), ending process.");
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} About to kill process...");
                Process.GetCurrentProcess().Kill();
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} Killed process.");
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} About to exit process...");
                Environment.Exit(0);
                logger.LogInformation($"{DateTime.Now:HH:mm:ss.fff} Exited process");
                //throw new Exception("END TASK!!");
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
