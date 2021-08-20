using System;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

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
            logger.LogInformation("***** OnConnectionUpAsync *****");
            return Task.CompletedTask;
        }

        public override Task OnConnectionDownAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation("***** OnConnectionDownAsync *****");

            return Task.CompletedTask;
        }

        public override Task OnCircuitOpenedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation($"***** OnCircuitOpenedAsync {circuit.Id} *****");
            return base.OnCircuitOpenedAsync(circuit, cancellationToken);
        }

        public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            //logger.LogWarning($"***** OnCircuitClosedAsync {circuit.Id} *****");
            logger.LogInformation($"About to call StopApplication()");

            lifetime.StopApplication();
            logger.LogInformation($"Called StopApplication(), ending process.");
            System.Diagnostics.Process.GetCurrentProcess().Kill();
            Environment.Exit(0);
            //throw new Exception("END TASK!!");
            return base.OnCircuitClosedAsync(circuit, cancellationToken);
        }
    }

}
