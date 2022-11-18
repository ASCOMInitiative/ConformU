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
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} ***** OnConnectionUpAsync {circuit.Id} *****");
            return base.OnConnectionUpAsync(circuit, cancellationToken);
        }

        public override Task OnConnectionDownAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} ***** OnConnectionDownAsync {circuit.Id} *****");
            Environment.Exit(0);
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} ***** OnConnectionDownAsync - AFter EXIT *****");
            return Task.CompletedTask; //base.OnConnectionDownAsync(circuit, cancellationToken);
        }

        public override Task OnCircuitOpenedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} ***** OnCircuitOpenedAsync {circuit.Id} *****");
            return base.OnCircuitOpenedAsync(circuit, cancellationToken);
        }

        public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff}  ***** OnCircuitClosedAsync {circuit.Id} *****");
            return base.OnCircuitClosedAsync(circuit, cancellationToken);
        }
    }

}
