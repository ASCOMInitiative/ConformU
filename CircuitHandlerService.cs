using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ConformU
{
    public class CircuitHandlerService : CircuitHandler
    {
        private readonly IHostApplicationLifetime lifetime = null; // This is required if the StopApplication method is used.
        readonly ILogger<CircuitHandlerService> logger;
        readonly object connectionsLockObject = new();

        private readonly List<string> connections;
        public CircuitHandlerService(IHostApplicationLifetime lifetime, ILogger<CircuitHandlerService> logger)
        {
            this.lifetime = lifetime; // This is required if the StopApplication method is used.
            this.logger = logger;

            // Create a new connections object if required
            if (connections is null)
            {
                lock (connectionsLockObject)
                {
                    connections = new();
                }
            }
        }

        public override Task OnConnectionUpAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation("***** OnConnectionUpAsync - Circuit {circuit.Id} is up. Connection count: {connections.Count} *****", circuit.Id, connections.Count);
            return Task.CompletedTask;
        }

        public override Task OnConnectionDownAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            logger.LogInformation("***** OnConnectionDownAsync - Circuit {circuit.Id} is down Connection count: {connections.Count} *****", circuit.Id, connections.Count);
            return Task.CompletedTask;
        }

        public override Task OnCircuitOpenedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            lock (connectionsLockObject)
            {
                try
                {
                    // Add the circuit to the list of circuits if not already in the list (it shouldn't be!)
                    if (!connections.Contains(circuit.Id))
                    {
                        logger.LogInformation("***** OnCircuitOpenedAsync - Adding connection {circuit.Id} Connection count: {connections.Count} *****", circuit.Id, connections.Count);
                        connections.Add(circuit.Id);
                        logger.LogInformation("***** OnCircuitOpenedAsync - Added connection {circuit.Id} Connection count: {connections.Count} *****", circuit.Id, connections.Count);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogInformation("***** OnCircuitOpenedAsync {circuit.Id} *****\r\n {ex}", circuit.Id, ex);
                }

                return Task.CompletedTask;
            }
        }

        public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
        {
            // Include a short delay to allow any new circuits to establish themselves before checking whether the application should close down
            Task.Delay(TimeSpan.FromSeconds(1),cancellationToken).Wait(cancellationToken);
            lock (connectionsLockObject)
            {
                try
                {
                    // Remove the circuit from the circuit list if present (it should be!)
                    if (connections.Contains(circuit.Id))
                    {
                        logger.LogInformation("***** OnCircuitClosedAsync - Removing connection {circuit.Id} Connection count: {connections.Count} *****", circuit.Id, connections.Count);
                        bool success = connections.Remove(circuit.Id);
                        logger.LogInformation("***** OnCircuitClosedAsync - Removed connection {circuit.Id} Connection count: {connections.Count}, Success: {success} *****", circuit.Id, connections.Count, success);
                    }

                    logger.LogInformation("***** OnCircuitClosedAsync - Before testing connection count: {connections.Count} *****", connections.Count);
                    // End the application if all circuits are closed
                    if (connections.Count == 0)
                    {
                        logger.LogInformation("***** OnCircuitClosedAsync - Calling StopApplication - {circuit.Id} Connection count: {connections.Count} *****", circuit.Id, connections.Count);
                        lifetime.StopApplication();
                    }
                }
                catch (Exception ex)
                {
                    logger.LogInformation("***** OnCircuitClosedAsync {circuit.Id} *****\r\n {ex}", circuit.Id, ex);
                }
                logger.LogInformation("***** OnCircuitClosedAsync - OnCircuitClosedAsync. Connection count: {connections.Count} *****", connections.Count);

                return Task.CompletedTask;
            }
        }
    }

}
