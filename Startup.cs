using Blazorise;
using Blazorise.Bootstrap;
using Blazorise.Icons.FontAwesome;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Radzen;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;

namespace ConformU
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; set; }

        /// <summary>
        /// This method gets called by the runtime before calling the Configure method. Use this method to add services to the container.
        /// For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        /// </summary>
        /// <param name="services"></param>
        public void ConfigureServices(IServiceCollection services)
        {
            // Blazor infrastructure
            services.AddRazorPages();
            services.AddServerSideBlazor(options =>
            {
                options.DisconnectedCircuitRetentionPeriod = TimeSpan.FromSeconds(Double.Parse(Configuration.GetValue<string>(Globals.COMMAND_OPTION_CONNECTION_TIMEOUT)));
                options.JSInteropDefaultCallTimeout = TimeSpan.FromSeconds(2);
                options.DetailedErrors = true;
            });

            // Add window resize listener
            services.AddSingleton<BrowserResizeService>();

            // Radzen services
            services.AddScoped<NotificationService>();

            // Add event handler to detect when the browser closes
            services.AddSingleton<CircuitHandler, CircuitHandlerService>();

            // Add Blazorise services
            services.AddBlazorise(options =>
             {
                 options.Immediate = true; // optional
             })
                .AddBootstrapProviders()
                .AddFontAwesomeIcons();
        }

        /// <summary>
        /// This method gets called by the runtime after services have been configured. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app"></param>
        /// <param name="env"></param>
        /// <param name="logger"></param>
        public void Configure(IApplicationBuilder app, IHostApplicationLifetime lifetime, IWebHostEnvironment env, ILogger<Startup> logger)
        {
            //applicationLifetime.ApplicationStopping.Register(DisposeObject, conformConfiguration);
            //applicationLifetime.ApplicationStopping.Register(DisposeObject, conformLogger);
            logger.LogInformation("***** Environment.CurrentDirectory: {Environment.CurrentDirectory}", Environment.CurrentDirectory);
            logger.LogInformation("***** Environment.ProcessPath: {Environment.ProcessPath}", Environment.ProcessPath);
            logger.LogInformation("***** Environment.ProcessorCount: {Environment.ProcessorCount}", Environment.ProcessorCount);

            // Start a browser if this a production application i.e. we are not running in Visual Studio
            if (!Debugger.IsAttached)
            {
                // Register a callback that will run after the application is fully configured to open the browser  (This can't be done earlier because the operating port isn't determined until Kestrel is fully running)
                lifetime.ApplicationStarted.Register(() =>
                {
                    StartBrowser(app.ServerFeatures, logger);
                });
            }

            // other config
            if (env.IsDevelopment()) // Running in 
            {
                app.UseDeveloperExceptionPage();
                logger.LogInformation("***** Using development environment");
            }
            else
            {
                //app.UseExceptionHandler("/Error");
                app.UseDeveloperExceptionPage();
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                //app.UseHsts();
                logger.LogInformation("***** Using modified production environment");
            }

            //
            // MAKE HTTP ONLY NO HTTPS SUPPORT
            //
            //app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapBlazorHub();
                endpoints.MapFallbackToPage("/_Host");
            });

        }

        /// <summary>
        /// Start a browser for the user to communicate with the server application
        /// </summary>
        /// <param name="features">Kestrel features collection</param>
        /// <param name="logger">ILogger instance</param>
        public static void StartBrowser(IFeatureCollection features, ILogger<Startup> logger)
        {
            IServerAddressesFeature addressFeature = features.Get<IServerAddressesFeature>();

            switch (addressFeature.Addresses.Count)
            {
                case 0: // FAILURE: No IP addresses are being used
                    logger.LogCritical("Cannot start browser because Kestrel is reporting that no IP addresses are in use.");
                    break;

                case 1: // SUCCESS: One IP address is being used 
                    OpenBrowser(logger, addressFeature);
                    break;

                default: // UNEXPECTED: More than one IP address is in use so we will just pick the first one and hope
                    logger.LogWarning("Kestrel is listening on more than one IP address so we will use the first one.");

                    // Iterate over the returned addresses
                    foreach (string addressItem in addressFeature.Addresses)
                    {
                        logger.LogInformation("Kestrel is listening on address: {address}", addressItem);
                    }

                    OpenBrowser(logger, addressFeature);

                    break;
            }
        }

        /// <summary>
        /// Open a browser using the correct syntax for each operating system
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="addressFeature"></param>
        private static void OpenBrowser(ILogger<Startup> logger, IServerAddressesFeature addressFeature)
        {
            string address;
            bool success;

            // Get the first (or only) address
            address = addressFeature.Addresses.First();
            logger.LogInformation("Listening on address: {address}, Substring: {address.Substring(8)}", address, address[7..]);

            // Parse the address (ignoring the 7 character http:// prefix) as an endpoint in order that we can use the IP port later
            success = IPEndPoint.TryParse(address[7..], out IPEndPoint iPEndPoint);
            logger.LogInformation("Success: {success}", success);

            // If the IP end point parsed OK open the browser
            if (success) // Address parsed OK
            {
                logger.LogInformation("Starting browser on localhost port: {iPEndPoint.Port}", iPEndPoint.Port);
                string url = $"http://localhost:{iPEndPoint.Port}/";

                // Start the browser using the appropriate command for each OS
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    logger.LogCritical(
                        "Unknown OS Platform, cannot start browser. Framework: {RuntimeInformation.FrameworkDescription}, " +
                        "Process architecture: {RuntimeInformation.ProcessArchitecture}, " +
                        "OS Architecture: {RuntimeInformation.OSArchitecture}, " +
                        "Run time identifier: {RuntimeInformation.RuntimeIdentifier}",
                        RuntimeInformation.FrameworkDescription, RuntimeInformation.ProcessArchitecture, RuntimeInformation.OSArchitecture, RuntimeInformation.RuntimeIdentifier);
                }
            }
            else // Address did not parse OK
            {
                logger.LogCritical("Cannot start browser, failed to parse address: {address}", address);
            }
        }
    }
}