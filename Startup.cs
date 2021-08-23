using BlazorPro.BlazorSize;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Radzen;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; set; }

        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            // Blazor infrastructure
            services.AddRazorPages();
            //services.AddServerSideBlazor();
            services.AddServerSideBlazor(options => {
                options.DisconnectedCircuitRetentionPeriod = TimeSpan.FromSeconds(10);
            });
         

                        // Conform components
                        string loggerName;
            // Set log name with casing appropriate to OS
            if (OperatingSystem.IsWindows())
            {
                loggerName = "ConformU";
            }
            else
            {
                loggerName = "conformu";
            }

            string logFileName = Configuration.GetValue<string>(ConformConstants.COMMAND_OPTION_LOGFILENAME) ?? "";
            string logFilePath = Configuration.GetValue<string>(ConformConstants.COMMAND_OPTION_LOGFILEPATH) ?? "";

            // Use fully qualified file name if present, otherwise use log file path and relative file name
            if (Path.IsPathFullyQualified(logFileName)) // Full file name and path provided so split into path and filename and ignore any supplied log file path
            {
                logFilePath = Path.GetDirectoryName(logFileName);
                logFileName = Path.GetFileName(logFileName);
            }
            else // Relative file name so use supplied log file name and path
            {
                // No action required
            }

            ConformLogger conformLogger = new(logFileName, logFilePath, loggerName, true);  // Create a logger component
            conformLogger.Debug = true;
            services.AddSingleton(conformLogger); // Add the logger component to the list of injectable services

            // Create a ConformConfiguration service
            ConformConfiguration conformConfiguration = new(conformLogger, Configuration.GetValue<string>(ConformConstants.COMMAND_OPTION_SETTINGS)); // Create a configuration settings component

            // Enable Alpaca discovery if a command line option requires this
            string debugDiscovery = Configuration.GetValue<string>(ConformConstants.COMMAND_OPTION_SHOW_DISCOVERY) ?? "";
            if (!string.IsNullOrEmpty(debugDiscovery)) conformConfiguration.DebugDiscovery = true;

            // Add the configuration component to the list of injectable services
            services.AddSingleton(conformConfiguration);

            // Resizeable screen log text area infrastructure
            services.AddResizeListener(options =>
                {
                    options.ReportRate = 100; // Milliseconds between update notifications (I think - documentation not clear)
                    options.EnableLogging = false; // Better performance
                    options.SuppressInitEvent = false; // Ensure the event fires when the application is first loaded
                });

            // Radzen services
            services.AddScoped<NotificationService>();

            // Add event handler to detect when the browser closes
            services.AddSingleton<CircuitHandler, CircuitHandlerService>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env, ILogger<Startup> logger)
        {

            if (env.IsDevelopment())
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

            //foreach (var c in Configuration.AsEnumerable())
            //{
            //    Console.WriteLine($"{c.Key,-40}:{c.Value}");
            //}
        }
    }
}