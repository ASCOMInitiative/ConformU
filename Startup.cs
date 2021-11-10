using Blazorise;
using Blazorise.Bootstrap;
using Blazorise.Icons.FontAwesome;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Radzen;
using System;
using System.IO;

namespace ConformU
{
    public class Startup
    {
        ConformConfiguration conformConfiguration;
        ConformLogger conformLogger;

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
                options.DisconnectedCircuitRetentionPeriod = TimeSpan.FromSeconds(300);
            });

            // Conform components
            #region Conform logger

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

            string logFileName = Configuration.GetValue<string>(Globals.COMMAND_OPTION_LOGFILENAME) ?? "";
            string logFilePath = Configuration.GetValue<string>(Globals.COMMAND_OPTION_LOGFILEPATH) ?? "";

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

            conformLogger = new(logFileName, logFilePath, loggerName, true);  // Create a logger component

            // Enable logging of application start-up if a command line option requires this
            string debugStartup = Configuration.GetValue<string>(Globals.COMMAND_OPTION_DEBUG_STARTUP) ?? "";
            if (!string.IsNullOrEmpty(debugStartup)) conformLogger.Debug = true;

            // Add the logger component to the list of injectable services
            services.AddSingleton(conformLogger);
            #endregion

            #region Conform configuration

            // Create a ConformConfiguration service
            conformConfiguration = new(conformLogger, Configuration.GetValue<string>(Globals.COMMAND_OPTION_SETTINGS)); // Create a configuration settings component

            // Enable logging of Alpaca discovery if a command line option requires this
            string debugDiscovery = Configuration.GetValue<string>(Globals.COMMAND_OPTION_DEBUG_DISCOVERY) ?? "";
            if (!string.IsNullOrEmpty(debugDiscovery)) conformConfiguration.Settings.TraceDiscovery = true;

            // Set the results filename if supplied on the command line
            string resultsFileName = Configuration.GetValue<string>(Globals.COMMAND_OPTION_RESULTS_FILENAME) ?? "";
            if (!string.IsNullOrEmpty(resultsFileName)) conformConfiguration.Settings.ResultsFileName = resultsFileName;

            // Add the configuration component to the list of injectable services
            services.AddSingleton(conformConfiguration);

            #endregion

            // Add BlazorPro screen resize listener 
            //services.AddResizeListener(options =>
            //    {
            //        options.ReportRate = 10; // Milliseconds between update notifications (I think - documentation not clear)
            //        options.EnableLogging = false; // Better performance
            //        options.SuppressInitEvent = false; // Ensure the event fires when the application is first loaded
            //    });

            // Add window resize listener
            services.AddSingleton<BrowserResizeService>();

            // Radzen services
            services.AddScoped<NotificationService>();

            // Add event handler to detect when the browser closes
            services.AddSingleton<CircuitHandler, CircuitHandlerService>();

            // Add Blazorise services
            services.AddBlazorise(options =>
             {
                 options.ChangeTextOnKeyPress = true; // optional
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
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env, ILogger<Startup> logger)
        {
            //applicationLifetime.ApplicationStopping.Register(DisposeObject, conformConfiguration);
            //applicationLifetime.ApplicationStopping.Register(DisposeObject, conformLogger);

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

        }
    }
}