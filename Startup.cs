using BlazorPro.BlazorSize;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Radzen;
using System;
using System.Collections.Generic;
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

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            // Blazor infrastructure
            services.AddRazorPages();
            services.AddServerSideBlazor();

            // Conform components
            ConformLogger conformLogger = new("ConformU", true);  // Create a logger component
            services.AddSingleton(conformLogger); // Add the logger component to the list of injectable services

            ConformConfiguration conformConfiguration = new(conformLogger); // Create a configuration settings component
            services.AddSingleton(conformConfiguration); // Add the configuration component to the list of injectable services

            // Resizeable screen log text area infrastructure
            //services.AddMediaQueryService();
            //services.AddScoped<ResizeListener>();
            services.AddResizeListener(options =>
                {
                    options.ReportRate = 100; // Milliseconds between update notifications (I think - documentation not clear)
                    options.EnableLogging = false; // Better performance
                    options.SuppressInitEvent = false; // Ensure the event fires when the application is first loaded
                });

            // Radzen services
            services.AddScoped<NotificationService>();
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
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
                logger.LogInformation("***** Using production environment");
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