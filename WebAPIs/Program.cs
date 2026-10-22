using Libraries;
using Libraries.LibModels;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);
IConfiguration configuration = builder.Configuration;
IHostEnvironment environment = builder.Environment;
IServiceCollection services = builder.Services;

// ==========================================
// 1. CONFIGURE SERVICES
// ==========================================

// Add controllers
services.AddControllers();

// Required for Swagger endpoint discovery in .NET 6+
services.AddEndpointsApiExplorer();

// Add HTTP context access
services.AddHttpContextAccessor();

// Add CORS
services.AddCors(options => options.AddPolicy(configuration.GetValue<string>("AllowedOriginPolicy"), builder =>
    builder.WithOrigins(configuration.GetValue<string>("AllowedHostsCORS"))
           .AllowAnyMethod()
           .AllowAnyHeader()));

// Add MVC controllers as services
services.AddMvc().AddControllersAsServices();

// Add session management
services.AddSession();

// Register the Swagger generator
services.AddSwaggerGen();

// Authentication & Authorization
services.AddAuthentication();
services.AddAuthorization();

// Bind configuration thresholds
services.Configure<MigrationThresholds>(configuration.GetSection("MigrationThresholds"));

// Custom Modular Registrations from your Libraries
services.ResolveIntegration();
services.ResolveServices();
services.ResolveRepositories();

// ==========================================
// 2. BUILD APPLICATION
// ==========================================
await using WebApplication app = builder.Build();

// ==========================================
// 3. CONFIGURE HTTP REQUEST PIPELINE
// ==========================================
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();

    // Enable Swagger JSON endpoint and Interactive UI
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "WCMS Migration API v1");
    });
}
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();

// Serve the landing page (wwwroot/index.html) and static assets.
// UseDefaultFiles must run before UseStaticFiles so "/" resolves to index.html.
app.UseDefaultFiles();
app.UseStaticFiles();

// Routing must come before CORS and Endpoints
app.UseRouting();

// CORS policy evaluation (placed between UseRouting and UseEndpoints)
app.UseCors(configuration.GetValue<string>("AllowedOriginPolicy"));

app.UseSession();
app.UseAuthentication();
app.UseAuthorization();

app.UseEndpoints(endpoints => endpoints.MapControllers());

// Run the application
await app.RunAsync();