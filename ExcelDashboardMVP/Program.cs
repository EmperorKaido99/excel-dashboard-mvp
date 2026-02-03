using ExcelDashboardMVP.Components;
using ExcelDashboardMVP.Services;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// EPPlus 8.x license
ExcelPackage.License.SetNonCommercialPersonal("ExcelDashboardMVP");

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Scoped — created fresh per request (original import service)
builder.Services.AddScoped<ExcelImportService>();

// Singleton — one instance for the whole app lifetime so imported
// data stays in memory and is reachable from any page.
builder.Services.AddSingleton<ExcelDataService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();