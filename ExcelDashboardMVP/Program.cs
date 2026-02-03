using ExcelDashboardMVP.Components;
using ExcelDashboardMVP.Services;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// EPPlus 8.x license
ExcelPackage.License = new LicenseInfo { LicenseType = LicenseType.NonCommercial };

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Register Excel Import Service
builder.Services.AddScoped<ExcelImportService>();

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