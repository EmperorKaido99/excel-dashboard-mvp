using ExcelDashboardMVP.Components;
using ExcelDashboardMVP.Services;
using OfficeOpenXml;
using MudBlazor.Services;

var builder = WebApplication.CreateBuilder(args);

// EPPlus 8.x non-commercial licence
ExcelPackage.License.SetNonCommercialPersonal("ExcelDashboardMVP");

// Razor + interactive server components
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// MudBlazor (theme, dialog, snackbar, resize-listener, etc.)
builder.Services.AddMudServices();

// ── ExcelDataService ──────────────────────────────────────────────────────
builder.Services.AddSingleton<ExcelDataService>();

// ── ExcelImportService ────────────────────────────────────────────────────
// Stateless — Scoped is fine.
builder.Services.AddScoped<ExcelImportService>();

var app = builder.Build();

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