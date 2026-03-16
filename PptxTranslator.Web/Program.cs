using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Http.Features;
using PptxTranslator.Web;
using PptxTranslator.Web.Components;
using PptxTranslator.Web.Features.Translation;

var builder = WebApplication.CreateBuilder(args);

// Add service defaults & Aspire client integrations.
builder.AddServiceDefaults();

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 1024 * 1024 * 1024; //1 GB
});

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Data protection
string dpDirName = "dpkeys-sosuweb";
Directory.CreateDirectory(dpDirName);
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(new DirectoryInfo(dpDirName))
    .SetApplicationName("SosuWeb");

builder.Services.AddOutputCache();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseAntiforgery();
app.UseOutputCache();
app.MapStaticAssets();
app.MapTranslateFeature();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.MapDefaultEndpoints();

app.Run();
