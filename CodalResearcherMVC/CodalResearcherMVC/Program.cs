using CodalResearcherMVC.Entities;
using CodalResearcherMVC.Services;
using DocumentFormat.OpenXml.InkML;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddHttpClient(); 
builder.Services.AddSingleton<CodalScraperSelenium>(); 
builder.Services.AddScoped<HtmlExcelParser>();
builder.Services.AddScoped<ExcelTableService>();

// Add SignalR
builder.Services.AddSignalR(options =>
{
    options.EnableDetailedErrors = true;
    options.KeepAliveInterval = TimeSpan.FromSeconds(15);
    options.ClientTimeoutInterval = TimeSpan.FromSeconds(30);
});

IConfiguration configuration = builder.Configuration;
builder.Services.AddDbContext<ExcelTableDbContext>(options =>
    options.UseSqlServer(configuration.GetConnectionString("ExcelTableDb")));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Codal}/{action=Index}/{id?}");

app.MapHub<ProcessingHub>("/processingHub");

app.Run();