using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using QuattroFacturatieProgramma.ViewModels;
using QuattroFacturatieProgramma.Helpers;
using OfficeOpenXml;

namespace QuattroFacturatieProgramma;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var builder = MauiApp.CreateBuilder();
        builder
            .UseMauiApp<App>()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

        // Configuration setup - probeer eerst appsettings.json, fallback naar in-memory
        var configuration = CreateConfiguration();
        builder.Services.AddSingleton<IConfiguration>(configuration);

        // ViewModels registreren
        builder.Services.AddSingleton<MainViewModel>();

        // Pages registreren
        builder.Services.AddSingleton<MainPage>();

        // Helpers en Services registreren
        builder.Services.AddTransient<MollieApiHelper>();
        builder.Services.AddTransient<FactuurService>();

        // KlantHelper toevoegen - Singleton omdat er maar één Excel bestand is
        builder.Services.AddSingleton<KlantHelper>(provider =>
        {
            var excelPad = Path.Combine(FileSystem.AppDataDirectory, "uren 2025.xlsx");
            return new KlantHelper(excelPad);
        });

        // LogoHelper is static, dus geen registratie nodig

#if DEBUG
        builder.Services.AddLogging(configure => configure.AddDebug());
#endif

        return builder.Build();
    }

    private static IConfiguration CreateConfiguration()
    {
        var configBuilder = new ConfigurationBuilder();

        try
        {
            // Probeer eerst appsettings.json te laden
            configBuilder.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            var config = configBuilder.Build();

            // Check of Mollie API key bestaat in appsettings.json
            var mollieKey = config["Mollie:ApiKey"];
            if (!string.IsNullOrEmpty(mollieKey))
            {
                Console.WriteLine("✅ Configuratie geladen uit appsettings.json");
                return config;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ Kon appsettings.json niet laden: {ex.Message}");
        }

        // Fallback naar in-memory configuratie
        Console.WriteLine("⚠️ Gebruik fallback configuratie - voeg appsettings.json toe voor productie");
        var inMemoryConfig = new Dictionary<string, string>
        {
            ["Mollie:ApiKey"] = "live_GPnfJy9SF2sakpxdbSHmnMJqj6GGSN",
            ["Factuurinstellingen:BedrijfsLogo"] = "quattro_logo.png",
            ["Factuurinstellingen:BedrijfsNaam"] = "QUATTRO BOUW & VASTGOED ADVIES BV",
            ["Factuurinstellingen:Adres"] = "Willinkhof 3",
            ["Factuurinstellingen:Postcode"] = "6006 RG",
            ["Factuurinstellingen:Plaats"] = "Weert",
            ["Factuurinstellingen:Website"] = "www.quattrobouwenenvastgoedadvies.nl",
            ["Factuurinstellingen:Email"] = "info@quattrobouwenenvastgoedadvies.nl",
            ["Factuurinstellingen:KvK"] = "75108542",
            ["Factuurinstellingen:BTW"] = "NL860145438B01",
            ["Factuurinstellingen:IBAN"] = "NL30 RABO 0347 6704 07",
            ["Factuurinstellingen:BetalingsTermijn"] = "14",
            ["Factuurinstellingen:BTWPercentage"] = "21.0",
            ["QrCode:EpcFallback"] = "true",
            ["QrCode:PixelsPerModule"] = "20",
            ["QrCode:MaxSize"] = "100",
            // Excel configuratie toevoegen
            ["Excel:UrenBestand"] = "uren 2025.xlsx",
            ["Excel:Werkblad"] = "Realisatie 2025"
        };

        configBuilder = new ConfigurationBuilder();
        configBuilder.AddInMemoryCollection(inMemoryConfig);
        return configBuilder.Build();
    }
}