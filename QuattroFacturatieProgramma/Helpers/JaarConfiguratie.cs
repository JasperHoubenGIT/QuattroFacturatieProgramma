using System;
using System.Collections.Generic;
using System.Globalization;

namespace QuattroFacturatieProgramma.Helpers
{
    /// <summary>
    /// Centrale configuratie voor jaar en maand logica met offset
    /// </summary>
    public static class JaarConfiguratie
    {
        /// <summary>
        /// Bepaalt welk boekjaar gebruikt moet worden op basis van huidige maand
        /// Januari = vorig jaar (December facturen), Februari+ = huidig jaar
        /// </summary>
        public static int BepaalBoekjaar()
        {
            var nu = DateTime.Now;

            // In januari doen we december van vorig jaar
            if (nu.Month == 1)
            {
                Console.WriteLine($"📅 Januari gedetecteerd - gebruik vorig jaar: {nu.Year - 1}");
                return nu.Year - 1;  // Januari 2026 → 2025
            }
            else
            {
                Console.WriteLine($"📅 Maand {nu.Month} gedetecteerd - gebruik huidig jaar: {nu.Year}");
                return nu.Year;      // Februari+ 2026 → 2026
            }
        }

        /// <summary>
        /// Bepaalt welk jaar gebruikt moet worden voor een specifieke maand
        /// </summary>
        public static int BepaalJaarVoorMaand(string maandNaam)
        {
            var nu = DateTime.Now;
            var maandNummer = ConverteerMaandNaamNaarNummer(maandNaam);

            // Als we in januari zijn en december selecteren, is dat vorig jaar
            if (nu.Month == 1 && maandNummer == 12)
            {
                Console.WriteLine($"📅 December factuur in januari - gebruik vorig jaar: {nu.Year - 1}");
                return nu.Year - 1;
            }

            // Anders gebruik het bepaalde boekjaar
            var boekjaar = BepaalBoekjaar();
            Console.WriteLine($"📅 {maandNaam} factuur - gebruik jaar: {boekjaar}");
            return boekjaar;
        }

        /// <summary>
        /// Bepaalt de eerste dag van de maand voor factuur datum
        /// </summary>
        public static DateTime BepaalEersteVanMaand(string maandNaam)
        {
            var jaar = BepaalJaarVoorMaand(maandNaam);
            var maandNummer = ConverteerMaandNaamNaarNummer(maandNaam);

            var datum = new DateTime(jaar, maandNummer, 1);
            Console.WriteLine($"📅 Eerste van maand voor {maandNaam}: {datum:dd-MM-yyyy}");
            return datum;
        }

        /// <summary>
        /// Excel bestand naam op basis van boekjaar
        /// </summary>
        public static string ExcelBestandNaam => $"Uren {BepaalBoekjaar()}.xlsx";

        /// <summary>
        /// Realisatie sheet naam op basis van boekjaar
        /// </summary>
        public static string RealisatieSheetNaam => $"Realisatie {BepaalBoekjaar()}";

        /// <summary>
        /// Converteert maandnaam naar nummer
        /// </summary>
        public static int ConverteerMaandNaamNaarNummer(string maandNaam)
        {
            var maandMapping = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["januari"] = 1,
                ["februari"] = 2,
                ["maart"] = 3,
                ["april"] = 4,
                ["mei"] = 5,
                ["juni"] = 6,
                ["juli"] = 7,
                ["augustus"] = 8,
                ["september"] = 9,
                ["oktober"] = 10,
                ["november"] = 11,
                ["december"] = 12
            };

            if (maandMapping.TryGetValue(maandNaam, out int maandNummer))
            {
                return maandNummer;
            }

            Console.WriteLine($"⚠️ Onbekende maand '{maandNaam}', gebruik huidige maand");
            return DateTime.Now.Month;
        }

        /// <summary>
        /// Converteert maandnummer naar naam
        /// </summary>
        public static string ConverteerMaandNummerNaarNaam(int maandNummer)
        {
            var maandNamen = new[]
            {
                "", "Januari", "Februari", "Maart", "April", "Mei", "Juni",
                "Juli", "Augustus", "September", "Oktober", "November", "December"
            };

            if (maandNummer >= 1 && maandNummer <= 12)
            {
                return maandNamen[maandNummer];
            }

            return "Onbekend";
        }

        /// <summary>
        /// Test de jaar/maand logica voor verschillende scenario's
        /// </summary>
        public static void TestJaarMaandLogica()
        {
            Console.WriteLine("=== Test Jaar/Maand Logica ===");
            Console.WriteLine($"Huidige datum: {DateTime.Now:dd-MM-yyyy}");
            Console.WriteLine($"Huidig boekjaar: {BepaalBoekjaar()}");
            Console.WriteLine($"Excel bestand: {ExcelBestandNaam}");
            Console.WriteLine($"Realisatie sheet: {RealisatieSheetNaam}");
            Console.WriteLine("");

            // Test verschillende maanden
            var testMaanden = new[] { "December", "Januari", "Februari", "Maart" };

            foreach (var maand in testMaanden)
            {
                var jaar = BepaalJaarVoorMaand(maand);
                var eersteVanMaand = BepaalEersteVanMaand(maand);
                Console.WriteLine($"🗓️ {maand}: Jaar={jaar}, Eerste van maand={eersteVanMaand:dd-MM-yyyy}");
            }

            Console.WriteLine("=== Test voltooid ===");
        }

        /// <summary>
        /// Geeft uitleg over de jaar/maand logica
        /// </summary>
        public static void ToonJaarMaandUitleg()
        {
            Console.WriteLine("📋 Jaar/Maand Logica Uitleg:");
            Console.WriteLine("");
            Console.WriteLine("🗓️ Factuur Timing:");
            Console.WriteLine("- Januari: Facturen voor December van vorig jaar");
            Console.WriteLine("- Februari+: Facturen voor vorige maand van huidig jaar");
            Console.WriteLine("");
            Console.WriteLine("📁 Bestand Selectie:");
            Console.WriteLine("- Januari 2026 → 'Uren 2025.xlsx' (December 2025 facturen)");
            Console.WriteLine("- Februari 2026 → 'Uren 2026.xlsx' (Januari 2026 facturen)");
            Console.WriteLine("- Maart 2026 → 'Uren 2026.xlsx' (Februari 2026 facturen)");
            Console.WriteLine("");
            Console.WriteLine("🎯 Sheet Namen:");
            Console.WriteLine("- Januari 2026 → 'Realisatie 2025'");
            Console.WriteLine("- Februari+ 2026 → 'Realisatie 2026'");
            Console.WriteLine("");
            Console.WriteLine($"✅ Huidige configuratie: {ExcelBestandNaam} → {RealisatieSheetNaam}");
        }

        /// <summary>
        /// Controleert of we in een overgangsperiode zitten (januari)
        /// </summary>
        public static bool IsOvergangsPeriode()
        {
            return DateTime.Now.Month == 1;
        }

        /// <summary>
        /// Geeft waarschuwing als we in overgangsperiode zitten
        /// </summary>
        public static void ControleerOvergangsPeriode()
        {
            if (IsOvergangsPeriode())
            {
                Console.WriteLine("⚠️ OVERGANGSPERIODE GEDETECTEERD:");
                Console.WriteLine($"   We zitten in januari {DateTime.Now.Year}");
                Console.WriteLine($"   Facturen voor december {DateTime.Now.Year - 1}");
                Console.WriteLine($"   Gebruikt bestand: {ExcelBestandNaam}");
                Console.WriteLine($"   Gebruikt sheet: {RealisatieSheetNaam}");
                Console.WriteLine("   Controleer of dit correct is!");
            }
        }
    }
}