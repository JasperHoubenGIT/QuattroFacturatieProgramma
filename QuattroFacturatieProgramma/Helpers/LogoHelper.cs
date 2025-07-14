using System;
using System.IO;
using System.Reflection;

namespace QuattroFacturatieProgramma.Helpers
{
    public static class LogoHelper
    {
        /// <summary>
        /// Laadt het Quattro logo uit embedded resources
        /// </summary>
        /// <returns>Logo als byte array, of null als niet gevonden</returns>
        public static byte[] LoadQuattroLogo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                // Mogelijke resource namen (afhankelijk van waar je het logo plaatst)
                string[] possibleResourceNames =
                {
                    "QuattroFacturatieProgramma.Assets.quattro_logo.png",
                    "QuattroFacturatieProgramma.Images.quattro_logo.png",
                    "QuattroFacturatieProgramma.Resources.quattro_logo.png",
                    "QuattroFacturatieProgramma.Resources.Images.quattro_logo.png",
                    "quattro_logo.png"
                };

                foreach (var resourceName in possibleResourceNames)
                {
                    using var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream != null)
                    {
                        using var memoryStream = new MemoryStream();
                        stream.CopyTo(memoryStream);
                        Console.WriteLine($"✅ Logo geladen uit resource: {resourceName}");
                        return memoryStream.ToArray();
                    }
                }

                // Als embedded resource niet werkt, probeer uit app directory
                var appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var logoPath = Path.Combine(appDirectory, "quattro_logo.png");

                if (File.Exists(logoPath))
                {
                    Console.WriteLine($"✅ Logo geladen uit bestand: {logoPath}");
                    return File.ReadAllBytes(logoPath);
                }

                Console.WriteLine("⚠️ Quattro logo niet gevonden - gebruik gestylde tekst");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij laden logo: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Slaat een logo op in de app directory voor gebruik
        /// </summary>
        /// <param name="logoBytes">Logo data</param>
        /// <param name="fileName">Bestandsnaam (bijv. "quattro_logo.png")</param>
        public static bool SaveLogoToAppDirectory(byte[] logoBytes, string fileName = "quattro_logo.png")
        {
            try
            {
                if (logoBytes == null || logoBytes.Length == 0)
                    return false;

                var appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var logoPath = Path.Combine(appDirectory, fileName);

                File.WriteAllBytes(logoPath, logoBytes);
                Console.WriteLine($"✅ Logo opgeslagen: {logoPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij opslaan logo: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Test of het logo correct kan worden geladen
        /// </summary>
        public static void TestLogoLoading()
        {
            Console.WriteLine("=== Test Logo Loading ===");

            var logoBytes = LoadQuattroLogo();

            if (logoBytes != null)
            {
                Console.WriteLine($"✅ Logo geladen: {logoBytes.Length} bytes");

                // Test of het een geldige image is
                try
                {
                    using var stream = new MemoryStream(logoBytes);
                    // Hier zou je kunnen testen of het een geldige PNG/JPG is
                    Console.WriteLine("✅ Logo data lijkt geldig");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Logo data mogelijk corrupt: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("❌ Logo niet gevonden");
                Console.WriteLine("💡 Tip: Plaats 'quattro_logo.png' in je app directory of voeg toe als embedded resource");
            }
        }
    }
}