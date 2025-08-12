using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using QuattroFacturatieProgramma.Helpers;
using QuattroFacturatieProgramma.ViewModels;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Geom;
using Path = System.IO.Path;

namespace QuattroFacturatieProgramma.Helpers
{
    public class FactuurService : IDisposable
    {
        private readonly IConfiguration _configuration;
        private readonly MollieApiHelper _mollieHelper;
        private readonly KlantHelper _klantHelper;

        public FactuurService(IConfiguration configuration, KlantHelper klantHelper = null)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _mollieHelper = new MollieApiHelper(configuration);
            _klantHelper = klantHelper;
        }

        /// <summary>
        /// Controleert of een factuur is betaald via Mollie
        /// </summary>
        /// <param name="paymentId">Het Mollie payment ID</param>
        /// <returns>True als betaald, false als niet betaald</returns>
        public async Task<bool> IsFactuurBetaaldAsync(string paymentId)
        {
            if (string.IsNullOrWhiteSpace(paymentId))
                return false;

            try
            {
                var status = await _mollieHelper.ControleerPaymentStatusAsync(paymentId);
                return status?.IsBetaald ?? false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fout bij controleren betalingsstatus: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Haalt uitgebreide betaalstatus op
        /// </summary>
        /// <param name="paymentId">Het Mollie payment ID</param>
        /// <returns>PaymentStatus object met alle details</returns>
        public async Task<PaymentStatus> GetBetalingsstatusAsync(string paymentId)
        {
            if (string.IsNullOrWhiteSpace(paymentId))
                return null;

            try
            {
                return await _mollieHelper.ControleerPaymentStatusAsync(paymentId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fout bij ophalen betalingsstatus: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Genereert een factuur PDF met QR-code voor betaling
        /// </summary>
        /// <param name="klantNaam">Naam van de klant</param>
        /// <param name="maand">Maand van de factuur</param>
        /// <param name="bedrag">Bedrag exclusief BTW</param>
        /// <param name="bestandsPad">Pad waar PDF wordt opgeslagen</param>
        /// <param name="factuurnummer">Factuurnummer</param>
        /// <param name="klantEmail">Email van klant (optioneel)</param>
        /// <returns>Payment ID van Mollie (null als EPC QR-code wordt gebruikt)</returns>
        public async Task<string> GenereerFactuurMetQrCodeAsync(
            string klantNaam,
            string maand,
            double bedrag,
            string bestandsPad,
            string factuurnummer,
            string klantEmail = null)
        {
            try
            {
                // Validatie
                if (string.IsNullOrWhiteSpace(klantNaam))
                    throw new ArgumentException("Klantnaam is verplicht", nameof(klantNaam));

                if (string.IsNullOrWhiteSpace(maand))
                    throw new ArgumentException("Maand is verplicht", nameof(maand));

                if (bedrag <= 0)
                    throw new ArgumentException("Bedrag moet groter zijn dan 0", nameof(bedrag));

                if (string.IsNullOrWhiteSpace(bestandsPad))
                    throw new ArgumentException("Bestandspad is verplicht", nameof(bestandsPad));

                if (string.IsNullOrWhiteSpace(factuurnummer))
                    throw new ArgumentException("Factuurnummer is verplicht", nameof(factuurnummer));

                // Zorg dat de directory bestaat
                var directory = Path.GetDirectoryName(bestandsPad);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var paymentId = await GenereerPDFFactuurAsync(klantNaam, maand, bedrag, bestandsPad, factuurnummer, klantEmail);

                return paymentId;
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij genereren factuur met QR-code: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Genereert PDF factuur met twee pagina's
        /// </summary>
        private async Task<string> GenereerPDFFactuurAsync(string klantNaam, string maand, double bedrag, string bestandsPad, string factuurnummer, string klantEmail = null)
        {
            using var stream = new FileStream(bestandsPad, FileMode.Create);
            using var writer = new PdfWriter(stream);
            using var pdf = new PdfDocument(writer);
            using var document = new Document(pdf, PageSize.A4);
            document.SetMargins(50, 50, 50, 50);

            // GEBRUIK NIEUWE FACTUURDATUM LOGICA
            // Factuurdatum = huidige datum (wanneer factuur wordt gemaakt)
            var eersteVanMaand = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            // Fonts
            var titelFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var bedrijfsFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var headerFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var kleinFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var accentFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            // BEREKEN TOTAAL UREN ÉÉN KEER
            double totaalUren = PdfHelper.BerekenTotaalUren(klantNaam, maand, _klantHelper);

            // Genereer QR-code voor betaling
            byte[] qrCodeBytes = null;
            string paymentId = null;

            try
            {
                // Probeer eerst Mollie API
                var (qrCode, molliePaymentId) = await _mollieHelper.CreeerPaymentEnQrCodeAsync(
                    (decimal)bedrag * 1.21m, // Inclusief BTW
                    factuurnummer,
                    klantEmail,
                    klantNaam);

                qrCodeBytes = qrCode;
                paymentId = molliePaymentId;
            }
            catch (Exception ex)
            {
                // Fallback naar EPC QR-code als Mollie faalt
                Console.WriteLine($"Mollie QR-code generatie gefaald, gebruik EPC fallback: {ex.Message}");
                try
                {
                    qrCodeBytes = QrBetalingHelper.GenereerBetalingsQrCode(
                        (decimal)bedrag * 1.21m,
                        "NL30RABO0347670407",
                        "Quattro Bouw & Vastgoed Advies BV",
                        $"Factuur {factuurnummer}");
                }
                catch (Exception epcEx)
                {
                    Console.WriteLine($"Ook EPC QR-code generatie gefaald: {epcEx.Message}");
                    qrCodeBytes = null;
                }
            }

            // Probeer logo te laden
            byte[] logoBytes = null;
            try
            {
                logoBytes = LogoHelper.LoadQuattroLogo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Kon logo niet laden: {ex.Message}");
            }

            // Genereer factuur pagina MET totaal uren
            await PdfHelper.MaakFactuurPaginaAsync(document, klantNaam, maand, bedrag, factuurnummer, eersteVanMaand,
                titelFont, bedrijfsFont, headerFont, normalFont, kleinFont, accentFont, qrCodeBytes, paymentId, logoBytes, _klantHelper, totaalUren);

            // Nieuwe pagina voor urenverantwoording
            document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

            // Genereer urenverantwoording pagina MET dezelfde totaal uren
            PdfHelper.MaakUrenverantwoordingPagina(document, maand, klantNaam, headerFont, normalFont, kleinFont, accentFont, _klantHelper, totaalUren);

            document.Close();

            return paymentId;
        }

        /// <summary>
        /// Test de QR-code functionaliteit
        /// </summary>
        public async Task TestQrCodeFunctionaliteitAsync()
        {
            try
            {
                Console.WriteLine("=== Test QR-code Functionaliteit ===");

                // Test EPC QR-code
                Console.WriteLine("1. Test EPC QR-code...");
                try
                {
                    var epcQr = QrBetalingHelper.GenereerBetalingsQrCode(
                        1250.75m,
                        "NL30RABO0347670407",
                        "Quattro Bouw & Vastgoed Advies BV",
                        "Test Factuur 2025_001"
                    );
                    Console.WriteLine($"   ✅ EPC QR-code gegenereerd ({epcQr.Length} bytes)");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ❌ EPC QR-code test gefaald: {ex.Message}");
                }

                // Test Mollie verbinding
                Console.WriteLine("2. Test Mollie verbinding...");
                try
                {
                    var verbindingOk = await _mollieHelper.TestVerbindingAsync();
                    if (verbindingOk)
                    {
                        Console.WriteLine("   ✅ Mollie API verbinding OK");

                        // Test beschikbare methoden
                        var methoden = await _mollieHelper.GetBeschikbareMethodenAsync();
                        Console.WriteLine($"   ✅ Beschikbare betaalmethoden: {string.Join(", ", methoden)}");
                    }
                    else
                    {
                        Console.WriteLine("   ❌ Mollie API verbinding gefaald");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ❌ Mollie verbinding test gefaald: {ex.Message}");
                }

                // Test Mollie QR-code (alleen als verbinding OK is)
                Console.WriteLine("3. Test Mollie QR-code...");
                try
                {
                    var (mollieQr, paymentId) = await _mollieHelper.CreeerPaymentEnQrCodeAsync(
                        1250.75m,
                        "TEST_2025_001",
                        "test@example.com",
                        "Test Klant"
                    );
                    Console.WriteLine($"   ✅ Mollie QR-code gegenereerd ({mollieQr.Length} bytes)");
                    Console.WriteLine($"   ✅ Payment ID: {paymentId}");

                    // Test status check
                    var status = await _mollieHelper.ControleerPaymentStatusAsync(paymentId);
                    Console.WriteLine($"   ✅ Status check: {status.Status}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ❌ Mollie QR-code test gefaald: {ex.Message}");
                }

                Console.WriteLine("=== Test voltooid ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Test gefaald: {ex.Message}");
            }
        }

        /// <summary>
        /// Test de volledige factuur generatie
        /// </summary>
        public async Task TestFactuurGeneratieAsync()
        {
            try
            {
                Console.WriteLine("=== Test Factuur Generatie ===");

                var testPad = Path.Combine(Path.GetTempPath(), "test_factuur.pdf");

                var paymentId = await GenereerFactuurMetQrCodeAsync(
                    "Test Klant BV",
                    "December",
                    1250.50,
                    testPad,
                    "TEST_2025_001",
                    "test@example.com"
                );

                if (File.Exists(testPad))
                {
                    var fileInfo = new FileInfo(testPad);
                    Console.WriteLine($"✅ Test factuur gegenereerd ({fileInfo.Length} bytes)");
                    Console.WriteLine($"✅ Bestand opgeslagen: {testPad}");

                    if (!string.IsNullOrEmpty(paymentId))
                    {
                        Console.WriteLine($"✅ Payment ID: {paymentId}");
                    }
                }
                else
                {
                    Console.WriteLine("❌ Test factuur niet gevonden");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Factuur generatie test gefaald: {ex.Message}");
            }
        }

        /// <summary>
        /// Test de uren berekening functionaliteit
        /// </summary>
        public void TestUrenBerekenging(string klantNaam, string maand)
        {
            try
            {
                Console.WriteLine("=== Test Uren Berekening ===");

                if (_klantHelper == null)
                {
                    Console.WriteLine("❌ KlantHelper niet beschikbaar voor test");
                    return;
                }

                var totaalUren = PdfHelper.BerekenTotaalUren(klantNaam, maand, _klantHelper);
                Console.WriteLine($"✅ Totaal uren voor {klantNaam} in {maand}: {totaalUren}");

                // Test met verschillende klanten
                var testKlanten = new[] { "Test Klant 1", "Test Klant 2", "Onbekende Klant" };
                foreach (var testKlant in testKlanten)
                {
                    var testUren = PdfHelper.BerekenTotaalUren(testKlant, maand, _klantHelper);
                    Console.WriteLine($"   {testKlant}: {testUren} uren");
                }

                Console.WriteLine("=== Test voltooid ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Uren berekening test gefaald: {ex.Message}");
            }
        }

        /// <summary>
        /// Genereert een factuur met gebruik van de gecentraliseerde uren berekening
        /// </summary>
        public async Task<string> GenereerFactuurMetCentraleUrenBerekenging(
            string klantNaam,
            string maand,
            double bedrag,
            string bestandsPad,
            string factuurnummer,
            string klantEmail = null)
        {
            try
            {
                Console.WriteLine($"🔄 Start factuur generatie voor {klantNaam} - {maand}");

                // Validatie
                ValideerFactuurParameters(klantNaam, maand, bedrag, bestandsPad, factuurnummer);

                // Zorg dat de directory bestaat
                var directory = Path.GetDirectoryName(bestandsPad);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                    Console.WriteLine($"📁 Directory aangemaakt: {directory}");
                }

                // CENTRALE UREN BEREKENING - dit is het belangrijkste deel
                double totaalUren = PdfHelper.BerekenTotaalUren(klantNaam, maand, _klantHelper);
                Console.WriteLine($"⏱️ Totaal uren berekend: {totaalUren}");

                // Genereer de PDF met de berekende uren
                var paymentId = await GenereerPDFMetBerekendeUren(
                    klantNaam, maand, bedrag, bestandsPad, factuurnummer, klantEmail, totaalUren);

                Console.WriteLine($"✅ Factuur gegenereerd: {bestandsPad}");
                return paymentId;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij genereren factuur: {ex.Message}");
                throw new Exception($"Fout bij genereren factuur met centrale uren berekening: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Valideert de factuur parameters
        /// </summary>
        private void ValideerFactuurParameters(string klantNaam, string maand, double bedrag, string bestandsPad, string factuurnummer)
        {
            if (string.IsNullOrWhiteSpace(klantNaam))
                throw new ArgumentException("Klantnaam is verplicht", nameof(klantNaam));

            if (string.IsNullOrWhiteSpace(maand))
                throw new ArgumentException("Maand is verplicht", nameof(maand));

            if (bedrag <= 0)
                throw new ArgumentException("Bedrag moet groter zijn dan 0", nameof(bedrag));

            if (string.IsNullOrWhiteSpace(bestandsPad))
                throw new ArgumentException("Bestandspad is verplicht", nameof(bestandsPad));

            if (string.IsNullOrWhiteSpace(factuurnummer))
                throw new ArgumentException("Factuurnummer is verplicht", nameof(factuurnummer));
        }

        /// <summary>
        /// Genereert PDF met vooraf berekende uren
        /// </summary>
        private async Task<string> GenereerPDFMetBerekendeUren(
            string klantNaam, string maand, double bedrag, string bestandsPad,
            string factuurnummer, string klantEmail, double totaalUren)
        {
            using var stream = new FileStream(bestandsPad, FileMode.Create);
            using var writer = new PdfWriter(stream);
            using var pdf = new PdfDocument(writer);
            using var document = new Document(pdf, PageSize.A4);
            document.SetMargins(50, 50, 50, 50);

            // GEBRUIK NIEUWE FACTUURDATUM LOGICA
            // Factuurdatum = huidige datum (wanneer factuur wordt gemaakt)
            var eersteVanMaand = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            // Fonts
            var titelFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var bedrijfsFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var headerFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var kleinFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var accentFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            // Genereer QR-code voor betaling
            byte[] qrCodeBytes = null;
            string paymentId = null;

            try
            {
                // Probeer eerst Mollie API
                var (qrCode, molliePaymentId) = await _mollieHelper.CreeerPaymentEnQrCodeAsync(
                    (decimal)bedrag * 1.21m, // Inclusief BTW
                    factuurnummer,
                    klantEmail,
                    klantNaam);

                qrCodeBytes = qrCode;
                paymentId = molliePaymentId;
                Console.WriteLine($"💳 Mollie QR-code gegenereerd voor payment: {paymentId}");
            }
            catch (Exception ex)
            {
                // Fallback naar EPC QR-code als Mollie faalt
                Console.WriteLine($"⚠️ Mollie QR-code generatie gefaald, gebruik EPC fallback: {ex.Message}");
                try
                {
                    qrCodeBytes = QrBetalingHelper.GenereerBetalingsQrCode(
                        (decimal)bedrag * 1.21m,
                        "NL30RABO0347670407",
                        "Quattro Bouw & Vastgoed Advies BV",
                        $"Factuur {factuurnummer}");
                    Console.WriteLine("💳 EPC QR-code gegenereerd als fallback");
                }
                catch (Exception epcEx)
                {
                    Console.WriteLine($"❌ Ook EPC QR-code generatie gefaald: {epcEx.Message}");
                    qrCodeBytes = null;
                }
            }

            // Probeer logo te laden
            byte[] logoBytes = null;
            try
            {
                logoBytes = LogoHelper.LoadQuattroLogo();
                Console.WriteLine("🎨 Logo geladen");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Kon logo niet laden: {ex.Message}");
            }

            // Genereer factuur pagina MET vooraf berekende totaal uren
            await PdfHelper.MaakFactuurPaginaAsync(document, klantNaam, maand, bedrag, factuurnummer, eersteVanMaand,
                titelFont, bedrijfsFont, headerFont, normalFont, kleinFont, accentFont, qrCodeBytes, paymentId, logoBytes, _klantHelper, totaalUren);

            // Nieuwe pagina voor urenverantwoording
            document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

            // Genereer urenverantwoording pagina MET dezelfde vooraf berekende totaal uren
            PdfHelper.MaakUrenverantwoordingPagina(document, maand, klantNaam, headerFont, normalFont, kleinFont, accentFont, _klantHelper, totaalUren);

            document.Close();

            return paymentId;
        }

        /// <summary>
        /// Geeft statistieken over de uren berekening
        /// </summary>
        public void ToonUrenStatistieken(string maand)
        {
            if (_klantHelper == null)
            {
                Console.WriteLine("❌ KlantHelper niet beschikbaar voor statistieken");
                return;
            }

            try
            {
                Console.WriteLine($"=== Uren Statistieken voor {maand} ===");

                // Dit zou je kunnen uitbreiden met meer statistieken
                // Bijvoorbeeld: totaal uren per klant, gemiddelde uren, etc.

                Console.WriteLine("📊 Uren statistieken functionaliteit beschikbaar");
                Console.WriteLine("   - Gebruik BerekenTotaalUren() voor specifieke klanten");
                Console.WriteLine("   - Alle berekeningen worden gecentraliseerd uitgevoerd");
                Console.WriteLine("   - Consistente resultaten gegarandeerd op beide PDF pagina's");

                Console.WriteLine("=== Statistieken voltooid ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij uren statistieken: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _mollieHelper?.Dispose();
        }
    }
}