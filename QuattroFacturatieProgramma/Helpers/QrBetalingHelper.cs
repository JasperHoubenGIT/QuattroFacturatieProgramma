using System;
using System.Globalization;
using System.Text;
using QRCoder;

namespace QuattroFacturatieProgramma.Helpers
{
    public static class QrBetalingHelper
    {
        /// <summary>
        /// Genereert een EPC QR-code voor SEPA betalingen
        /// </summary>
        /// <param name="bedrag">Het te betalen bedrag</param>
        /// <param name="iban">IBAN van de ontvanger</param>
        /// <param name="ontvanger">Naam van de ontvanger</param>
        /// <param name="omschrijving">Omschrijving van de betaling</param>
        /// <returns>QR-code als byte array (PNG format)</returns>
        public static byte[] GenereerBetalingsQrCode(decimal bedrag, string iban, string ontvanger, string omschrijving)
        {
            try
            {
                // Validatie
                if (bedrag <= 0)
                    throw new ArgumentException("Bedrag moet groter zijn dan 0");

                if (string.IsNullOrWhiteSpace(iban))
                    throw new ArgumentException("IBAN is verplicht");

                if (string.IsNullOrWhiteSpace(ontvanger))
                    throw new ArgumentException("Ontvanger naam is verplicht");

                // Clean IBAN (remove spaces)
                iban = iban.Replace(" ", "").ToUpperInvariant();

                // Valideer IBAN format (basis check)
                if (!IsValidIbanFormat(iban))
                    throw new ArgumentException("Ongeldig IBAN format");

                // Maak EPC QR-code data volgens European Payments Council standaard
                var epcData = CreateEpcQrData(bedrag, iban, ontvanger, omschrijving ?? "");

                // Genereer QR-code
                using var qrGenerator = new QRCodeGenerator();
                using var qrCodeData = qrGenerator.CreateQrCode(epcData, QRCodeGenerator.ECCLevel.M);
                using var qrCode = new PngByteQRCode(qrCodeData);

                return qrCode.GetGraphic(20); // 20 pixels per module
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij genereren QR-code: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Maakt EPC QR-code data string volgens standaard
        /// </summary>
        private static string CreateEpcQrData(decimal bedrag, string iban, string ontvanger, string omschrijving)
        {
            var sb = new StringBuilder();

            // EPC QR-code format versie 002
            sb.AppendLine("BCD");              // Service Tag
            sb.AppendLine("002");              // Version
            sb.AppendLine("1");                // Character set (UTF-8)
            sb.AppendLine("SCT");              // Identification (SEPA Credit Transfer)
            sb.AppendLine("");                 // BIC (empty for domestic transfers)
            sb.AppendLine(TruncateField(ontvanger, 70));           // Beneficiary Name (max 70 chars)
            sb.AppendLine(iban);               // Beneficiary Account (IBAN)
            sb.AppendLine($"EUR{bedrag.ToString("F2", CultureInfo.InvariantCulture)}"); // Amount
            sb.AppendLine("");                 // Purpose (empty)
            sb.AppendLine(TruncateField(omschrijving, 140));       // Remittance Information (max 140 chars)
            sb.Append("");                     // Beneficiary to originator information (empty, no newline)

            return sb.ToString();
        }

        /// <summary>
        /// Truncate field to maximum length (UTF-8 safe)
        /// </summary>
        private static string TruncateField(string input, int maxLength)
        {
            if (string.IsNullOrEmpty(input))
                return "";

            if (input.Length <= maxLength)
                return input;

            return input.Substring(0, maxLength);
        }

        /// <summary>
        /// Basis IBAN format validatie
        /// </summary>
        private static bool IsValidIbanFormat(string iban)
        {
            if (string.IsNullOrWhiteSpace(iban))
                return false;

            // IBAN moet tussen 15 en 34 karakters zijn
            if (iban.Length < 15 || iban.Length > 34)
                return false;

            // Eerste 2 karakters moeten letters zijn (landcode)
            if (!char.IsLetter(iban[0]) || !char.IsLetter(iban[1]))
                return false;

            // Derde en vierde karakter moeten cijfers zijn (check digits)
            if (!char.IsDigit(iban[2]) || !char.IsDigit(iban[3]))
                return false;

            // Rest moet alphanumeriek zijn
            for (int i = 4; i < iban.Length; i++)
            {
                if (!char.IsLetterOrDigit(iban[i]))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Test functie voor QR-code generatie
        /// </summary>
        public static void TestQrCodeGeneratie()
        {
            try
            {
                Console.WriteLine("=== Test QR-code Generatie ===");

                var qrBytes = GenereerBetalingsQrCode(
                    1250.75m,
                    "NL30RABO0347670407",
                    "Quattro Bouw & Vastgoed Advies BV",
                    "Test Factuur 2025_001"
                );

                Console.WriteLine($"✅ QR-code succesvol gegenereerd ({qrBytes.Length} bytes)");

                // Optioneel: sla QR-code op voor visuele controle
                var testPath = Path.Combine(Path.GetTempPath(), "test_qr_code.png");
                File.WriteAllBytes(testPath, qrBytes);
                Console.WriteLine($"✅ Test QR-code opgeslagen: {testPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ QR-code test gefaald: {ex.Message}");
            }
        }
    }
}