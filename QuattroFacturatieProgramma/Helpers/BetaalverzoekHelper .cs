using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using QRCoder;

namespace QuattroFacturatieProgramma.Helpers
{
    /// <summary>
    /// Betaalverzoek.nl API Helper
    /// Nederlandse payment request service die werkt met alle banken
    /// </summary>
    public class BetaalverzoekHelper : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _baseUrl = "https://api.betaalverzoek.nl/v1";
        private readonly IConfiguration _configuration;

        public BetaalverzoekHelper(IConfiguration configuration)
        {
            _configuration = configuration;
            _apiKey = configuration["Betaalverzoek:ApiKey"] ?? throw new InvalidOperationException("Betaalverzoek API key niet gevonden");

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "QuattroFacturatieProgramma/1.0");
        }

        /// <summary>
        /// Creëert een betaalverzoek met lange geldigheid
        /// </summary>
        /// <param name="bedrag">Bedrag inclusief BTW</param>
        /// <param name="omschrijving">Omschrijving</param>
        /// <param name="ontvangerIban">Jouw IBAN</param>
        /// <param name="ontvangerNaam">Jouw bedrijfsnaam</param>
        /// <param name="geldigDagen">Dagen geldig (max 365)</param>
        /// <returns>Tuple met QR-code en betaalverzoek ID</returns>
        public async Task<(byte[] QrCode, string BetaalverzoekId)> CreeerBetaalverzoekAsync(
            decimal bedrag,
            string omschrijving,
            string ontvangerIban,
            string ontvangerNaam,
            int geldigDagen = 30)
        {
            try
            {
                Console.WriteLine($"🔄 Betaalverzoek.nl - €{bedrag:F2}, {geldigDagen} dagen geldig");

                var request = new
                {
                    amount = bedrag,
                    currency = "EUR",
                    description = omschrijving,
                    creditor = new
                    {
                        name = ontvangerNaam,
                        iban = ontvangerIban
                    },
                    validUntil = DateTime.UtcNow.AddDays(geldigDagen).ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    callbackUrl = "https://quattrobouwenenvastgoedadvies.nl/webhook/betaalverzoek", // Optioneel
                    reference = $"QUATTRO_{DateTime.Now:yyyyMMdd}_{Guid.NewGuid().ToString("N")[..8]}"
                };

                var json = JsonSerializer.Serialize(request, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync($"{_baseUrl}/payment-requests", content);

                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"Betaalverzoek API fout: {response.StatusCode} - {responseContent}");
                }

                var responseData = JsonSerializer.Deserialize<JsonElement>(responseContent);

                var betaalverzoekId = responseData.GetProperty("id").GetString();
                var paymentUrl = responseData.GetProperty("paymentUrl").GetString();

                Console.WriteLine($"✅ Betaalverzoek succesvol - ID: {betaalverzoekId}");
                Console.WriteLine($"🔗 Payment URL: {paymentUrl}");

                // Genereer QR-code
                var qrCodeBytes = GenereerQrCodeVoorUrl(paymentUrl);

                return (qrCodeBytes, betaalverzoekId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Betaalverzoek fout: {ex.Message}");
                throw new Exception($"Betaalverzoek.nl fout: {ex.Message}", ex);
            }
        }

        private byte[] GenereerQrCodeVoorUrl(string url)
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.M);
            using var qrCode = new PngByteQRCode(qrCodeData);
            return qrCode.GetGraphic(20);
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}