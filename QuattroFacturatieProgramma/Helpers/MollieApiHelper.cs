using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using QRCoder;

namespace QuattroFacturatieProgramma.Helpers
{
    public class PaymentStatus
    {
        public string Status { get; set; } = "";
        public bool IsBetaald => Status == "paid";
        public string PaymentId { get; set; } = "";
        public decimal? Bedrag { get; set; }
        public DateTime? BetaalDatum { get; set; }
        public string Omschrijving { get; set; } = "";
    }

    public class MollieApiHelper : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _baseUrl = "https://api.mollie.com/v2";

        public MollieApiHelper(IConfiguration configuration)
        {
            _apiKey = configuration["Mollie:ApiKey"] ?? throw new InvalidOperationException("Mollie API key niet gevonden in configuratie");

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "QuattroFacturatieProgramma/1.0");
        }

        /// <summary>
        /// Creëert een Mollie payment en genereert QR-code
        /// </summary>
        /// <param name="bedrag">Bedrag inclusief BTW</param>
        /// <param name="factuurnummer">Factuurnummer voor referentie</param>
        /// <param name="klantEmail">Email van klant (optioneel)</param>
        /// <param name="klantNaam">Naam van klant</param>
        /// <returns>Tuple met QR-code bytes en payment ID</returns>
        public async Task<(byte[] QrCode, string PaymentId)> CreeerPaymentEnQrCodeAsync(
            decimal bedrag,
            string factuurnummer,
            string klantEmail = null,
            string klantNaam = null)
        {
            try
            {
                // Validatie
                if (bedrag <= 0)
                    throw new ArgumentException("Bedrag moet groter zijn dan 0");

                if (string.IsNullOrWhiteSpace(factuurnummer))
                    throw new ArgumentException("Factuurnummer is verplicht");

                // Maak payment request
                var paymentRequest = new
                {
                    amount = new
                    {
                        currency = "EUR",
                        value = bedrag.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)
                    },
                    description = $"Factuur {factuurnummer}",
                    redirectUrl = "https://quattrobouwenenvastgoedadvies.nl/betaling-voltooid",
                    webhookUrl = "https://quattrobouwenenvastgoedadvies.nl/webhook/mollie", // Optioneel
                    metadata = new
                    {
                        factuurnummer = factuurnummer,
                        klant_naam = klantNaam ?? "",
                        klant_email = klantEmail ?? ""
                    },
                    method = new[] { "ideal", "bancontact", "sofort", "creditcard" } // Toegestane betaalmethoden
                };

                var json = JsonSerializer.Serialize(paymentRequest);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // Verstuur request naar Mollie
                var response = await _httpClient.PostAsync($"{_baseUrl}/payments", content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"Mollie API fout: {response.StatusCode} - {errorContent}");
                }

                var responseJson = await response.Content.ReadAsStringAsync();
                var paymentResponse = JsonSerializer.Deserialize<JsonElement>(responseJson);

                // Haal payment URL en ID op
                var paymentUrl = paymentResponse.GetProperty("_links").GetProperty("checkout").GetProperty("href").GetString();
                var paymentId = paymentResponse.GetProperty("id").GetString();

                if (string.IsNullOrEmpty(paymentUrl) || string.IsNullOrEmpty(paymentId))
                    throw new InvalidOperationException("Ongeldig response van Mollie API");

                // Genereer QR-code voor de payment URL
                var qrCodeBytes = GenereerQrCodeVoorUrl(paymentUrl);

                return (qrCodeBytes, paymentId);
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij creëren Mollie payment: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Controleert de status van een payment
        /// </summary>
        /// <param name="paymentId">Mollie payment ID</param>
        /// <returns>PaymentStatus object</returns>
        public async Task<PaymentStatus> ControleerPaymentStatusAsync(string paymentId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(paymentId))
                    throw new ArgumentException("Payment ID is verplicht");

                var response = await _httpClient.GetAsync($"{_baseUrl}/payments/{paymentId}");

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"Mollie API fout: {response.StatusCode} - {errorContent}");
                }

                var responseJson = await response.Content.ReadAsStringAsync();
                var paymentData = JsonSerializer.Deserialize<JsonElement>(responseJson);

                var status = new PaymentStatus
                {
                    PaymentId = paymentData.GetProperty("id").GetString() ?? "",
                    Status = paymentData.GetProperty("status").GetString() ?? "",
                    Omschrijving = paymentData.GetProperty("description").GetString() ?? ""
                };

                // Bedrag ophalen
                if (paymentData.TryGetProperty("amount", out var amountProperty) &&
                    amountProperty.TryGetProperty("value", out var valueProperty))
                {
                    if (decimal.TryParse(valueProperty.GetString(), out var bedrag))
                        status.Bedrag = bedrag;
                }

                // Betaaldatum ophalen (indien betaald)
                if (status.IsBetaald && paymentData.TryGetProperty("paidAt", out var paidAtProperty))
                {
                    if (DateTime.TryParse(paidAtProperty.GetString(), out var paidAt))
                        status.BetaalDatum = paidAt;
                }

                return status;
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij controleren payment status: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Genereert QR-code voor een URL
        /// </summary>
        private byte[] GenereerQrCodeVoorUrl(string url)
        {
            try
            {
                using var qrGenerator = new QRCodeGenerator();
                using var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.M);
                using var qrCode = new PngByteQRCode(qrCodeData);

                return qrCode.GetGraphic(20); // 20 pixels per module
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij genereren QR-code voor URL: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Test de Mollie API verbinding
        /// </summary>
        public async Task<bool> TestVerbindingAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/methods");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Haalt beschikbare betaalmethoden op
        /// </summary>
        public async Task<string[]> GetBeschikbareMethodenAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/methods");

                if (!response.IsSuccessStatusCode)
                    return Array.Empty<string>();

                var responseJson = await response.Content.ReadAsStringAsync();
                var methodsData = JsonSerializer.Deserialize<JsonElement>(responseJson);

                var methods = new List<string>();

                if (methodsData.TryGetProperty("_embedded", out var embedded) &&
                    embedded.TryGetProperty("methods", out var methodsArray))
                {
                    foreach (var method in methodsArray.EnumerateArray())
                    {
                        if (method.TryGetProperty("id", out var idProperty))
                        {
                            var methodId = idProperty.GetString();
                            if (!string.IsNullOrEmpty(methodId))
                                methods.Add(methodId);
                        }
                    }
                }

                return methods.ToArray();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}