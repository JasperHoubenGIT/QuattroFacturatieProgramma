using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
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
        private readonly IConfiguration _configuration;
        private readonly string _baseUrl = "https://api.mollie.com/v2";

        public MollieApiHelper(IConfiguration configuration)
        {
            _configuration = configuration;
            _apiKey = configuration["Mollie:ApiKey"] ?? throw new InvalidOperationException("Mollie API key niet gevonden in configuratie");
            
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "QuattroFacturatieProgramma/1.0");
            
            // Langere timeout voor live API
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }
        public async Task<(byte[] QrCode, string PaymentLinkId)> CreeerPaymentLinkEnQrCodeAsync(
            decimal bedrag,
            string factuurnummer,
            string? klantEmail = null,
            string? klantNaam = null)
        {
            try
            {
                if (bedrag <= 0)
                    throw new ArgumentException("Bedrag moet groter zijn dan 0");

                if (string.IsNullOrWhiteSpace(factuurnummer))
                    throw new ArgumentException("Factuurnummer is verplicht");

                var request = new
                {
                    description = $"Factuur {factuurnummer}",
                    amount = new
                    {
                        currency = "EUR",
                        value = bedrag.ToString("F2", CultureInfo.InvariantCulture)
                    },
                    redirectUrl = "https://jouwdomein.nl/bedankt", // Pas aan naar jouw omgeving
                    webhookUrl = "https://jouwdomein.nl/webhook", // Optioneel
                    metadata = new
                    {
                        factuurnummer,
                        klant_naam = klantNaam ?? "",
                        klant_email = klantEmail ?? ""
                    },
                    expiresAt = DateTime.UtcNow.AddDays(15).ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                var json = JsonSerializer.Serialize(request, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync($"{_baseUrl}/payment-links", content);
                var responseJson = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Mollie Payment Link error: {response.StatusCode} - {responseJson}");
                    throw new Exception($"Mollie API fout: {response.StatusCode} - {responseJson}");
                }

                var data = JsonSerializer.Deserialize<JsonElement>(responseJson);
                var linkId = data.GetProperty("id").GetString();
                var linkUrl = data.GetProperty("url").GetString();

                if (string.IsNullOrEmpty(linkId) || string.IsNullOrEmpty(linkUrl))
                    throw new Exception("Kon payment link niet ophalen uit response");

                var qrCode = GenereerQrCodeVoorUrl(linkUrl);

                return (qrCode, linkId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij creëren payment link: {ex.Message}");
                throw;
            }
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
    string? klantEmail = null,
    string? klantNaam = null)
        {
            try
            {
                // Validatie
                if (bedrag <= 0)
                    throw new ArgumentException("Bedrag moet groter zijn dan 0");

                if (string.IsNullOrWhiteSpace(factuurnummer))
                    throw new ArgumentException("Factuurnummer is verplicht");

                Console.WriteLine($"🔄 Mollie API aanroep - Bedrag: €{bedrag:F2}, Factuur: {factuurnummer}");
                Console.WriteLine($"🔑 API Key type: {(_apiKey.StartsWith("live_") ? "LIVE" : "TEST")}");

                var paymentRequest = new
                {
                    amount = new
                    {
                        currency = "EUR",
                        value = bedrag.ToString("F2", CultureInfo.InvariantCulture)
                    },
                    description = $"Factuur {factuurnummer}",
                    redirectUrl = "https://quattrobouwenenvastgoedadvies.nl/betaling-voltooid",
                    metadata = new
                    {
                        factuurnummer = factuurnummer,
                        klant_naam = klantNaam ?? "",
                        klant_email = klantEmail ?? ""
                    },
                    locale = "nl_NL"
                    // ⚠️ GEEN "method" opnemen = betaling blijft lang geldig
                };

                var json = JsonSerializer.Serialize(paymentRequest, new JsonSerializerOptions
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                Console.WriteLine($"📤 Mollie request: {json}");

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync($"{_baseUrl}/payments", content);

                Console.WriteLine($"📡 Mollie response status: {response.StatusCode}");

                var responseContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"📥 Mollie response: {responseContent}");

                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"Mollie API fout: {response.StatusCode} - {responseContent}");

                var paymentResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);
                var paymentId = paymentResponse.GetProperty("id").GetString();
                var checkoutUrl = paymentResponse.GetProperty("_links").GetProperty("checkout").GetProperty("href").GetString();

                if (string.IsNullOrEmpty(paymentId) || string.IsNullOrEmpty(checkoutUrl))
                    throw new InvalidOperationException("Ongeldig response van Mollie API");

                Console.WriteLine($"✅ Payment aangemaakt - ID: {paymentId}");
                Console.WriteLine($"🔗 Checkout URL: {checkoutUrl}");

                // Genereer QR-code
                var qrCodeBytes = GenereerQrCodeVoorUrl(checkoutUrl);
                Console.WriteLine($"📱 QR code gegenereerd: {qrCodeBytes?.Length ?? 0} bytes");

                return (qrCodeBytes, paymentId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij Mollie payment: {ex.Message}");
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
                var qrGenerator = new QRCodeGenerator();
                var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.M);
                var qrCode = new PngByteQRCode(qrCodeData);
                
                var pixelsPerModule = _configuration["QrCode:PixelsPerModule"] != null 
                    ? int.Parse(_configuration["QrCode:PixelsPerModule"]!) 
                    : 20;
                return qrCode.GetGraphic(pixelsPerModule);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ QR code generatie fout: {ex.Message}");
                throw new Exception($"QR code generatie fout: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Test de Mollie API verbinding
        /// </summary>
        public async Task<bool> TestVerbindingAsync()
        {
            try
            {
                Console.WriteLine("🔄 Test Mollie API verbinding...");
                var response = await _httpClient.GetAsync($"{_baseUrl}/methods");
                Console.WriteLine($"📡 API test response: {response.StatusCode}");
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"✅ API verbinding succesvol");
                    return true;
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"❌ API verbinding gefaald: {response.StatusCode} - {errorContent}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ API verbinding exception: {ex.Message}");
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