using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using QRCoder;

namespace QuattroFacturatieProgramma.Helpers
{
    public class RabobankR2PStatus
    {
        public string Status { get; set; } = "";
        public bool IsBetaald => Status == "ACCEPTED" || Status == "PAID";
        public string RequestId { get; set; } = "";
        public decimal? Bedrag { get; set; }
        public DateTime? BetaalDatum { get; set; }
        public string Omschrijving { get; set; } = "";
        public string PaymentUrl { get; set; } = "";
    }

    /// <summary>
    /// Rabobank Request to Pay API Helper
    /// Vergelijkbaar met Tikkie maar dan van Rabobank
    /// </summary>
    public class RabobankR2PHelper : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _baseUrl = "https://api.rabobank.nl/openapi/sandbox/payments/request-to-pay/v1";
        private readonly IConfiguration _configuration;

        public RabobankR2PHelper(IConfiguration configuration)
        {
            _configuration = configuration;
            _apiKey = configuration["Rabobank:ApiKey"] ?? throw new InvalidOperationException("Rabobank API key niet gevonden");

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("X-IBM-Client-Id", _apiKey);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "QuattroFacturatieProgramma/1.0");
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        /// <summary>
        /// Creëert een Request to Pay en genereert QR-code
        /// </summary>
        /// <param name="bedrag">Bedrag inclusief BTW</param>
        /// <param name="omschrijving">Omschrijving van de betaling</param>
        /// <param name="creditorIban">Jouw Rabobank IBAN</param>
        /// <param name="geldigDagen">Dagen geldig (max 90)</param>
        /// <returns>Tuple met QR-code bytes en request ID</returns>
        public async Task<(byte[] QrCode, string RequestId)> CreeerRequestToPayAsync(
            decimal bedrag,
            string omschrijving,
            string creditorIban,
            int geldigDagen = 30)
        {
            try
            {
                Console.WriteLine($"🔄 Rabobank R2P aanroep - €{bedrag:F2}, geldig {geldigDagen} dagen");

                var requestId = Guid.NewGuid().ToString();
                var expiryDate = DateTime.UtcNow.AddDays(geldigDagen);

                var paymentRequest = new
                {
                    requestId = requestId,
                    creditor = new
                    {
                        iban = creditorIban,
                        name = "Quattro Bouw & Vastgoed Advies BV"
                    },
                    paymentInformation = new
                    {
                        instructedAmount = new
                        {
                            currency = "EUR",
                            amount = bedrag.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)
                        },
                        remittanceInformation = new
                        {
                            unstructured = omschrijving
                        }
                    },
                    requestedExecutionDate = DateTime.UtcNow.ToString("yyyy-MM-dd"),
                    expiryDateTime = expiryDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                var json = JsonSerializer.Serialize(paymentRequest, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    WriteIndented = true
                });

                Console.WriteLine($"📤 R2P request: {json}");

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync($"{_baseUrl}/payment-requests", content);

                var responseContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"📥 R2P response ({response.StatusCode}): {responseContent}");

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"Rabobank R2P fout: {response.StatusCode} - {responseContent}");
                }

                var paymentResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);

                var returnedRequestId = paymentResponse.GetProperty("requestId").GetString();
                var paymentUrl = paymentResponse.GetProperty("_links").GetProperty("qrCode").GetProperty("href").GetString();

                Console.WriteLine($"✅ R2P succesvol - ID: {returnedRequestId}");
                Console.WriteLine($"🔗 Payment URL: {paymentUrl}");

                // Genereer QR-code
                var qrCodeBytes = GenereerQrCodeVoorUrl(paymentUrl);

                return (qrCodeBytes, returnedRequestId ?? requestId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ R2P fout: {ex.Message}");
                throw new Exception($"Rabobank Request to Pay fout: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Controleert status van Request to Pay
        /// </summary>
        public async Task<RabobankR2PStatus> ControleerR2PStatusAsync(string requestId)
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/payment-requests/{requestId}");

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"R2P status fout: {response.StatusCode} - {errorContent}");
                }

                var responseJson = await response.Content.ReadAsStringAsync();
                var statusData = JsonSerializer.Deserialize<JsonElement>(responseJson);

                var status = new RabobankR2PStatus
                {
                    RequestId = statusData.GetProperty("requestId").GetString() ?? "",
                    Status = statusData.GetProperty("status").GetString() ?? ""
                };

                // Bedrag ophalen
                if (statusData.TryGetProperty("paymentInformation", out var paymentInfo) &&
                    paymentInfo.TryGetProperty("instructedAmount", out var amount) &&
                    amount.TryGetProperty("amount", out var amountValue))
                {
                    if (decimal.TryParse(amountValue.GetString(), out var bedrag))
                        status.Bedrag = bedrag;
                }

                return status;
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij R2P status check: {ex.Message}", ex);
            }
        }

        private byte[] GenereerQrCodeVoorUrl(string url)
        {
            try
            {
                using var qrGenerator = new QRCodeGenerator();
                using var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.M);
                using var qrCode = new PngByteQRCode(qrCodeData);

                return qrCode.GetGraphic(20);
            }
            catch (Exception ex)
            {
                throw new Exception($"QR-code generatie fout: {ex.Message}", ex);
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}