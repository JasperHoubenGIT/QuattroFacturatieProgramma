using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeOpenXml;
using QuattroFacturatieProgramma.Models;
using QuattroFacturatieProgramma.Helpers;
using QuattroFacturatieProgramma.Views;
using Microsoft.Extensions.Configuration;
using Application = Microsoft.Maui.Controls.Application;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Geom;
using Path = System.IO.Path;

namespace QuattroFacturatieProgramma.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly Dictionary<string, List<string>> _klantenPerMaand = new();
        private readonly Dictionary<string, Dictionary<string, double>> _klantenBedragen = new();
        private string _huidigeExcelPad = "";
        private readonly IConfiguration _configuration;
        private readonly KlantHelper _klantHelper;

        [ObservableProperty]
        private ObservableCollection<string> _maanden = new()
        {
            "Januari", "Februari", "Maart", "April", "Mei", "Juni",
            "Juli", "Augustus", "September", "Oktober", "November", "December"
        };

        [ObservableProperty]
        private ObservableCollection<KlantItem> _klanten = new();

        [ObservableProperty]
        private string? _geselecteerdeMaand;

        [ObservableProperty]
        private bool _isExcelGeladen;

        [ObservableProperty]
        private string _statusTekst = "Selecteer eerst een Excel-bestand";

        [ObservableProperty]
        private bool _isFacturenAanHetGenereren = false;

        [ObservableProperty]
        private double _factuurProgress = 0;

        [ObservableProperty]
        private string _factuurProgressTekst = "";

        public MainViewModel(IConfiguration configuration, KlantHelper klantHelper)
        {
            _configuration = configuration;
            _klantHelper = klantHelper;

            // Test en toon jaar/maand logica bij opstarten
            JaarConfiguratie.ControleerOvergangsPeriode();
            Console.WriteLine($"📅 Jaar configuratie: {JaarConfiguratie.ExcelBestandNaam} → {JaarConfiguratie.RealisatieSheetNaam}");
        }

        partial void OnGeselecteerdeMaandChanged(string? value)
        {
            if (!string.IsNullOrEmpty(value) && IsExcelGeladen)
            {
                LaadKlantenVoorMaand(value);
            }
        }

        [RelayCommand]
        private async Task LaadExcelBestand()
        {
            try
            {
                var result = await FilePicker.PickAsync(new PickOptions
                {
                    PickerTitle = "Selecteer Excel bestand",
                    FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                    {
                        { DevicePlatform.iOS, new[] { "com.microsoft.excel.xlsx" } },
                        { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
                        { DevicePlatform.WinUI, new[] { ".xlsx" } },
                        { DevicePlatform.macOS, new[] { "xlsx" } }
                    })
                });

                if (result != null)
                {
                    _huidigeExcelPad = result.FullPath;
                    await LaadKlantenUitExcel(result.FullPath);
                }
            }
            catch (Exception ex)
            {
                await Application.Current!.MainPage!.DisplayAlert("Fout", $"Kon Excel-bestand niet laden: {ex.Message}", "OK");
            }
        }

        private async Task LaadKlantenUitExcel(string pad)
        {
            try
            {
                _klantenPerMaand.Clear();
                _klantenBedragen.Clear();

                // Update KlantHelper met het juiste Excel bestand pad
                _klantHelper.ZetExcelBestandPad(pad);

                using var workbook = new ExcelPackage(new FileInfo(pad));
                var sheet = workbook.Workbook.Worksheets[JaarConfiguratie.RealisatieSheetNaam]; // Dynamisch sheet naam

                // Laad maandnamen
                var maandNamen = new Dictionary<int, string>();
                for (int kol = 2; kol <= 13; kol++)
                {
                    var maandNaam = sheet.Cells[3, kol].Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(maandNaam))
                    {
                        maandNamen[kol] = maandNaam;
                    }
                }

                // AANGEPAST: Dynamische klanten detectie
                var klantRijen = ZoekKlantRijen(sheet);

                // Laad klanten en bedragen - DYNAMISCH in plaats van hardcoded rij 4-31
                foreach (int rij in klantRijen)
                {
                    var klant = sheet.Cells[rij, 1].Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(klant))
                    {
                        foreach (var maandInfo in maandNamen)
                        {
                            int kol = maandInfo.Key;
                            string maand = maandInfo.Value;

                            var celWaarde = sheet.Cells[rij, kol].Value?.ToString() ?? "";
                            if (double.TryParse(celWaarde, out double bedrag) && bedrag > 0)
                            {
                                if (!_klantenPerMaand.ContainsKey(maand))
                                    _klantenPerMaand[maand] = new List<string>();

                                if (!_klantenPerMaand[maand].Contains(klant))
                                    _klantenPerMaand[maand].Add(klant);

                                if (!_klantenBedragen.ContainsKey(maand))
                                    _klantenBedragen[maand] = new Dictionary<string, double>();

                                _klantenBedragen[maand][klant] = bedrag;
                            }
                        }
                    }
                }

                var totaalKlanten = _klantenPerMaand.Values.SelectMany(x => x).Distinct().Count();
                StatusTekst = $"Excel geladen: {totaalKlanten} unieke klanten gevonden";
                IsExcelGeladen = true;

                await Application.Current!.MainPage!.DisplayAlert("Succes",
                    $"Klanten succesvol geladen uit Excel.\nTotaal {totaalKlanten} unieke klanten gevonden.", "OK");
            }
            catch (Exception ex)
            {
                StatusTekst = "Fout bij laden Excel";
                await Application.Current!.MainPage!.DisplayAlert("Fout", $"Fout bij laden Excel: {ex.Message}", "OK");
            }
        }

        /// <summary>
        /// Zoekt dynamisch alle rijen met klanten/projecten in het Excel bestand
        /// </summary>
        private List<int> ZoekKlantRijen(ExcelWorksheet sheet)
        {
            var klantRijen = new List<int>();

            // Zoek "Factuurgegevens:" header
            int startRij = ZoekFactuurgegevensHeader(sheet);
            if (startRij == -1)
            {
                // Fallback naar oude methode als header niet gevonden
                for (int rij = 4; rij <= 31; rij++)
                {
                    klantRijen.Add(rij);
                }
                return klantRijen;
            }

            // Lees vanaf rij na de header
            int huidigeRij = startRij + 1;

            while (huidigeRij <= 200) // Veiligheidsklep
            {
                var celWaarde = sheet.Cells[huidigeRij, 1].Value?.ToString() ?? "";

                // Stop condities
                if (string.IsNullOrEmpty(celWaarde))
                {
                    // Check of er 3 lege rijen op rij zijn (einde lijst)
                    if (IsConsecutieveLegeCellen(sheet, huidigeRij, 1, 3))
                    {
                        break;
                    }
                }
                else if (IsEindeVanLijst(celWaarde))
                {
                    break; // Stop bij BTW, totaal, etc.
                }
                else
                {
                    // Geldige klant/project regel
                    klantRijen.Add(huidigeRij);
                }

                huidigeRij++;
            }

            return klantRijen;
        }

        /// <summary>
        /// Zoekt de rij waar "Factuurgegevens:" staat
        /// </summary>
        private int ZoekFactuurgegevensHeader(ExcelWorksheet sheet)
        {
            for (int rij = 1; rij <= 20; rij++)
            {
                var celWaarde = sheet.Cells[rij, 1].Value?.ToString() ?? "";
                if (celWaarde.ToLower().Contains("factuurgegevens"))
                {
                    return rij;
                }
            }
            return -1; // Niet gevonden
        }

        /// <summary>
        /// Controleert of we aan het einde van de klanten lijst zijn
        /// </summary>
        private bool IsEindeVanLijst(string waarde)
        {
            var eindeMarkers = new[] { "btw", "totaal", "subtotaal", "kosten", "budget" };
            return eindeMarkers.Any(marker => waarde.ToLower().Contains(marker));
        }

        /// <summary>
        /// Controleert of er meerdere lege cellen na elkaar komen
        /// </summary>
        private bool IsConsecutieveLegeCellen(ExcelWorksheet sheet, int startRij, int kolom, int aantalCellen)
        {
            for (int i = 0; i < aantalCellen; i++)
            {
                var celWaarde = sheet.Cells[startRij + i, kolom].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(celWaarde))
                {
                    return false;
                }
            }
            return true;
        }

        private void LaadKlantenVoorMaand(string maand)
        {
            Klanten.Clear();

            if (_klantenPerMaand.TryGetValue(maand, out var klanten))
            {
                foreach (var klant in klanten.Distinct().OrderBy(k => k))
                {
                    var bedrag = _klantenBedragen[maand].GetValueOrDefault(klant, 0);
                    Klanten.Add(new KlantItem
                    {
                        Naam = klant,
                        Bedrag = bedrag,
                        IsGeselecteerd = false
                    });
                }
            }

            StatusTekst = $"{Klanten.Count} klanten voor {maand}";
        }

        [RelayCommand]
        private void SelecteerAlleKlanten()
        {
            foreach (var klant in Klanten)
            {
                klant.IsGeselecteerd = true;
            }
        }

        [RelayCommand]
        private void DeselecteerAlleKlanten()
        {
            foreach (var klant in Klanten)
            {
                klant.IsGeselecteerd = false;
            }
        }

        [RelayCommand]
        private async Task GenereerFacturen()
        {
            if (string.IsNullOrEmpty(GeselecteerdeMaand))
            {
                await Application.Current!.MainPage!.DisplayAlert("Waarschuwing", "Selecteer eerst een maand.", "OK");
                return;
            }

            var geselecteerdeKlanten = Klanten.Where(k => k.IsGeselecteerd).ToList();
            if (geselecteerdeKlanten.Count == 0)
            {
                await Application.Current!.MainPage!.DisplayAlert("Waarschuwing", "Selecteer ten minste één klant.", "OK");
                return;
            }

            try
            {
                // Dynamisch pad ophalen uit instellingen
                string outputMap = Preferences.Get("FactuurOpslagPad", "");

                if (string.IsNullOrWhiteSpace(outputMap) || !Directory.Exists(outputMap))
                {
                    bool naarInstellingen = await Application.Current!.MainPage!.DisplayAlert(
                        "Opslagpad ontbreekt",
                        "Het opslagpad voor facturen is niet ingesteld of bestaat niet.\nWil je dit nu instellen?",
                        "Ja", "Nee");

                    if (naarInstellingen)
                    {
                        await Shell.Current.GoToAsync(nameof(InstellingenPagina));
                    }
                    return;
                }

                // START PROGRESS TRACKING OP UI THREAD
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    IsFacturenAanHetGenereren = true;
                    FactuurProgress = 0;
                    FactuurProgressTekst = $"Voorbereiden van {geselecteerdeKlanten.Count} facturen...";
                });

                await GenereerFacturenAsync(geselecteerdeKlanten, GeselecteerdeMaand, outputMap);

                // EINDE PROGRESS TRACKING OP UI THREAD
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    IsFacturenAanHetGenereren = false;
                    FactuurProgress = 100;
                    FactuurProgressTekst = "Facturen succesvol gegenereerd!";
                });

                await Application.Current!.MainPage!.DisplayAlert("Succes",
                    $"Facturen zijn opgeslagen in:\n{outputMap}", "OK");

                // Reset progress na dialog OP UI THREAD
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    FactuurProgressTekst = "";
                });
            }
            catch (Exception ex)
            {
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    IsFacturenAanHetGenereren = false;
                    FactuurProgressTekst = "Fout bij genereren facturen";
                });

                await Application.Current!.MainPage!.DisplayAlert("Fout", $"Fout bij genereren facturen: {ex.Message}", "OK");

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    FactuurProgressTekst = "";
                });
            }
        }

        private async Task GenereerFacturenAsync(List<KlantItem> klanten, string maand, string outputMap)
        {
            int aantalGegenereerd = 0;
            var fouten = new List<string>();

            foreach (var klant in klanten)
            {
                try
                {
                    // BEREKEN DYNAMISCH BEDRAG: Uren × Uurtarief
                    double totaalUren = PdfHelper.BerekenTotaalUren(klant.Naam, maand, _klantHelper);
                    double uurtarief = HaalUurtariefOp(klant.Naam);
                    double berekendBedrag = totaalUren * uurtarief;

                    Console.WriteLine($"💰 {klant.Naam}: {totaalUren} uren × €{uurtarief} = €{berekendBedrag:F2}");

                    string factuurnummer = FactuurnummerHelper.GenereerVolgendFactuurnummer(outputMap, maand);
                    string factuurMap = FactuurnummerHelper.BepaalFactuurMap(outputMap, maand);
                    string bestandsNaam = $"{factuurnummer}_{klant.Naam.Replace(" ", "_")}.pdf";
                    string volledigPad = Path.Combine(factuurMap, bestandsNaam);

                    // Gebruik berekend bedrag in plaats van klant.Bedrag
                    await GenereerPDFFactuurMetTweeBladzijdenAsync(klant.Naam, maand, berekendBedrag, volledigPad, factuurnummer);
                    aantalGegenereerd++;
                }
                catch (Exception ex)
                {
                    fouten.Add($"{klant.Naam}: {ex.Message}");
                }
            }

            string bericht = $"Facturen gegenereerd: {aantalGegenereerd}";
            if (fouten.Count > 0)
            {
                bericht += $"\nFouten: {fouten.Count}\n" + string.Join("\n", fouten.Take(5));
                if (fouten.Count > 5)
                    bericht += $"\n... en {fouten.Count - 5} meer";
            }

            await Application.Current!.MainPage!.DisplayAlert("Resultaat", bericht, "OK");
        }

        /// <summary>
        /// Haalt het uurtarief op voor een klant
        /// </summary>
        private double HaalUurtariefOp(string klantNaam)
        {
            try
            {
                var klantNAW = _klantHelper.HaalKlantNAWGegevensOp(klantNaam);
                return klantNAW?.PrijsPerUur ?? 107.5; // Default als niet gevonden
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Kon uurtarief niet ophalen voor {klantNaam}: {ex.Message}");
                return 107.5; // Default
            }
        }

        private async Task GenereerPDFFactuurMetTweeBladzijdenAsync(string klantNaam, string maand, double bedrag, string bestandsPad, string factuurnummer, string klantEmail = null)
        {
            using var stream = new FileStream(bestandsPad, FileMode.Create);
            using var writer = new PdfWriter(stream);
            using var pdf = new PdfDocument(writer);
            using var document = new Document(pdf, PageSize.A4);
            document.SetMargins(50, 50, 50, 50);

            // GEBRUIK DYNAMISCHE DATUM LOGICA
            var eersteVanMaand = JaarConfiguratie.BepaalEersteVanMaand(maand);

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
                if (_configuration != null)
                {
                    using var mollieHelper = new MollieApiHelper(_configuration);
                    var (qrCode, molliePaymentId) = await mollieHelper.CreeerPaymentEnQrCodeAsync(
                        (decimal)bedrag * 1.21m, // Inclusief BTW
                        factuurnummer,
                        klantEmail,
                        klantNaam);

                    qrCodeBytes = qrCode;
                    paymentId = molliePaymentId;
                }
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
        }

        [RelayCommand]
        private async Task OpenInstellingen()
        {
            await Shell.Current.GoToAsync(nameof(InstellingenPagina));
        }
    }
}