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

        // ========== NIEUWE DATUM CONFIGURATIE PROPERTIES ==========
        [ObservableProperty]
        private bool _gebruikAutomatischeDatum = true;

        [ObservableProperty]
        private DateTime _handmaligeDatum = DateTime.Today;

        [ObservableProperty]
        private DateTime _factuurdatum = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

        // Property change handlers voor datum logica
        partial void OnGebruikAutomatischeDatumChanged(bool value)
        {
            Console.WriteLine($"🔄 GebruikAutomatischeDatum gewijzigd naar: {value}");
            Preferences.Set("GebruikAutomatischeDatum", value);
            BerekenFactuurdatum();
        }

        partial void OnHandmaligeDatumChanged(DateTime value)
        {
            Console.WriteLine($"🗓️ HandmaligeDatum gewijzigd naar: {value:dd-MM-yyyy}");
            Preferences.Set("HandmaligeDatum", value.ToString());
            if (!GebruikAutomatischeDatum)
            {
                BerekenFactuurdatum();
            }
        }

        private void BerekenFactuurdatum()
        {
            if (GebruikAutomatischeDatum)
            {
                Factuurdatum = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                Console.WriteLine($"📅 Automatische factuurdatum: {Factuurdatum:dd-MM-yyyy}");
            }
            else
            {
                Factuurdatum = HandmaligeDatum;
                Console.WriteLine($"📅 Handmatige factuurdatum: {Factuurdatum:dd-MM-yyyy}");
            }
        }
        // ========== EINDE NIEUWE DATUM CONFIGURATIE ==========

        [RelayCommand]
        private async Task SelecteerMaand()
        {
            try
            {
                Console.WriteLine("🎯 SelecteerMaand aangeroepen");

                if (Maanden == null || Maanden.Count == 0)
                {
                    Console.WriteLine("❌ Geen maanden beschikbaar");
                    return;
                }

                // Toon action sheet met maanden
                var result = await Application.Current!.MainPage!.DisplayActionSheet(
                    "Selecteer een maand",
                    "Annuleren",
                    null,
                    Maanden.ToArray());

                Console.WriteLine($"🎯 Gebruiker selecteerde: {result}");

                // Controleer of er iets geselecteerd is (niet "Annuleren" of null)
                if (!string.IsNullOrEmpty(result) && result != "Annuleren" && Maanden.Contains(result))
                {
                    Console.WriteLine($"✅ Geldige maand geselecteerd: {result}");
                    GeselecteerdeMaand = result;
                }
                else
                {
                    Console.WriteLine($"❌ Ongeldige selectie: {result}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout in SelecteerMaand: {ex.Message}");
                await Application.Current!.MainPage!.DisplayAlert("Fout", $"Kon maand niet selecteren: {ex.Message}", "OK");
            }
        }

        [RelayCommand]
        private void SelecteerMaandDirect(string maand)
        {
            try
            {
                Console.WriteLine($"🎯 SelecteerMaandDirect aangeroepen met: {maand}");

                if (string.IsNullOrEmpty(maand))
                {
                    Console.WriteLine("❌ Lege maand parameter");
                    return;
                }

                if (Maanden?.Contains(maand) == true)
                {
                    Console.WriteLine($"✅ Geldige maand geselecteerd: {maand}");
                    GeselecteerdeMaand = maand;
                }
                else
                {
                    Console.WriteLine($"❌ Ongeldige maand: {maand}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout in SelecteerMaandDirect: {ex.Message}");
            }
        }

        [RelayCommand]
        private async Task TestMaanden()
        {
            Console.WriteLine("🧪 === TEST MAANDEN ===");
            Console.WriteLine($"📅 Maanden count: {Maanden?.Count ?? 0}");

            if (Maanden != null)
            {
                foreach (var maand in Maanden)
                {
                    Console.WriteLine($"   - {maand}");
                }
            }

            Console.WriteLine($"🎯 GeselecteerdeMaand: '{GeselecteerdeMaand}'");
            Console.WriteLine($"📊 IsExcelGeladen: {IsExcelGeladen}");
            Console.WriteLine($"🔄 IsFacturenAanHetGenereren: {IsFacturenAanHetGenereren}");

            await Application.Current!.MainPage!.DisplayAlert("Test",
                $"Maanden: {Maanden?.Count ?? 0}\nGeselecteerd: {GeselecteerdeMaand}", "OK");
        }

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

            // Zorg dat Maanden collectie gevuld is
            Maanden = new ObservableCollection<string>
            {
                "Januari", "Februari", "Maart", "April", "Mei", "Juni",
                "Juli", "Augustus", "September", "Oktober", "November", "December"
            };

            Console.WriteLine($"📅 Maanden collectie geïnitialiseerd: {Maanden.Count} maanden");

            // Test en toon jaar/maand logica bij opstarten
            JaarConfiguratie.ControleerOvergangsPeriode();
            Console.WriteLine($"📅 Jaar configuratie: {JaarConfiguratie.ExcelBestandNaam} → {JaarConfiguratie.RealisatieSheetNaam}");

            // ========== DATUM VOORKEUREN LADEN ==========
            GebruikAutomatischeDatum = Preferences.Get("GebruikAutomatischeDatum", true);

            if (DateTime.TryParse(Preferences.Get("HandmaligeDatum", DateTime.Today.ToString()), out DateTime opgeslagenDatum))
            {
                HandmaligeDatum = opgeslagenDatum;
            }
            else
            {
                HandmaligeDatum = DateTime.Today;
            }

            // Bereken initiële factuurdatum
            BerekenFactuurdatum();

            Console.WriteLine($"🔧 Datum voorkeuren geladen - Automatisch: {GebruikAutomatischeDatum}, Handmalig: {HandmaligeDatum:dd-MM-yyyy}, Factuurdatum: {Factuurdatum:dd-MM-yyyy}");
        }

        partial void OnGeselecteerdeMaandChanged(string? value)
        {
            Console.WriteLine($"🗓️ OnGeselecteerdeMaandChanged aangeroepen: '{value}'");
            Console.WriteLine($"📊 IsExcelGeladen: {IsExcelGeladen}");
            Console.WriteLine($"📅 Beschikbare maanden: {string.Join(", ", Maanden)}");

            if (!string.IsNullOrEmpty(value) && IsExcelGeladen)
            {
                Console.WriteLine($"✅ Laad klanten voor maand: {value}");
                LaadKlantenVoorMaand(value);
            }
            else
            {
                Console.WriteLine($"❌ Kan klanten niet laden - Value: '{value}', IsExcelGeladen: {IsExcelGeladen}");
            }
        }

        [RelayCommand]
        private async Task TestMollieApi()
        {
            try
            {
                await Application.Current!.MainPage!.DisplayAlert("Info", "Test Mollie API verbinding...", "OK");

                using var mollieHelper = new MollieApiHelper(_configuration);

                // Test verbinding
                var verbindingOk = await mollieHelper.TestVerbindingAsync();

                if (!verbindingOk)
                {
                    await Application.Current!.MainPage!.DisplayAlert("Fout", "Mollie API verbinding gefaald", "OK");
                    return;
                }

                await Application.Current!.MainPage!.DisplayAlert("Succes", "Mollie API verbinding werkt!", "OK");

                // Test payment creation
                var (qrCode, paymentLinkId) = await mollieHelper.CreeerPaymentEnQrCodeAsync(
                    10.00m, // Test bedrag
                    "TEST-001", // Test factuurnummer
                    "test@example.com",
                    "Test Klant");

                await Application.Current!.MainPage!.DisplayAlert("Test Resultaat",
                    $"Payment ID: {paymentLinkId}\n" +
                    $"QR Code: {(qrCode != null ? $"{qrCode.Length} bytes" : "NULL")}", "OK");
            }
            catch (Exception ex)
            {
                await Application.Current!.MainPage!.DisplayAlert("Fout",
                    $"Mollie API test gefaald:\n{ex.Message}", "OK");
            }
        }

        [RelayCommand]
        private async Task LaadExcelBestand()
        {
            try
            {
                Console.WriteLine("🔄 FilePicker wordt gestart...");

                var result = await FilePicker.PickAsync(new PickOptions
                {
                    PickerTitle = "Selecteer Excel bestand",
                    FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                    {
                        { DevicePlatform.iOS, new[] { "com.microsoft.excel.xlsx", "org.openxmlformats.spreadsheetml.sheet" } },
                        { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
                        { DevicePlatform.WinUI, new[] { ".xlsx" } },
                        { DevicePlatform.macOS, new[] { "xlsx", "xls" } },
                        { DevicePlatform.MacCatalyst, new[] { "xlsx", "xls", "public.spreadsheet", "com.microsoft.excel.xlsx", "public.data" } }
                    })
                });

                if (result != null)
                {
                    Console.WriteLine($"📄 Bestand geselecteerd: {result.FileName}");
                    Console.WriteLine($"📍 Pad: {result.FullPath}");

                    // Controleer bestandsextensie
                    var extension = Path.GetExtension(result.FileName).ToLower();
                    if (extension != ".xlsx" && extension != ".xls")
                    {
                        await Application.Current!.MainPage!.DisplayAlert("Fout",
                            "Selecteer een Excel bestand (.xlsx of .xls)", "OK");
                        return;
                    }

                    _huidigeExcelPad = result.FullPath;
                    await LaadKlantenUitExcel(result.FullPath);
                }
                else
                {
                    Console.WriteLine("❌ Geen bestand geselecteerd");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fout bij selecteren bestand: {ex.Message}");
                await Application.Current!.MainPage!.DisplayAlert("Fout", $"Kon Excel-bestand niet selecteren: {ex.Message}", "OK");
            }
        }

        private async Task LaadKlantenUitExcel(string pad)
        {
            try
            {
                _klantenPerMaand.Clear();
                _klantenBedragen.Clear();
                _klantHelper.ZetExcelBestandPad(pad);

                using var workbook = new ExcelPackage(new FileInfo(pad));
                var sheet = workbook.Workbook.Worksheets[JaarConfiguratie.RealisatieSheetNaam];

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

                // Zoek klanten rijen
                var klantRijen = ZoekKlantRijen(sheet);

                // DEBUG: Toon alle klanten uit kolom 1
                var alleKlantenUitKolom1 = new List<string>();
                var klantenMetBedragen = new HashSet<string>();

                foreach (int rij in klantRijen)
                {
                    var klant = sheet.Cells[rij, 1].Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(klant))
                    {
                        alleKlantenUitKolom1.Add($"R{rij}: {klant}");

                        // Check of klant bedragen heeft
                        bool heeftBedrag = false;
                        foreach (var maandInfo in maandNamen)
                        {
                            int kol = maandInfo.Key;
                            string maand = maandInfo.Value;
                            var celWaarde = sheet.Cells[rij, kol].Value?.ToString() ?? "";

                            if (double.TryParse(celWaarde, out double bedrag) && bedrag > 0)
                            {
                                heeftBedrag = true;
                                klantenMetBedragen.Add(klant);

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

                // Toon debug info via alerts
                var klantenZonderBedragen = alleKlantenUitKolom1
                    .Select(k => k.Split(": ")[1])
                    .Where(k => !klantenMetBedragen.Contains(k))
                    .ToList();

                var totaalKlanten = _klantenPerMaand.Values.SelectMany(x => x).Distinct().Count();

                // DEBUG ALERT 1: Overzicht
                await Application.Current!.MainPage!.DisplayAlert("Debug Info 1",
                    $"📋 Klanten uit kolom 1: {alleKlantenUitKolom1.Count}\n" +
                    $"💰 Klanten met bedragen: {klantenMetBedragen.Count}\n" +
                    $"🎯 Unieke klanten (result): {totaalKlanten}\n" +
                    $"❌ Klanten zonder bedragen: {klantenZonderBedragen.Count}", "OK");

                // DEBUG ALERT 2: Klanten lijst (eerste 10)
                var eersteKlanten = string.Join("\n", alleKlantenUitKolom1.Take(10));
                await Application.Current!.MainPage!.DisplayAlert("Debug Info 2 - Eerste 10 klanten",
                    eersteKlanten, "OK");

                // DEBUG ALERT 3: Laatste klanten
                var laatsteKlanten = string.Join("\n", alleKlantenUitKolom1.Skip(10));
                if (!string.IsNullOrEmpty(laatsteKlanten))
                {
                    await Application.Current!.MainPage!.DisplayAlert("Debug Info 3 - Overige klanten",
                        laatsteKlanten, "OK");
                }

                // DEBUG ALERT 4: Klanten zonder bedragen
                if (klantenZonderBedragen.Any())
                {
                    var zonderBedragen = string.Join("\n", klantenZonderBedragen);
                    await Application.Current!.MainPage!.DisplayAlert("Debug Info 4 - Klanten ZONDER bedragen",
                        zonderBedragen, "OK");
                }

                StatusTekst = $"Excel geladen: {totaalKlanten} unieke klanten gevonden";
                IsExcelGeladen = true;

                await Application.Current!.MainPage!.DisplayAlert("Succes",
                    $"Klanten succesvol geladen uit Excel.\n\n" +
                    $"📊 Totaal {totaalKlanten} unieke klanten gevonden\n" +
                    $"📅 Maanden: {string.Join(", ", _klantenPerMaand.Keys)}", "OK");
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
            Console.WriteLine($"🔄 LaadKlantenVoorMaand: {maand}");

            Klanten.Clear();

            if (_klantenPerMaand.TryGetValue(maand, out var klanten))
            {
                Console.WriteLine($"👥 {klanten.Count} klanten gevonden voor {maand}");

                foreach (var klant in klanten.Distinct().OrderBy(k => k))
                {
                    var bedrag = _klantenBedragen[maand].GetValueOrDefault(klant, 0);
                    Console.WriteLine($"   💰 {klant}: €{bedrag:F2}");

                    Klanten.Add(new KlantItem
                    {
                        Naam = klant,
                        Bedrag = bedrag,
                        IsGeselecteerd = false
                    });
                }
            }
            else
            {
                Console.WriteLine($"❌ Geen klanten gevonden voor maand: {maand}");
                Console.WriteLine($"📋 Beschikbare maanden: {string.Join(", ", _klantenPerMaand.Keys)}");
            }

            StatusTekst = $"{Klanten.Count} klanten voor {maand}";
            Console.WriteLine($"✅ Status: {StatusTekst}");
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

            // ========== GEBRUIK NIEUWE FACTUURDATUM LOGICA ==========
            var eersteVanMaand = Factuurdatum; // Gebruik de berekende factuurdatum

            Console.WriteLine($"📅 Factuurdatum voor PDF: {eersteVanMaand:dd-MM-yyyy} (Automatisch: {GebruikAutomatischeDatum})");

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