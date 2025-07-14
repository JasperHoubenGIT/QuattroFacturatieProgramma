using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace QuattroFacturatieProgramma.Helpers
{
    public class KlantHelper
    {
        private string _excelBestand;

        public KlantHelper(string excelBestand)
        {
            _excelBestand = excelBestand;
        }

        /// <summary>
        /// Wijzigt het Excel bestand pad (voor FilePicker scenario)
        /// </summary>
        public void ZetExcelBestandPad(string nieuwPad)
        {
            _excelBestand = nieuwPad;
            Console.WriteLine($"📁 Excel bestand pad gewijzigd naar: {nieuwPad}");
        }

        /// <summary>
        /// Zoekt het tabblad dat matcht met de klantnaam
        /// </summary>
        public string ZoekKlantTabblad(string klantNaam)
        {
            try
            {
                using var workbook = new XLWorkbook(_excelBestand);
                var worksheetNamen = workbook.Worksheets.Select(w => w.Name).ToList();

                // Exacte match eerst
                var exacteMatch = worksheetNamen.FirstOrDefault(naam =>
                    naam.Equals(klantNaam, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(exacteMatch))
                {
                    return exacteMatch;
                }

                // Deel van naam match
                var gedeeltelijkeMatch = worksheetNamen.FirstOrDefault(naam =>
                    naam.ToLower().Contains(klantNaam.ToLower()) ||
                    klantNaam.ToLower().Contains(naam.ToLower()));

                if (!string.IsNullOrEmpty(gedeeltelijkeMatch))
                {
                    return gedeeltelijkeMatch;
                }

                // Intelligente fuzzy match
                var fuzzyMatch = ZoekFuzzyMatch(klantNaam, worksheetNamen);
                return fuzzyMatch;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fout bij zoeken tabblad voor {klantNaam}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Fuzzy matching voor tabblad namen
        /// </summary>
        private string ZoekFuzzyMatch(string klantNaam, List<string> worksheetNamen)
        {
            var klantWoorden = klantNaam.ToLower()
                .Split(' ', '_', '-', '.', ',')
                .Where(w => w.Length > 2)
                .ToList();

            var besteMatch = "";
            int hoogsteScore = 0;

            foreach (var worksheetNaam in worksheetNamen)
            {
                // Skip system sheets
                if (worksheetNaam.ToLower().Contains("realisatie") ||
                    worksheetNaam.ToLower().Contains("prognose") ||
                    worksheetNaam.ToLower().Contains("blad"))
                {
                    continue;
                }

                var worksheetWoorden = worksheetNaam.ToLower()
                    .Split(' ', '_', '-', '.', ',')
                    .Where(w => w.Length > 2)
                    .ToList();

                int score = 0;
                foreach (var klantWoord in klantWoorden)
                {
                    if (worksheetWoorden.Any(w => w.Contains(klantWoord) || klantWoord.Contains(w)))
                    {
                        score++;
                    }
                }

                if (score > hoogsteScore && score > 0)
                {
                    hoogsteScore = score;
                    besteMatch = worksheetNaam;
                }
            }

            return hoogsteScore > 0 ? besteMatch : null;
        }

        /// <summary>
        /// Haalt NAW gegevens op uit klant-specifiek tabblad - EXACT VOOR JOUW STRUCTUUR
        /// </summary>
        public KlantNAWGegevens HaalKlantNAWGegevensOp(string klantNaam)
        {
            var defaultNAW = new KlantNAWGegevens
            {
                Naam = klantNaam,
                Adres = "T.a.v. de heer G.P.J. Houben",
                Straat = "Willinkhof 3",
                Postcode = "6006 RG",
                Stad = "Weert",
                PrijsPerUur = 107.5
            };

            var tabbladNaam = ZoekKlantTabblad(klantNaam);
            if (string.IsNullOrEmpty(tabbladNaam))
            {
                Console.WriteLine($"⚠️ Geen tabblad gevonden voor {klantNaam}, gebruik default NAW");
                return defaultNAW;
            }

            try
            {
                using var workbook = new XLWorkbook(_excelBestand);
                var worksheet = workbook.Worksheet(tabbladNaam);

                // EXACTE IMPLEMENTATIE VOOR JOUW STRUCTUUR
                var nawGegevens = new KlantNAWGegevens
                {
                    // Rij 1: B="Naam Project", C="QHR_Weert_Laarveld 10 woningen"
                    Naam = worksheet.Cell(1, 3).GetString().Trim(),

                    // Rij 2: B="T.a.v.", C="de heer G.P.J. Houben"
                    Adres = worksheet.Cell(2, 3).GetString().Trim(),

                    // Rij 3: B="Straatnaam", C="Willinkhof 3"
                    Straat = worksheet.Cell(3, 3).GetString().Trim(),

                    // Rij 4: B="Postcode", C="6006 RG"
                    Postcode = worksheet.Cell(4, 3).GetString().Trim(),

                    // Rij 5: B="Stad", C="Weert"
                    Stad = worksheet.Cell(5, 3).GetString().Trim(),

                    // Rij 6: B="Prijs per uur", C="107.5" - VERBETERDE UITLEZING
                    PrijsPerUur = ParsePrijsPerUurUitCell(worksheet.Cell(6, 3))
                };

                // Vul ontbrekende velden met defaults
                if (string.IsNullOrEmpty(nawGegevens.Naam))
                    nawGegevens.Naam = klantNaam;
                if (string.IsNullOrEmpty(nawGegevens.Adres))
                    nawGegevens.Adres = defaultNAW.Adres;
                if (string.IsNullOrEmpty(nawGegevens.Straat))
                    nawGegevens.Straat = defaultNAW.Straat;
                if (string.IsNullOrEmpty(nawGegevens.Postcode))
                    nawGegevens.Postcode = defaultNAW.Postcode;
                if (string.IsNullOrEmpty(nawGegevens.Stad))
                    nawGegevens.Stad = defaultNAW.Stad;
                if (nawGegevens.PrijsPerUur <= 0)
                    nawGegevens.PrijsPerUur = defaultNAW.PrijsPerUur;

                Console.WriteLine($"✅ NAW gegevens gevonden voor {nawGegevens.Naam} (tabblad: {tabbladNaam})");
                return nawGegevens;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Fout bij ophalen NAW voor {klantNaam}: {ex.Message}");
                return defaultNAW;
            }
        }

        /// <summary>
        /// Parseert prijs per uur direct uit Excel cel (betere methode)
        /// </summary>
        private double ParsePrijsPerUurUitCell(IXLCell cell)
        {
            try
            {
                // Probeer eerst als getal
                if (cell.TryGetValue(out double getalWaarde))
                {
                    Console.WriteLine($"✅ Prijs als getal gelezen: {getalWaarde}");
                    return getalWaarde;
                }

                // Als dat niet lukt, probeer als tekst
                var tekstWaarde = cell.GetString().Trim();
                Console.WriteLine($"🔍 Prijs als tekst: '{tekstWaarde}'");

                return ParsePrijsPerUur(tekstWaarde);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Fout bij uitlezen prijs cel: {ex.Message}");
                return 107.5;
            }
        }

        /// <summary>
        /// Parseert prijs per uur vanuit Excel cel
        /// </summary>
        private double ParsePrijsPerUur(string prijsText)
        {
            if (string.IsNullOrEmpty(prijsText)) return 107.5;

            // Debug: toon wat er wordt gelezen
            Console.WriteLine($"🔍 ParsePrijsPerUur input: '{prijsText}'");

            // Vervang komma door punt voor decimal parsing
            prijsText = prijsText.Replace(",", ".");

            // Probeer direct te parsen
            if (double.TryParse(prijsText, out double prijs))
            {
                Console.WriteLine($"✅ Prijs geparsed: {prijs}");
                return prijs;
            }

            // Probeer met Nederlandse cultuur
            if (double.TryParse(prijsText.Replace(".", ","), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CreateSpecificCulture("nl-NL"), out prijs))
            {
                Console.WriteLine($"✅ Prijs geparsed (NL): {prijs}");
                return prijs;
            }

            Console.WriteLine($"⚠️ Kon prijs niet parsen: '{prijsText}', gebruik default 107.5");
            return 107.5; // Default
        }

        /// <summary>
        /// Haalt urenverantwoording op uit klant-specifiek tabblad - GEFILTERD OP MAAND
        /// </summary>
        public List<UrenRegel> HaalUrenverantwoordingOp(string klantNaam, string maand)
        {
            var tabbladNaam = ZoekKlantTabblad(klantNaam);
            if (string.IsNullOrEmpty(tabbladNaam))
            {
                Console.WriteLine($"⚠️ Geen tabblad gevonden voor {klantNaam}, gebruik fallback uren");
                return CreateFallbackUrenRegels(klantNaam, maand);
            }

            try
            {
                using var workbook = new XLWorkbook(_excelBestand);
                var worksheet = workbook.Worksheet(tabbladNaam);

                var alleUrenRegels = new List<UrenRegel>();

                // STAP 1: Lees ALLE uren uit het tabblad
                for (int rij = 10; rij <= 200; rij++)
                {
                    var datumCel = worksheet.Cell(rij, 2).GetString().Trim();      // Kolom B
                    var activiteitCel = worksheet.Cell(rij, 3).GetString().Trim(); // Kolom C
                    var urenCel = worksheet.Cell(rij, 4).GetString().Trim();       // Kolom D
                    var opmerkingenCel = worksheet.Cell(rij, 5).GetString().Trim(); // Kolom E - OPMERKINGEN!

                    // Stop als alle cellen leeg zijn
                    if (string.IsNullOrEmpty(datumCel) &&
                        string.IsNullOrEmpty(activiteitCel) &&
                        string.IsNullOrEmpty(urenCel))
                    {
                        // Check volgende paar rijen voor zekerheid
                        bool isEchtEinde = true;
                        for (int checkRij = rij + 1; checkRij <= rij + 3; checkRij++)
                        {
                            if (!string.IsNullOrEmpty(worksheet.Cell(checkRij, 2).GetString()) ||
                                !string.IsNullOrEmpty(worksheet.Cell(checkRij, 3).GetString()) ||
                                !string.IsNullOrEmpty(worksheet.Cell(checkRij, 4).GetString()))
                            {
                                isEchtEinde = false;
                                break;
                            }
                        }

                        if (isEchtEinde)
                        {
                            break;
                        }
                    }

                    // Voeg toe als er activiteit is
                    if (!string.IsNullOrEmpty(activiteitCel))
                    {
                        var werkDatum = ConverteerExcelDatumNaarDateTime(datumCel);

                        alleUrenRegels.Add(new UrenRegel
                        {
                            Datum = werkDatum?.ToString("dd-MMM") ?? DateTime.Now.ToString("dd-MMM"),
                            Werkzaamheden = activiteitCel,
                            Uren = string.IsNullOrEmpty(urenCel) ? "1,0" : urenCel.Replace(".", ","),
                            Opmerkingen = opmerkingenCel, // ECHTE OPMERKINGEN UIT EXCEL!
                            WerkDatum = werkDatum // Voor filtering
                        });

                        // DEBUG: Log wat er gelezen wordt
                        Console.WriteLine($"🔍 Regel {rij}: Datum='{datumCel}', Activiteit='{activiteitCel}', Uren='{urenCel}', Opmerkingen='{opmerkingenCel}'");
                    }
                }

                // STAP 2: Filter op gewenste maand EN JAAR
                var doelMaand = ConverteerMaandNaamNaarNummer(maand);
                var doelJaar = JaarConfiguratie.BepaalJaarVoorMaand(maand); // Gebruik dynamische jaar logica

                var gefilterdUrenRegels = alleUrenRegels.Where(regel =>
                    regel.WerkDatum.HasValue &&
                    regel.WerkDatum.Value.Month == doelMaand &&
                    regel.WerkDatum.Value.Year == doelJaar // Correct jaar op basis van logica
                ).ToList();

                Console.WriteLine($"📊 {alleUrenRegels.Count} totale uren gevonden, {gefilterdUrenRegels.Count} in {maand} {doelJaar} voor {klantNaam}");

                if (gefilterdUrenRegels.Count == 0)
                {
                    Console.WriteLine($"⚠️ Geen uren gevonden voor {maand} {doelJaar} in tabblad {tabbladNaam}, gebruik fallback");
                    return CreateFallbackUrenRegels(klantNaam, maand);
                }

                // Sorteer op datum
                gefilterdUrenRegels = gefilterdUrenRegels.OrderBy(r => r.WerkDatum).ToList();

                Console.WriteLine($"✅ {gefilterdUrenRegels.Count} uren regels gevonden voor {klantNaam} in {maand} {doelJaar} (tabblad: {tabbladNaam})");
                return gefilterdUrenRegels;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Fout bij ophalen uren voor {klantNaam}: {ex.Message}");
                return CreateFallbackUrenRegels(klantNaam, maand);
            }
        }

        /// <summary>
        /// Converteert Excel serial date naar DateTime
        /// </summary>
        private DateTime? ConverteerExcelDatumNaarDateTime(string datumText)
        {
            if (string.IsNullOrEmpty(datumText)) return null;

            // Als het een Excel serial date nummer is (zoals 45678)
            if (int.TryParse(datumText, out int excelSerialDate))
            {
                try
                {
                    // Excel serial date naar DateTime
                    // Excel telt dagen vanaf 1 januari 1900, maar heeft een bug voor leapyear
                    var baseDate = new DateTime(1900, 1, 1);
                    var actualDate = baseDate.AddDays(excelSerialDate - 2); // -2 voor Excel bug correctie
                    return actualDate;
                }
                catch
                {
                    return null;
                }
            }

            // Als het al een datum string is
            if (DateTime.TryParse(datumText, out DateTime datum))
            {
                return datum;
            }

            // Probeer andere datum formaten
            var dateFormats = new[] { "dd-MM-yyyy", "dd/MM/yyyy", "dd-MM-yy", "dd/MM/yy", "dd-MMM-yyyy", "dd-MMM-yy" };
            foreach (var format in dateFormats)
            {
                if (DateTime.TryParseExact(datumText, format, null, System.Globalization.DateTimeStyles.None, out datum))
                {
                    return datum;
                }
            }

            return null;
        }

        /// <summary>
        /// Converteert maandnaam naar maandnummer
        /// </summary>
        private int ConverteerMaandNaamNaarNummer(string maandNaam)
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

            return maandMapping.TryGetValue(maandNaam, out int maandNummer) ? maandNummer : DateTime.Now.Month;
        }

        /// <summary>
        /// Maakt fallback uren regels als tabblad niet gevonden
        /// </summary>
        private List<UrenRegel> CreateFallbackUrenRegels(string klantNaam, string maand)
        {
            return new List<UrenRegel>
            {
                new UrenRegel
                {
                    Datum = DateTime.Now.ToString("dd-MMM"),
                    Werkzaamheden = $"Advieswerkzaamheden voor {klantNaam}",
                    Uren = "1,0",
                    Opmerkingen = "", // Lege opmerking als fallback
                    WerkDatum = DateTime.Now
                }
            };
        }

        // Bestaande methodes voor klanten lijst...
        public List<string> HaalKlantenOp()
        {
            var klanten = new List<string>();

            try
            {
                using var workbook = new XLWorkbook(_excelBestand);
                var worksheet = workbook.Worksheet(JaarConfiguratie.RealisatieSheetNaam); // Dynamisch sheet naam

                int startRij = ZoekFactuurgegevensHeader(worksheet);
                if (startRij == -1)
                {
                    throw new Exception("Factuurgegevens header niet gevonden");
                }

                int huidigeRij = startRij + 1;

                while (true)
                {
                    var celWaarde = worksheet.Cell(huidigeRij, 1).GetString().Trim();

                    if (string.IsNullOrEmpty(celWaarde) ||
                        IsEindeVanLijst(celWaarde) ||
                        IsConsecutieveLegeCel(worksheet, huidigeRij, 1))
                    {
                        break;
                    }

                    klanten.Add(celWaarde);
                    huidigeRij++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Fout bij lezen klanten: {ex.Message}");
            }

            return klanten;
        }

        private int ZoekFactuurgegevensHeader(IXLWorksheet worksheet)
        {
            for (int rij = 1; rij <= 20; rij++)
            {
                var celWaarde = worksheet.Cell(rij, 1).GetString().ToLower();
                if (celWaarde.Contains("factuurgegevens"))
                {
                    return rij;
                }
            }
            return -1;
        }

        private bool IsEindeVanLijst(string waarde)
        {
            var eindeMarkers = new[] { "btw", "totaal", "subtotaal", "kosten", "budget" };
            return eindeMarkers.Any(marker => waarde.ToLower().Contains(marker));
        }

        private bool IsConsecutieveLegeCel(IXLWorksheet worksheet, int startRij, int kolom)
        {
            for (int i = 0; i < 3; i++)
            {
                var celWaarde = worksheet.Cell(startRij + i, kolom).GetString().Trim();
                if (!string.IsNullOrEmpty(celWaarde))
                {
                    return false;
                }
            }
            return true;
        }
    }

    /// <summary>
    /// NAW gegevens model - UITGEBREID VOOR JOUW STRUCTUUR
    /// </summary>
    public class KlantNAWGegevens
    {
        public string Naam { get; set; }
        public string Adres { get; set; }
        public string Straat { get; set; }
        public string Postcode { get; set; }
        public string Stad { get; set; }
        public double PrijsPerUur { get; set; }
    }

    /// <summary>
    /// Urenverantwoording regel model
    /// </summary>
    public class UrenRegel
    {
        public string Datum { get; set; }
        public string Werkzaamheden { get; set; }
        public string Uren { get; set; }
        public string Opmerkingen { get; set; }
        public DateTime? WerkDatum { get; set; } // Voor filtering op maand
    }

    /// <summary>
    /// Klant informatie model
    /// </summary>
    public class KlantInfo
    {
        public string Naam { get; set; }
        public double Bedrag { get; set; }
    }
}