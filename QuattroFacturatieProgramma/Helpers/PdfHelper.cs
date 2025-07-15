using System.Reflection;
using static QuattroFacturatieProgramma.Helpers.KlantHelper;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Geom;
using iText.IO.Image;
using iText.Layout.Borders;
using iText.Layout.Renderer;
using iText.Kernel.Pdf.Canvas.Draw;
using SystemPath = System.IO.Path;
using VerticalAlignment = iText.Layout.Properties.VerticalAlignment;
using Border = iText.Layout.Borders.Border;
using Cell = iText.Layout.Element.Cell;
using TextAlignment = iText.Layout.Properties.TextAlignment;
using HorizontalAlignment = iText.Layout.Properties.HorizontalAlignment;
using Image = iText.Layout.Element.Image;

namespace QuattroFacturatieProgramma.Helpers
{
    public static class PdfHelper
    {
        /// <summary>
        /// Berekent totaal uren voor een klant en maand - CENTRALE METHODE
        /// </summary>
        public static double BerekenTotaalUren(string klantNaam, string maand, KlantHelper klantHelper)
        {
            if (klantHelper == null)
                return 0.0;

            try
            {
                var urenRegels = klantHelper.HaalUrenverantwoordingOp(klantNaam, maand);
                if (urenRegels == null || urenRegels.Count == 0)
                    return 0.0;

                double totaalUren = 0;
                foreach (var regel in urenRegels)
                {
                    double uren = ParseUrenWaarde(regel.Uren);
                    totaalUren += uren;
                    Console.WriteLine($"🔢 Uren regel: '{regel.Uren}' → {uren} (totaal: {totaalUren})");
                }

                Console.WriteLine($"✅ Totaal uren berekend voor {klantNaam} in {maand}: {totaalUren}");
                return totaalUren;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Kon totaal uren niet berekenen: {ex.Message}");
                return 0.0;
            }
        }

        public static async Task MaakFactuurPaginaAsync(Document document, string klantNaam, string maand, double bedrag,
            string factuurnummer, DateTime eersteVanMaand, PdfFont titelFont, PdfFont bedrijfsFont,
            PdfFont headerFont, PdfFont normalFont, PdfFont kleinFont, PdfFont accentFont,
            byte[] qrCodeBytes, string paymentId, byte[] logoBytes = null, KlantHelper klantHelper = null,
            double totaalUren = 0)
        {
            // ===== HEADER SECTIE =====
            document.Add(new LineSeparator(new SolidLine(2f))
                .SetStrokeColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(10));

            // Haal klant-specifieke gegevens op
            KlantNAWGegevens klantNAW = null;
            double uurtarief = 107.50;

            if (klantHelper != null)
            {
                try
                {
                    klantNAW = klantHelper.HaalKlantNAWGegevensOp(klantNaam);
                    if (klantNAW != null)
                    {
                        uurtarief = klantNAW.PrijsPerUur;
                    }
                    Console.WriteLine($"💰 Klant: {klantNaam} | Uurtarief: €{uurtarief} | Totaal uren {maand}: {totaalUren}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Kon klant-specifieke gegevens niet ophalen: {ex.Message}");
                }
            }

            // ===== BEDRIJFSGEGEVENS SECTIE - VEEL PLATTER =====
            var bedrijfsTabel = new Table(UnitValue.CreatePercentArray(new float[] { 50f, 50f }))
                .SetWidth(UnitValue.CreatePercentValue(100))
                .SetMarginBottom(15);

            // Klant adres cell - LINKS MET GRIJS VAK
            var klantCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetBackgroundColor(new DeviceRgb(248, 248, 248))
                .SetPadding(10)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            klantCell.Add(new Paragraph("FACTUURADRES")
                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                .SetFontSize(9)
                .SetFontColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(5));

            if (klantNAW != null)
            {
                klantCell.Add(new Paragraph(klantNAW.Naam)
                    .SetFont(headerFont)
                    .SetFontSize(10)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph(klantNAW.Adres)
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph(klantNAW.Straat)
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph($"{klantNAW.Postcode} {klantNAW.Stad}")
                    .SetFont(normalFont)
                    .SetFontSize(9));
            }
            else
            {
                klantCell.Add(new Paragraph(klantNaam)
                    .SetFont(headerFont)
                    .SetFontSize(10)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph("T.a.v. de heer G.P.J. Houben")
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph("Willinkhof 3")
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetMarginBottom(1));
                klantCell.Add(new Paragraph("6006 RG Weert")
                    .SetFont(normalFont)
                    .SetFontSize(9));
            }

            bedrijfsTabel.AddCell(klantCell);

            // Bedrijfs info cell - RECHTS - VEEL PLATTER
            var bedrijfsCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(10)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            // Logo of QUATTRO styling - kleiner
            if (logoBytes != null && logoBytes.Length > 0)
            {
                try
                {
                    var logo = new Image(ImageDataFactory.Create(logoBytes))
                        .SetWidth(100)
                        .SetHeight(30)
                        .SetMarginBottom(3);
                    bedrijfsCell.Add(logo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Kon logo niet laden: {ex.Message}");
                    bedrijfsCell.Add(CreateQuattroStyledParagraph(bedrijfsFont).SetMarginBottom(3));
                }
            }
            else
            {
                bedrijfsCell.Add(CreateQuattroStyledParagraph(bedrijfsFont).SetMarginBottom(3));
            }

            // Alles veel dichter op elkaar
            bedrijfsCell.Add(new Paragraph("BOUW & VASTGOED ADVIES BV")
                .SetFont(headerFont)
                .SetFontSize(10)
                .SetMarginBottom(2));

            bedrijfsCell.Add(new Paragraph("Willinkhof 3")
                .SetFont(normalFont)
                .SetFontSize(9)
                .SetMarginBottom(0));
            bedrijfsCell.Add(new Paragraph("6006 RG Weert")
                .SetFont(normalFont)
                .SetFontSize(9)
                .SetMarginBottom(3));

            bedrijfsCell.Add(new Paragraph("www.quattrobouwenenvastgoedadvies.nl")
                .SetFont(kleinFont)
                .SetFontSize(8)
                .SetMarginBottom(0));
            bedrijfsCell.Add(new Paragraph("info@quattrobouwenenvastgoedadvies.nl")
                .SetFont(kleinFont)
                .SetFontSize(8)
                .SetMarginBottom(2));

            bedrijfsCell.Add(new Paragraph("KvK: 75108542")
                .SetFont(kleinFont)
                .SetFontSize(8)
                .SetMarginBottom(0));
            bedrijfsCell.Add(new Paragraph("BTW: NL860145438B01")
                .SetFont(kleinFont)
                .SetFontSize(8));

            bedrijfsTabel.AddCell(bedrijfsCell);
            document.Add(bedrijfsTabel);

            // ===== FACTUUR TITEL SECTIE =====
            document.Add(new Paragraph("FACTUUR")
                .SetFont(titelFont)
                .SetFontSize(18)
                .SetFontColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(3));

            document.Add(new LineSeparator(new SolidLine(1f))
                .SetStrokeColor(ColorConstants.DARK_GRAY)
                .SetWidth(UnitValue.CreatePercentValue(25))
                .SetMarginBottom(15));

            // ===== FACTUUR INFO SECTIE =====
            var factuurInfoTabel = new Table(UnitValue.CreatePercentArray(new float[] { 33f, 33f, 34f }))
                .SetWidth(UnitValue.CreatePercentValue(100))
                .SetMarginBottom(20);

            string[] labels = { "FACTUURDATUM", "FACTUURNUMMER", "VERVALDATUM" };
            string[] waarden = {
                eersteVanMaand.ToString("dd-MM-yyyy"),
                factuurnummer,
                eersteVanMaand.AddDays(14).ToString("dd-MM-yyyy")
            };

            // Headers
            foreach (var label in labels)
            {
                factuurInfoTabel.AddCell(new Cell()
                    .SetBorder(Border.NO_BORDER)
                    .SetBackgroundColor(new DeviceRgb(245, 245, 245))
                    .SetPadding(12)
                    .Add(new Paragraph(label)
                        .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                        .SetFontSize(10)
                        .SetFontColor(ColorConstants.DARK_GRAY)));
            }

            // Waarden
            foreach (var value in waarden)
            {
                factuurInfoTabel.AddCell(new Cell()
                    .SetBorder(Border.NO_BORDER)
                    .SetPadding(12)
                    .Add(new Paragraph(value)
                        .SetFont(normalFont)
                        .SetFontSize(11)));
            }

            document.Add(factuurInfoTabel);

            // ===== FACTUURREGELS SECTIE =====
            var table = new Table(UnitValue.CreatePercentArray(new float[] { 45f, 15f, 20f, 20f }))
                .SetWidth(UnitValue.CreatePercentValue(100))
                .SetMarginBottom(15);

            var headerColor = new DeviceRgb(50, 50, 50);
            table.AddHeaderCell(CreateHeaderCell("OMSCHRIJVING", accentFont, headerColor));
            table.AddHeaderCell(CreateHeaderCell("AANTAL", accentFont, headerColor, TextAlignment.CENTER));
            table.AddHeaderCell(CreateHeaderCell("TARIEF", accentFont, headerColor, TextAlignment.RIGHT));
            table.AddHeaderCell(CreateHeaderCell("BEDRAG", accentFont, headerColor, TextAlignment.RIGHT));

            // Factuurregels
            table.AddCell(CreateDataCell($"Advieswerkzaamheden {maand} {JaarConfiguratie.BepaalJaarVoorMaand(maand)}", normalFont));
            table.AddCell(CreateDataCell(FormateerUrenWaarde(totaalUren), normalFont, TextAlignment.CENTER));
            table.AddCell(CreateDataCell($"€ {uurtarief:N2}", normalFont, TextAlignment.RIGHT));
            table.AddCell(CreateDataCell($"€ {bedrag:N2}", normalFont, TextAlignment.RIGHT));

            document.Add(table);

            // ===== TOTALEN SECTIE =====
            double btw = bedrag * 0.21;
            double totaal = bedrag + btw;

            var totaalTabel = new Table(2)
                .SetWidth(UnitValue.CreatePercentValue(40))
                .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                .SetMarginBottom(25);

            totaalTabel.AddCell(CreateTotalCell("Subtotaal", normalFont, TextAlignment.LEFT));
            totaalTabel.AddCell(CreateTotalCell($"€ {bedrag:N2}", normalFont, TextAlignment.RIGHT));
            totaalTabel.AddCell(CreateTotalCell("BTW 21%", normalFont, TextAlignment.LEFT));
            totaalTabel.AddCell(CreateTotalCell($"€ {btw:N2}", normalFont, TextAlignment.RIGHT));

            // Totaal regel met meer emphasis
            totaalTabel.AddCell(new Cell()
                .SetBorderTop(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                .SetBorderBottom(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetPadding(15)
                .SetBackgroundColor(new DeviceRgb(240, 240, 240))
                .Add(new Paragraph("TOTAAL")
                    .SetFont(headerFont)
                    .SetFontSize(12)
                    .SetFontColor(ColorConstants.DARK_GRAY)));

            totaalTabel.AddCell(new Cell()
                .SetBorderTop(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                .SetBorderBottom(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetPadding(15)
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetBackgroundColor(new DeviceRgb(240, 240, 240))
                .Add(new Paragraph($"€ {totaal:N2}")
                    .SetFont(headerFont)
                    .SetFontSize(12)
                    .SetFontColor(ColorConstants.DARK_GRAY)));

            document.Add(totaalTabel);

            // ===== BETALINGSVOORWAARDEN SECTIE =====
            document.Add(new Paragraph("BETALINGSVOORWAARDEN")
                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                .SetFontSize(12)
                .SetFontColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(7));

            document.Add(new LineSeparator(new SolidLine(1f))
                .SetStrokeColor(ColorConstants.LIGHT_GRAY)
                .SetMarginBottom(7));

            // Betalings info met QR-code
            var betalingsTabel = new Table(UnitValue.CreatePercentArray(new float[] { 65f, 35f }))
                .SetWidth(UnitValue.CreatePercentValue(100));

            var betalingsInfo = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(2)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            betalingsInfo.Add(new Paragraph("Gelieve het totaalbedrag binnen 14 dagen over te maken:")
                .SetFont(normalFont)
                .SetFontSize(10));
            // .SetMarginBottom(2));

            betalingsInfo.Add(new Paragraph("IBAN: NL30 RABO 0347 6704 07")
                .SetFont(headerFont)
                .SetFontSize(10));
            //  .SetMarginBottom(2));

            // Styled Quattro tekst
            var quattroParagraph = new Paragraph()
                .SetFontSize(10);
               // .SetMarginBottom(2);
            quattroParagraph.Add(new Text("T.n.v. ").SetFont(normalFont));
            AddQuattroStyledPhrase(quattroParagraph, normalFont);
            betalingsInfo.Add(quattroParagraph);

            betalingsInfo.Add(new Paragraph($"O.v.v. Factuurnummer {factuurnummer}")
                .SetFont(normalFont)
                .SetFontSize(10));
            // .SetMarginBottom(2));

            betalingsInfo.Add(new Paragraph("Of scan de QR-code om direct te betalen →")
                .SetFont(kleinFont)
                .SetFontSize(9));
                //.SetMarginBottom(2));

            if (!string.IsNullOrWhiteSpace(paymentId))
                betalingsInfo.Add(new Paragraph($"Payment ID: {paymentId}")
                    .SetFont(kleinFont)
                    .SetFontSize(9));

            betalingsTabel.AddCell(betalingsInfo);

            // QR-code
            var qrCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(2)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            if (qrCodeBytes != null && qrCodeBytes.Length > 0)
            {
                try
                {
                    var qrImage = new Image(ImageDataFactory.Create(qrCodeBytes))
                        .SetWidth(100)
                        .SetHeight(100);
                    qrCell.Add(qrImage);
                    qrCell.Add(new Paragraph("Scan om te betalen")
                        .SetFont(kleinFont)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetMarginTop(2));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fout bij toevoegen QR-code: {ex.Message}");
                    qrCell.Add(new Paragraph("QR-code niet beschikbaar")
                        .SetFont(kleinFont));
                }
            }
            else
            {
                qrCell.Add(new Paragraph("QR-code niet beschikbaar")
                    .SetFont(kleinFont));
            }

            betalingsTabel.AddCell(qrCell);
            document.Add(betalingsTabel);
        }

        public static void MaakUrenverantwoordingPagina(Document document, string maand, string klantNaam,
            PdfFont headerFont, PdfFont normalFont, PdfFont kleinFont, PdfFont accentFont,
            KlantHelper klantHelper = null, double totaalUren = 0)
        {
            // ===== HEADER SECTIE =====
            document.Add(new LineSeparator(new SolidLine(2f))
                .SetStrokeColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(15));

            // ===== TITEL SECTIE =====
            document.Add(new Paragraph("URENVERANTWOORDING")
                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                .SetFontSize(18)
                .SetFontColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(10));

            document.Add(new LineSeparator(new SolidLine(1f))
                .SetStrokeColor(ColorConstants.DARK_GRAY)
                .SetWidth(UnitValue.CreatePercentValue(35))
                .SetMarginBottom(10));

            // ===== INFO SECTIE =====
            document.Add(new Paragraph($"Periode: {maand} {JaarConfiguratie.BepaalJaarVoorMaand(maand)}")
                .SetFont(normalFont)
                .SetFontSize(11)
                .SetMarginTop(0)
                .SetMarginBottom(0));

            document.Add(new Paragraph($"Klant: {klantNaam}")
                .SetFont(normalFont)
                .SetFontSize(11)
                .SetMarginTop(0)
                .SetMarginBottom(10));

            // ===== UREN TABEL SECTIE =====
            var urenTabel = new Table(UnitValue.CreatePercentArray(new float[] { 12f, 48f, 15f, 25f }))
                .SetWidth(UnitValue.CreatePercentValue(100))
                .SetMarginBottom(0);

            var headerColor = new DeviceRgb(50, 50, 50);
            urenTabel.AddHeaderCell(CreateHeaderCell("DATUM", accentFont, headerColor));
            urenTabel.AddHeaderCell(CreateHeaderCell("WERKZAAMHEDEN", accentFont, headerColor));
            urenTabel.AddHeaderCell(CreateHeaderCell("UREN", accentFont, headerColor, TextAlignment.CENTER));
            urenTabel.AddHeaderCell(CreateHeaderCell("OPMERKINGEN", accentFont, headerColor));

            // Probeer klant-specifieke urenverantwoording op te halen
            List<UrenRegel> urenRegels = null;
            if (klantHelper != null)
            {
                try
                {
                    urenRegels = klantHelper.HaalUrenverantwoordingOp(klantNaam, maand);
                    Console.WriteLine($"✅ {urenRegels?.Count ?? 0} uren regels gevonden voor {klantNaam} in {maand}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Kon urenverantwoording niet ophalen: {ex.Message}");
                }
            }

            if (urenRegels != null && urenRegels.Count > 0)
            {
                foreach (var regel in urenRegels)
                {
                    urenTabel.AddCell(CreateDataCell(regel.Datum, normalFont));
                    urenTabel.AddCell(CreateDataCell(regel.Werkzaamheden, normalFont));
                    urenTabel.AddCell(CreateDataCell(regel.Uren, normalFont, TextAlignment.CENTER));
                    string opmerkingen = regel.Opmerkingen == "Conform opdracht" ? "" : regel.Opmerkingen;
                    urenTabel.AddCell(CreateDataCell(opmerkingen, kleinFont));
                }

                document.Add(urenTabel);

                // ===== TOTAAL UREN SECTIE =====
                var totaalTabel = new Table(2)
                    .SetWidth(UnitValue.CreatePercentValue(40))
                    .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                    .SetMarginBottom(15);

                totaalTabel.AddCell(new Cell()
                    .SetBorderTop(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER)
                    .SetPadding(10)
                    .SetBackgroundColor(new DeviceRgb(245, 245, 245))
                    .Add(new Paragraph("TOTAAL UREN")
                        .SetFont(headerFont)
                        .SetFontSize(12)));

                totaalTabel.AddCell(new Cell()
                    .SetBorderTop(new SolidBorder(ColorConstants.DARK_GRAY, 2))
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER)
                    .SetPadding(10)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(245, 245, 245))
                    .Add(new Paragraph(FormateerUrenWaarde(totaalUren))
                        .SetFont(headerFont)
                        .SetFontSize(12)));

                document.Add(totaalTabel);
                Console.WriteLine($"✅ Totaal uren gebruikt in PDF: {FormateerUrenWaarde(totaalUren)}");
            }
            else
            {
                // Fallback naar standaard regel
                Console.WriteLine($"⚠️ Geen uren gevonden voor {klantNaam} in {maand}, gebruik fallback");

                urenTabel.AddCell(CreateDataCell("0-MAAND", normalFont));
                urenTabel.AddCell(CreateDataCell("DEZE CEL IS LEEG", normalFont));
                urenTabel.AddCell(CreateDataCell("0,0", normalFont, TextAlignment.CENTER));
                urenTabel.AddCell(CreateDataCell("", kleinFont));

                document.Add(urenTabel);

                var totaalTabel = new Table(2)
                    .SetWidth(UnitValue.CreatePercentValue(30))
                    .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                    .SetMarginBottom(30);

                totaalTabel.AddCell(CreateTotalCell("TOTAAL UREN", headerFont, TextAlignment.LEFT));
                totaalTabel.AddCell(CreateTotalCell(totaalUren > 0 ? FormateerUrenWaarde(totaalUren) : "1,0", headerFont, TextAlignment.CENTER));

                document.Add(totaalTabel);
            }

            // ===== OPMERKINGEN SECTIE =====
            document.Add(new Paragraph("AANVULLENDE OPMERKINGEN")
                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                .SetFontSize(12)
                .SetFontColor(ColorConstants.DARK_GRAY)
                .SetMarginBottom(10));

            document.Add(new LineSeparator(new SolidLine(1f))
                .SetStrokeColor(ColorConstants.LIGHT_GRAY)
                .SetMarginBottom(15));

            document.Add(new Paragraph("Alle werkzaamheden zijn uitgevoerd conform de opdrachtbevestiging en geldende voorwaarden.")
                .SetFont(normalFont)
                .SetFontSize(10));
        }

        /// <summary>
        /// Maakt een styled phrase voor Quattro tekst in betalingsinfo
        /// </summary>
        private static void AddQuattroStyledPhrase(Paragraph paragraph, PdfFont baseFont)
        {
            var orangeColor = new DeviceRgb(218, 119, 47);
            var blueColor = new DeviceRgb(70, 89, 155);

            paragraph.Add(new Text("Q").SetFont(baseFont).SetFontColor(orangeColor));
            paragraph.Add(new Text("uattro").SetFont(baseFont).SetFontColor(blueColor));
            paragraph.Add(new Text(" Bouw & Vastgoed Advies BV").SetFont(baseFont));
        }

        /// <summary>
        /// Parseert uren waarde uit Excel naar double
        /// </summary>
        private static double ParseUrenWaarde(string urenText)
        {
            if (string.IsNullOrWhiteSpace(urenText))
                return 0.0;

            string normalizedText = urenText.Trim().Replace(",", ".");

            if (double.TryParse(normalizedText, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }

            if (double.TryParse(urenText.Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CreateSpecificCulture("nl-NL"), out result))
            {
                return result;
            }

            if (double.TryParse(urenText.Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), out result))
            {
                return result;
            }

            var cleanText = System.Text.RegularExpressions.Regex.Replace(urenText, @"[^\d,.]", "");
            if (!string.IsNullOrEmpty(cleanText))
            {
                cleanText = cleanText.Replace(",", ".");
                if (double.TryParse(cleanText, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out result))
                {
                    return result;
                }
            }

            Console.WriteLine($"⚠️ Kon uren waarde niet parsen: '{urenText}'");
            return 0.0;
        }

        /// <summary>
        /// Formatteert uren waarde voor PDF output
        /// </summary>
        private static string FormateerUrenWaarde(double uren)
        {
            return uren.ToString("0.0", System.Globalization.CultureInfo.CreateSpecificCulture("nl-NL"));
        }

        /// <summary>
        /// Maakt een styled paragraph voor QUATTRO met oranje Q en blauwe uattro
        /// </summary>
        private static Paragraph CreateQuattroStyledParagraph(PdfFont baseFont)
        {
            var orangeColor = new DeviceRgb(218, 119, 47);
            var blueColor = new DeviceRgb(70, 89, 155);

            var paragraph = new Paragraph()
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(11);

            paragraph.Add(new Text("Q").SetFont(baseFont).SetFontColor(orangeColor));
            paragraph.Add(new Text("UATTRO").SetFont(baseFont).SetFontColor(blueColor));

            return paragraph;
        }

        // ===== CELL CREATION METHODS =====
        private static Cell CreateHeaderCell(string text, PdfFont font, DeviceRgb bgColor, TextAlignment alignment = TextAlignment.LEFT)
        {
            return new Cell()
                .SetBackgroundColor(bgColor)
                .SetPadding(12)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(alignment)
                .SetHeight(25)
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(11)
                    .SetFontColor(ColorConstants.WHITE)
                    .SetMargin(0));
        }

        private static Cell CreateDataCell(string text, PdfFont font, TextAlignment alignment = TextAlignment.LEFT)
        {
            return new Cell()
                .SetPadding(10)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderBottom(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetBackgroundColor(new DeviceRgb(250, 250, 250))
                .SetTextAlignment(alignment)
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(10)
                    .SetMargin(0));
        }

        private static Cell CreateTotalCell(string text, PdfFont font, TextAlignment alignment = TextAlignment.RIGHT)
        {
            return new Cell()
                .SetBorderTop(new SolidBorder(ColorConstants.LIGHT_GRAY, 1))
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetPadding(12)
                .SetTextAlignment(alignment)
                .SetBackgroundColor(new DeviceRgb(248, 248, 248))
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(10)
                    .SetMargin(0));
        }
    }
}