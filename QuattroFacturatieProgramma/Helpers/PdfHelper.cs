using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using iTextFont = iTextSharp.text.Font;
using iTextElement = iTextSharp.text.Element;
using iTextImage = iTextSharp.text.Image;
using System.Reflection;
using Element = iTextSharp.text.Element;
using static QuattroFacturatieProgramma.Helpers.KlantHelper;

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
            string factuurnummer, DateTime eersteVanMaand, iTextFont titelFont, iTextFont bedrijfsFont,
            iTextFont headerFont, iTextFont normalFont, iTextFont kleinFont, iTextFont accentFont,
            byte[] qrCodeBytes, string paymentId, byte[] logoBytes = null, KlantHelper klantHelper = null,
            double totaalUren = 0)
        {
            // Headerlijn
            var headerLijn = new LineSeparator(2f, 100f, BaseColor.DARK_GRAY, iTextElement.ALIGN_CENTER, -2);
            document.Add(headerLijn);
            document.Add(new Paragraph(" "));

            // Haal klant-specifieke gegevens op
            KlantNAWGegevens klantNAW = null;
            double uurtarief = 00.00; // Default

            if (klantHelper != null)
            {
                try
                {
                    klantNAW = klantHelper.HaalKlantNAWGegevensOp(klantNaam);

                    // Gebruik klant-specifiek uurtarief
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

            // Bedrijfsgegevens tabel
            var bedrijfsTabel = new PdfPTable(2) { WidthPercentage = 100 };
            bedrijfsTabel.SetWidths(new float[] { 55f, 45f });

            // Klant adres cell - GEBRUIK KLANT-SPECIFIEKE NAW GEGEVENS
            var klantCell = new PdfPCell
            {
                Border = Rectangle.NO_BORDER,
                BackgroundColor = new BaseColor(248, 248, 248),
                Padding = 15,
                VerticalAlignment = iTextElement.ALIGN_TOP
            };

            klantCell.AddElement(new Paragraph("FACTUURADRES", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.DARK_GRAY)));
            klantCell.AddElement(new Paragraph(" ", kleinFont));

            // Gebruik klant-specifieke NAW gegevens of fallback
            if (klantNAW != null)
            {
                klantCell.AddElement(new Paragraph(klantNAW.Naam, headerFont));
                klantCell.AddElement(new Paragraph(klantNAW.Adres, normalFont));
                klantCell.AddElement(new Paragraph(klantNAW.Straat, normalFont));
                klantCell.AddElement(new Paragraph($"{klantNAW.Postcode} {klantNAW.Stad}", normalFont));
            }
            else
            {
                // Fallback naar standaard adres
                klantCell.AddElement(new Paragraph(klantNaam, headerFont));
                klantCell.AddElement(new Paragraph("T.a.v. de heer G.P.J. Houben", normalFont));
                klantCell.AddElement(new Paragraph("Willinkhof 3", normalFont));
                klantCell.AddElement(new Paragraph("6006 RG Weert", normalFont));
            }

            bedrijfsTabel.AddCell(klantCell);

            // Bedrijfs info cell
            var bedrijfsCell = new PdfPCell
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = iTextElement.ALIGN_RIGHT,
                Padding = 15,
                VerticalAlignment = iTextElement.ALIGN_TOP
            };

            // Logo toevoegen als beschikbaar
            if (logoBytes != null && logoBytes.Length > 0)
            {
                try
                {
                    var logo = iTextImage.GetInstance(logoBytes);
                    logo.ScaleToFit(120f, 40f);
                    logo.Alignment = iTextElement.ALIGN_LEFT;
                    bedrijfsCell.AddElement(logo);
                    bedrijfsCell.AddElement(new Paragraph(" ", kleinFont));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Kon logo niet laden: {ex.Message}");
                    // Fallback naar styled tekst
                    bedrijfsCell.AddElement(CreateQuattroStyledParagraph(bedrijfsFont));
                }
            }
            else
            {
                // Styled QUATTRO tekst als fallback
                bedrijfsCell.AddElement(CreateQuattroStyledParagraph(bedrijfsFont));
            }

            bedrijfsCell.AddElement(new Paragraph("BOUW & VASTGOED ADVIES BV", normalFont));
            bedrijfsCell.AddElement(new Paragraph("Willinkhof 3", normalFont));
            bedrijfsCell.AddElement(new Paragraph("6006 RG Weert", normalFont));
            bedrijfsCell.AddElement(new Paragraph(" ", kleinFont));
            bedrijfsCell.AddElement(new Paragraph("www.quattrobouwenenvastgoedadvies.nl", kleinFont));
            bedrijfsCell.AddElement(new Paragraph("info@quattrobouwenenvastgoedadvies.nl", kleinFont));
            bedrijfsCell.AddElement(new Paragraph("KvK: 75108542", kleinFont));
            bedrijfsCell.AddElement(new Paragraph("BTW: NL860145438B01", kleinFont));
            bedrijfsTabel.AddCell(bedrijfsCell);

            document.Add(bedrijfsTabel);
            document.Add(new Paragraph(" "));

            // Factuur titel
            document.Add(new Paragraph("FACTUUR", titelFont) { SpacingAfter = 5f });
            document.Add(new LineSeparator(1f, 30f, BaseColor.DARK_GRAY, iTextElement.ALIGN_LEFT, -2));
            document.Add(new Paragraph(" "));

            // Factuur info tabel (zonder PERIODE)
            var factuurInfoTabel = new PdfPTable(3) { WidthPercentage = 100 };
            factuurInfoTabel.SetWidths(new float[] { 33f, 33f, 34f });

            string[] labels = { "FACTUURDATUM", "FACTUURNUMMER", "VERVALDATUM" };
            string[] waarden = {
                eersteVanMaand.ToString("dd-MM-yyyy"),
                factuurnummer,
                eersteVanMaand.AddDays(14).ToString("dd-MM-yyyy")
            };

            // Headers
            foreach (var label in labels)
            {
                factuurInfoTabel.AddCell(new PdfPCell(new Phrase(label, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.DARK_GRAY)))
                {
                    Border = Rectangle.NO_BORDER,
                    BackgroundColor = new BaseColor(248, 248, 248),
                    Padding = 8
                });
            }

            // Waarden
            foreach (var value in waarden)
            {
                factuurInfoTabel.AddCell(new PdfPCell(new Phrase(value, normalFont))
                {
                    Border = Rectangle.NO_BORDER,
                    Padding = 8
                });
            }

            document.Add(factuurInfoTabel);
            document.Add(new Paragraph(" "));

            // Factuurregels tabel
            var table = new PdfPTable(4) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 40f, 20f, 20f, 20f });

            var headerColor = new BaseColor(60, 60, 60);
            table.AddCell(CreateHeaderCell("OMSCHRIJVING", accentFont, headerColor));
            table.AddCell(CreateHeaderCell("AANTAL", accentFont, headerColor, iTextElement.ALIGN_CENTER));
            table.AddCell(CreateHeaderCell("TARIEF", accentFont, headerColor, iTextElement.ALIGN_RIGHT));
            table.AddCell(CreateHeaderCell("BEDRAG", accentFont, headerColor, iTextElement.ALIGN_RIGHT));

            // Factuurregels - MET KLANT-SPECIFIEKE GEGEVENS
            table.AddCell(CreateDataCell($"Advieswerkzaamheden {maand} {JaarConfiguratie.BepaalJaarVoorMaand(maand)}", normalFont));
            table.AddCell(CreateDataCell(FormateerUrenWaarde(totaalUren), normalFont, iTextElement.ALIGN_CENTER)); // TOTAAL UREN
            table.AddCell(CreateDataCell($"€ {uurtarief:N2}", normalFont, iTextElement.ALIGN_RIGHT)); // KLANT-SPECIFIEK TARIEF
            table.AddCell(CreateDataCell($"€ {bedrag:N2}", normalFont, iTextElement.ALIGN_RIGHT));

            document.Add(table);
            document.Add(new Paragraph(" "));

            // Totalen berekenen
            double btw = bedrag * 0.21;
            double totaal = bedrag + btw;

            // Totalen tabel
            var totaalTabel = new PdfPTable(2) { WidthPercentage = 45, HorizontalAlignment = iTextElement.ALIGN_RIGHT };
            totaalTabel.AddCell(CreateTotalCell("Subtotaal", normalFont));
            totaalTabel.AddCell(CreateTotalCell($"€ {bedrag:N2}", normalFont, iTextElement.ALIGN_RIGHT));
            totaalTabel.AddCell(CreateTotalCell("BTW 21%", normalFont));
            totaalTabel.AddCell(CreateTotalCell($"€ {btw:N2}", normalFont, iTextElement.ALIGN_RIGHT));
            totaalTabel.AddCell(CreateTotalCell("TOTAAL", headerFont));
            totaalTabel.AddCell(CreateTotalCell($"€ {totaal:N2}", headerFont, iTextElement.ALIGN_RIGHT));

            document.Add(totaalTabel);
            document.Add(new Paragraph(" "));

            // Betalingsvoorwaarden
            document.Add(new Paragraph("BETALINGSVOORWAARDEN", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11, BaseColor.DARK_GRAY))
            {
                SpacingBefore = 20f
            });
            document.Add(new LineSeparator(1f, 100f, BaseColor.LIGHT_GRAY, iTextElement.ALIGN_LEFT, -2));
            document.Add(new Paragraph(" "));

            // Betalings info met QR-code
            var betalingsTabel = new PdfPTable(2) { WidthPercentage = 100 };
            betalingsTabel.SetWidths(new float[] { 65f, 35f });

            // Betalingsinfo
            var betalingsInfo = new PdfPCell
            {
                Border = Rectangle.NO_BORDER,
                Padding = 5,
                VerticalAlignment = iTextElement.ALIGN_TOP
            };
            betalingsInfo.AddElement(new Paragraph("Gelieve het totaalbedrag binnen 14 dagen over te maken:", normalFont));
            betalingsInfo.AddElement(new Paragraph("IBAN: NL30 RABO 0347 6704 07", headerFont));

            // Styled Quattro tekst gebruiken
            var quattroPhrase = CreateQuattroStyledPhrase(normalFont);
            var quattroParagraph = new Paragraph();
            quattroParagraph.Add(new Chunk("T.n.v. ", normalFont));
            quattroParagraph.Add(quattroPhrase);
            betalingsInfo.AddElement(quattroParagraph);
            betalingsInfo.AddElement(new Paragraph($"O.v.v. Factuurnummer {factuurnummer}", normalFont));
            betalingsInfo.AddElement(new Paragraph("Of scan de QR-code om direct te betalen →", kleinFont));

            if (!string.IsNullOrWhiteSpace(paymentId))
                betalingsInfo.AddElement(new Paragraph($"Payment ID: {paymentId}", kleinFont));

            betalingsTabel.AddCell(betalingsInfo);

            // QR-code
            var qrCell = new PdfPCell
            {
                Border = Rectangle.NO_BORDER,
                Padding = 10,
                HorizontalAlignment = iTextElement.ALIGN_CENTER,
                VerticalAlignment = iTextElement.ALIGN_MIDDLE
            };

            if (qrCodeBytes != null && qrCodeBytes.Length > 0)
            {
                try
                {
                    var qrImage = iTextImage.GetInstance(qrCodeBytes);
                    qrImage.ScaleToFit(100f, 100f);
                    qrImage.Alignment = iTextElement.ALIGN_CENTER;
                    qrCell.AddElement(qrImage);
                    qrCell.AddElement(new Paragraph("Scan om te betalen", kleinFont)
                    {
                        Alignment = iTextElement.ALIGN_CENTER,
                        SpacingBefore = 5f
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fout bij toevoegen QR-code: {ex.Message}");
                    qrCell.AddElement(new Paragraph("QR-code niet beschikbaar", kleinFont));
                }
            }
            else
            {
                qrCell.AddElement(new Paragraph("QR-code niet beschikbaar", kleinFont));
            }

            betalingsTabel.AddCell(qrCell);
            document.Add(betalingsTabel);
        }

        public static void MaakUrenverantwoordingPagina(Document document, string maand, string klantNaam,
            iTextFont headerFont, iTextFont normalFont, iTextFont kleinFont, iTextFont accentFont,
            KlantHelper klantHelper = null, double totaalUren = 0)
        {
            // Header lijn
            var headerLijn = new LineSeparator(2f, 100f, BaseColor.DARK_GRAY, iTextElement.ALIGN_CENTER, -2);
            document.Add(headerLijn);
            document.Add(new Paragraph(" "));

            // Titel
            document.Add(new Paragraph("URENVERANTWOORDING", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18, BaseColor.DARK_GRAY)));
            document.Add(new LineSeparator(1f, 40f, BaseColor.DARK_GRAY, iTextElement.ALIGN_LEFT, -2));
            document.Add(new Paragraph(" "));

            document.Add(new Paragraph($"Periode: {maand} {JaarConfiguratie.BepaalJaarVoorMaand(maand)}", normalFont));
            document.Add(new Paragraph($"Klant: {klantNaam}", normalFont));
            document.Add(new Paragraph(" "));

            // Uren tabel - KOLOMBREEDTE AANGEPAST
            var urenTabel = new PdfPTable(4) { WidthPercentage = 100 };
            urenTabel.SetWidths(new float[] { 12f, 48f, 15f, 25f }); // UREN kolom breder

            var headerColor = new BaseColor(60, 60, 60);
            urenTabel.AddCell(CreateHeaderCell("DATUM", accentFont, headerColor));
            urenTabel.AddCell(CreateHeaderCell("WERKZAAMHEDEN", accentFont, headerColor));
            urenTabel.AddCell(CreateHeaderCell("UREN", accentFont, headerColor, iTextElement.ALIGN_CENTER));
            urenTabel.AddCell(CreateHeaderCell("OPMERKINGEN", accentFont, headerColor));

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

            // GEBRUIK KLANT-SPECIFIEKE EN MAAND-GEFILTERDE URENVERANTWOORDING
            if (urenRegels != null && urenRegels.Count > 0)
            {
                foreach (var regel in urenRegels)
                {
                    urenTabel.AddCell(CreateDataCell(regel.Datum, normalFont));
                    urenTabel.AddCell(CreateDataCell(regel.Werkzaamheden, normalFont));
                    urenTabel.AddCell(CreateDataCell(regel.Uren, normalFont, iTextElement.ALIGN_CENTER));
                    string opmerkingen = regel.Opmerkingen == "Conform opdracht" ? "" : regel.Opmerkingen;
                    urenTabel.AddCell(CreateDataCell(opmerkingen, kleinFont));
                }

                document.Add(urenTabel);
                document.Add(new Paragraph(" "));

                // Totaal uren - GEBRUIK MEEGEGEVEN WAARDE
                var totaalTabel = new PdfPTable(2) { WidthPercentage = 35, HorizontalAlignment = iTextElement.ALIGN_RIGHT };
                totaalTabel.AddCell(CreateTotalCell("TOTAAL UREN", headerFont));
                totaalTabel.AddCell(CreateTotalCell(FormateerUrenWaarde(totaalUren), headerFont, iTextElement.ALIGN_CENTER));

                Console.WriteLine($"✅ Totaal uren gebruikt in PDF: {FormateerUrenWaarde(totaalUren)}");

                document.Add(totaalTabel);
            }
            else
            {
                // Fallback naar standaard regel
                Console.WriteLine($"⚠️ Geen uren gevonden voor {klantNaam} in {maand}, gebruik fallback");

                urenTabel.AddCell(CreateDataCell("20-mei", normalFont));
                urenTabel.AddCell(CreateDataCell("Reactie naar architect met onderbouwing adviesrapport", normalFont));
                urenTabel.AddCell(CreateDataCell("1,0", normalFont, iTextElement.ALIGN_CENTER));
                urenTabel.AddCell(CreateDataCell("", kleinFont)); // Lege opmerking als fallback

                document.Add(urenTabel);
                document.Add(new Paragraph(" "));

                // Totaal uren - GEBRUIK MEEGEGEVEN WAARDE OF FALLBACK
                var totaalTabel = new PdfPTable(2) { WidthPercentage = 35, HorizontalAlignment = iTextElement.ALIGN_RIGHT };
                totaalTabel.AddCell(CreateTotalCell("TOTAAL UREN", headerFont));
                totaalTabel.AddCell(CreateTotalCell(totaalUren > 0 ? FormateerUrenWaarde(totaalUren) : "1,0", headerFont, iTextElement.ALIGN_CENTER));

                document.Add(totaalTabel);
            }

            document.Add(new Paragraph(" "));

            // Opmerkingen
            document.Add(new Paragraph("AANVULLENDE OPMERKINGEN", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11, BaseColor.DARK_GRAY))
            {
                SpacingBefore = 20f
            });
            document.Add(new LineSeparator(1f, 100f, BaseColor.LIGHT_GRAY, iTextElement.ALIGN_LEFT, -2));
            document.Add(new Paragraph(" "));
            document.Add(new Paragraph("Alle werkzaamheden zijn uitgevoerd conform de opdrachtbevestiging en geldende voorwaarden.", normalFont));
        }

        /// <summary>
        /// Parseert uren waarde uit Excel naar double
        /// </summary>
        private static double ParseUrenWaarde(string urenText)
        {
            if (string.IsNullOrWhiteSpace(urenText))
                return 0.0;

            // Vervang komma door punt voor consistent parsing
            string normalizedText = urenText.Trim().Replace(",", ".");

            // Probeer direct te parsen
            if (double.TryParse(normalizedText, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }

            // Probeer met Nederlandse cultuur (komma als decimaal)
            if (double.TryParse(urenText.Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CreateSpecificCulture("nl-NL"), out result))
            {
                return result;
            }

            // Probeer met Amerikaanse cultuur (punt als decimaal)
            if (double.TryParse(urenText.Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), out result))
            {
                return result;
            }

            // Laatste poging: extract alleen cijfers en punt/komma
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
            // Nederlandse formatting: 1,5 in plaats van 1.5
            return uren.ToString("0.0", System.Globalization.CultureInfo.CreateSpecificCulture("nl-NL"));
        }

        /// <summary>
        /// Maakt een styled paragraph voor QUATTRO met oranje Q en blauwe uattro
        /// </summary>
        private static Paragraph CreateQuattroStyledParagraph(iTextFont baseFont)
        {
            var orangeFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, baseFont.Size, new BaseColor(218, 119, 47)); // Oranje #DA772F
            var blueFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, baseFont.Size, new BaseColor(70, 89, 155)); // Blauw #46599B

            var paragraph = new Paragraph();
            paragraph.Alignment = iTextElement.ALIGN_LEFT;

            // Q in oranje
            paragraph.Add(new Chunk("Q", orangeFont));
            // uattro in blauw  
            paragraph.Add(new Chunk("UATTRO", blueFont));

            return paragraph;
        }

        /// <summary>
        /// Maakt een styled chunk voor Quattro tekst in betalingsinfo
        /// </summary>
        private static Phrase CreateQuattroStyledPhrase(iTextFont baseFont)
        {
            var orangeFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, baseFont.Size, new BaseColor(218, 119, 47));
            var blueFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, baseFont.Size, new BaseColor(70, 89, 155));

            var phrase = new Phrase();
            phrase.Add(new Chunk("Q", orangeFont));
            phrase.Add(new Chunk("uattro", blueFont));
            phrase.Add(new Chunk(" Bouw & Vastgoed Advies BV", baseFont));

            return phrase;
        }

        private static PdfPCell CreateHeaderCell(string text, iTextFont font, BaseColor bgColor, int alignment = iTextElement.ALIGN_LEFT)
        {
            return new PdfPCell(new Phrase(text, font))
            {
                BackgroundColor = bgColor,
                Padding = 12,
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = alignment,
                NoWrap = true, // Voorkom text wrapping
                MinimumHeight = 25f // Voldoende hoogte
            };
        }

        private static PdfPCell CreateDataCell(string text, iTextFont font, int alignment = iTextElement.ALIGN_LEFT)
        {
            return new PdfPCell(new Phrase(text, font))
            {
                Padding = 12,
                Border = Rectangle.BOTTOM_BORDER,
                BorderColor = BaseColor.LIGHT_GRAY,
                BackgroundColor = new BaseColor(252, 252, 252),
                HorizontalAlignment = alignment
            };
        }

        private static PdfPCell CreateTotalCell(string text, iTextFont font, int alignment = iTextElement.ALIGN_RIGHT)
        {
            return new PdfPCell(new Phrase(text, font))
            {
                Border = Rectangle.TOP_BORDER,
                BorderColor = BaseColor.DARK_GRAY,
                Padding = 10,
                HorizontalAlignment = alignment,
                BackgroundColor = new BaseColor(248, 248, 248)
            };
        }
    }
}