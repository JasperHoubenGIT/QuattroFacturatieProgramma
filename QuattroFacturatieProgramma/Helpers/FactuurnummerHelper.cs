using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace QuattroFacturatieProgramma.Helpers
{
    internal class FactuurnummerHelper
    {
        public static string GenereerVolgendFactuurnummer(string basisPad, string maandNaam)
        {
            string jaar = DateTime.Now.Year.ToString();
            string maandNummer = MaandNaarNummer(maandNaam);
            string mapNaam = $"{jaar}_{maandNummer}_{maandNaam}_Uitgaand";
            string volledigeMap = Path.Combine(basisPad, $"{jaar} Uitgaande facturen", mapNaam);

            if (!Directory.Exists(volledigeMap))
                Directory.CreateDirectory(volledigeMap);

            // Kijk in alle maandmappen van het jaar om het hoogste factuurnummer te vinden
            string jaarMap = Path.Combine(basisPad, $"{jaar} Uitgaande facturen");
            var alleBestanden = new List<string>();

            if (Directory.Exists(jaarMap))
            {
                var maandMappen = Directory.GetDirectories(jaarMap)
                    .Where(map => Path.GetFileName(map).StartsWith($"{jaar}_"));
                foreach (var maandMap in maandMappen)
                {
                    alleBestanden.AddRange(Directory.GetFiles(maandMap));
                }
            }

            // Aangepaste regex voor het nieuwe formaat: "factuur 2025_110_..."
            var regex = new Regex($@"^factuur {jaar}_(\d{{3}})");
            int hoogste = alleBestanden
                .Select(f => Path.GetFileNameWithoutExtension(f))
                .Select(name => regex.Match(name))
                .Where(m => m.Success)
                .Select(m => int.Parse(m.Groups[1].Value))
                .DefaultIfEmpty(0)
                .Max();

            int volgend = hoogste + 1;
            return $"factuur {jaar}_{volgend:000}";
        }

        // NIEUWE METHOD: Check of factuurnummer al bestaat
        public static bool FactuurnummerBestaat(string basisPad, string factuurnummer)
        {
            string jaar = DateTime.Now.Year.ToString();
            string jaarMap = Path.Combine(basisPad, $"{jaar} Uitgaande facturen");

            if (!Directory.Exists(jaarMap))
                return false;

            var maandMappen = Directory.GetDirectories(jaarMap)
                .Where(map => Path.GetFileName(map).StartsWith($"{jaar}_"));

            foreach (var maandMap in maandMappen)
            {
                var bestanden = Directory.GetFiles(maandMap);
                var bestaand = bestanden.Any(f =>
                    Path.GetFileNameWithoutExtension(f).StartsWith(factuurnummer));

                if (bestaand)
                    return true;
            }

            return false;
        }

        // NIEUWE METHOD: Genereer factuurnummer met opgegeven nummer
        public static string GenereerHandmatigFactuurnummer(int nummer)
        {
            string jaar = DateTime.Now.Year.ToString();
            return $"factuur {jaar}_{nummer:000}";
        }

        // NIEUWE METHOD: Krijg alle bestaande factuurnummers
        public static List<int> GetBestaandeFactuurnummers(string basisPad)
        {
            string jaar = DateTime.Now.Year.ToString();
            string jaarMap = Path.Combine(basisPad, $"{jaar} Uitgaande facturen");
            var nummers = new List<int>();

            if (!Directory.Exists(jaarMap))
                return nummers;

            var maandMappen = Directory.GetDirectories(jaarMap)
                .Where(map => Path.GetFileName(map).StartsWith($"{jaar}_"));

            var regex = new Regex($@"^factuur {jaar}_(\d{{3}})");

            foreach (var maandMap in maandMappen)
            {
                var bestanden = Directory.GetFiles(maandMap);
                var mapNummers = bestanden
                    .Select(f => Path.GetFileNameWithoutExtension(f))
                    .Select(name => regex.Match(name))
                    .Where(m => m.Success)
                    .Select(m => int.Parse(m.Groups[1].Value));

                nummers.AddRange(mapNummers);
            }

            return nummers.Distinct().OrderBy(n => n).ToList();
        }

        public static string BepaalFactuurMap(string basisPad, string maandNaam)
        {
            string jaar = DateTime.Now.Year.ToString();
            string maandNummer = MaandNaarNummer(maandNaam);
            return Path.Combine(basisPad, $"{jaar} Uitgaande facturen", $"{jaar}_{maandNummer}_{maandNaam}_Uitgaand");
        }

        private static string MaandNaarNummer(string maandNaam)
        {
            var maanden = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Januari", "01" }, { "Februari", "02" }, { "Maart", "03" },
                { "April", "04" }, { "Mei", "05" }, { "Juni", "06" },
                { "Juli", "07" }, { "Augustus", "08" }, { "September", "09" },
                { "Oktober", "10" }, { "November", "11" }, { "December", "12" }
            };
            return maanden.TryGetValue(maandNaam, out var nummer) ? nummer : "00";
        }
    }
}