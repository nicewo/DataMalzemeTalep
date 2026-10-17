using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using DataMalzemeTalep.Models;

namespace DataMalzemeTalep.Services;

public sealed class PdfCatalogImportService
{
    private sealed record PdfLine(double Y, List<Word> Words)
    {
        public string Text => string.Join(" ", Words.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text));
    }

    public List<Material> ReadMaterials(string pdfPath)
    {
        using var document = PdfDocument.Open(pdfPath);
        var results = new List<Material>();

        foreach (var page in document.GetPages())
        {
            var lines = BuildLines(page.GetWords());
            var headerIndex = lines.FindIndex(line =>
                Has(line.Text, "MARKA") && Has(line.Text, "AÇIKLAMA") &&
                Has(line.Text, "MİKTAR") && (Has(line.Text, "ÜRÜN") || Has(line.Text, "KODU")));

            if (headerIndex < 0)
                continue;

            var header = lines[headerIndex];
            var brandX = HeaderCenter(header.Words, w => Has(w.Text, "MARKA"));
            var modelX = HeaderCenter(header.Words, w => Has(w.Text, "ÜRÜN") || Has(w.Text, "KODU"));
            var descriptionX = HeaderCenter(header.Words, w => Has(w.Text, "AÇIKLAMA"));
            var quantityX = HeaderCenter(header.Words, w => Has(w.Text, "MİKTAR"));
            var unitX = HeaderCenter(header.Words, w => Has(w.Text, "BRİM") || Has(w.Text, "BİRİM"));

            if (brandX is null || modelX is null || descriptionX is null || quantityX is null || unitX is null)
                continue;

            var columns = new[] { brandX.Value, modelX.Value, descriptionX.Value, quantityX.Value, unitX.Value };
            for (var i = headerIndex + 1; i < lines.Count; i++)
            {
                var line = lines[i];
                if (line.Y >= header.Y)
                    continue;

                var fields = new string[5];
                foreach (var word in line.Words)
                {
                    var x = (word.BoundingBox.Left + word.BoundingBox.Right) / 2;
                    var nearest = Enumerable.Range(0, columns.Length)
                        .MinBy(index => Math.Abs(columns[index] - x));
                    fields[nearest] = Join(fields[nearest], word.Text);
                }

                var brand = Clean(fields[0]);
                var model = Clean(fields[1]);
                var description = Clean(fields[2]);
                var quantityText = Clean(fields[3]);
                var unit = Clean(fields[4]);

                // Ignore blank template rows, page numbers, and notes outside the item table.
                if (string.IsNullOrWhiteSpace(brand) &&
                    string.IsNullOrWhiteSpace(model) &&
                    string.IsNullOrWhiteSpace(description))
                    continue;

                if (IsFooterOrNote(brand, model, description))
                    break;

                // Product rows normally have a quantity and/or unit. This prevents
                // the company header and free-form notes from becoming catalog items.
                if (!Regex.IsMatch(quantityText, @"\d") && string.IsNullOrWhiteSpace(unit))
                    continue;

                results.Add(new Material
                {
                    Category = InferCategory(description),
                    Brand = brand,
                    Model = model,
                    Specifications = description,
                    Description = description,
                    Unit = NormalizeUnit(unit)
                });
            }
        }

        // Keep the first occurrence from the document; catalog merge logic decides
        // whether it updates an existing model or adds a new one.
        return results;
    }

    private static List<PdfLine> BuildLines(IEnumerable<Word> words)
    {
        var ordered = words
            .Where(w => !string.IsNullOrWhiteSpace(w.Text))
            .OrderByDescending(w => (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2)
            .ThenBy(w => w.BoundingBox.Left)
            .ToList();

        var lines = new List<PdfLine>();
        foreach (var word in ordered)
        {
            var y = (word.BoundingBox.Top + word.BoundingBox.Bottom) / 2;
            var line = lines.FirstOrDefault(candidate => Math.Abs(candidate.Y - y) <= 2.8);
            if (line is null)
                lines.Add(new PdfLine(y, [word]));
            else
                line.Words.Add(word);
        }

        return lines
            .Select(line => new PdfLine(line.Y, line.Words.OrderBy(w => w.BoundingBox.Left).ToList()))
            .OrderByDescending(line => line.Y)
            .ToList();
    }

    private static double? HeaderCenter(List<Word> words, Func<Word, bool> predicate)
    {
        var matches = words.Where(predicate).ToList();
        return matches.Count == 0
            ? null
            : matches.Average(w => (w.BoundingBox.Left + w.BoundingBox.Right) / 2);
    }

    private static string Join(string left, string right) =>
        string.IsNullOrWhiteSpace(left) ? right : left + " " + right;

    private static bool Has(string source, string value) =>
        Normalize(source).Contains(Normalize(value), StringComparison.Ordinal);

    private static string Normalize(string value) =>
        value.ToUpperInvariant().Replace('İ', 'I').Replace('Ş', 'S').Replace('Ğ', 'G')
            .Replace('Ü', 'U').Replace('Ö', 'O').Replace('Ç', 'C');

    private static string Clean(string value) =>
        Regex.Replace(value.Replace("\u00A0", " "), @"\s+", " ").Trim().Trim('|', ':', ';');

    private static bool IsFooterOrNote(string brand, string model, string description)
    {
        var text = Normalize($"{brand} {model} {description}");
        return text.Contains("DATA OTOMASYON") ||
               text.Contains("VERGI DAIRESI") ||
               text.Contains("VERGI NO") ||
               text.Contains("MAI L") ||
               text.Contains("ADRES") ||
               text.Contains("BAKIR BARALAR") ||
               text.Contains("TARAFINDAN BELIRLENECEK");
    }

    private static string NormalizeUnit(string unit)
    {
        var value = Normalize(unit).Trim().TrimEnd('.', ',');
        return value switch
        {
            "ADET" or "AD" => "Adet",
            "METRE" or "MT" or "M" => "Metre",
            "TAKIM" => "Takım",
            "SET" => "Set",
            "KG" or "KILO" => "Kg",
            _ => string.IsNullOrWhiteSpace(unit) ? "Adet" : unit.Trim()
        };
    }

    private static string InferCategory(string description)
    {
        var text = Normalize(description);
        if (ContainsAny(text, "INVERTER", "SURUCU", "VFD", "FREKANS")) return "Sürücüler ve Motor Kontrol";
        if (ContainsAny(text, "PLC", "CPU", "DIJITAL GIRIS", "ANALOG GIRIS", "HMI")) return "PLC ve Otomasyon";
        if (ContainsAny(text, "KABLO", "TTR", "KUMANDA KABLOSU")) return "Kablolar";
        if (ContainsAny(text, "KLEMENS", "KONTAK", "ROLE", "RÖLE", "RÖLE")) return "Röle ve Arayüz";
        if (ContainsAny(text, "SENSOR", "SWITCH", "SWIC", "PROXIMITY", "TERMOPLASTIK MAKARALI")) return "Sensörler";
        if (ContainsAny(text, "OTOMAT", "TMŞ", "TMS", "MOTOR KORUMA", "ACIL STOP", "BUTON", "SİGORTA", "SIGORTA")) return "Elektrik Malzemeleri";
        if (ContainsAny(text, "RAKOR", "HORTUM", "PNOMAT", "PNÜMAT")) return "Pnömatik";
        if (ContainsAny(text, "KABLO KANALI", "BUAT", "PANO", "BARA")) return "Pano ve Mekanik";
        return "İçe Aktarılan";
    }

    private static bool ContainsAny(string source, params string[] terms) =>
        terms.Any(term => source.Contains(Normalize(term), StringComparison.Ordinal));
}
