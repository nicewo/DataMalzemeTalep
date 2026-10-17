using ClosedXML.Excel;
using DataMalzemeTalep.Models;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Path = System.IO.Path;

namespace DataMalzemeTalep.Services;

public sealed class DataService
{
    public string Root { get; } = AppContext.BaseDirectory;
    public string CatalogPath => Path.Combine(Root, "catalog.xml");
    public string RequestsPath => Path.Combine(Root, "Listeler");

    public DataService()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(CatalogPath)!);
        Directory.CreateDirectory(RequestsPath);
        if (!File.Exists(CatalogPath))
            CreateCatalog();
    }

    public List<Material> LoadCatalog()
    {
        var materials = ReadCatalog(CatalogPath);

        // An older app version may already have created an empty catalog.xml.
        // In that case, recover the starter materials from the bundled template.
        if (materials.Count == 0)
        {
            var templatePath = Path.Combine(Root, "Data", "catalog.xml");
            if (File.Exists(templatePath))
            {
                var templateMaterials = ReadCatalog(templatePath);
                if (templateMaterials.Count > 0)
                {
                    WriteCatalog(CatalogPath, new MaterialCatalog { Materials = templateMaterials });
                    materials = templateMaterials;
                }
            }
        }

        return materials;
    }

    private static List<Material> ReadCatalog(string path)
    {
        if (!File.Exists(path))
            return [];

        using var stream = File.OpenRead(path);
        return ((MaterialCatalog?)new XmlSerializer(typeof(MaterialCatalog)).Deserialize(stream))?.Materials ?? [];
    }

    private void CreateCatalog()
    {
        var templatePath = Path.Combine(Root, "Data", "catalog.xml");
        var materials = ReadCatalog(templatePath);
        WriteCatalog(CatalogPath, new MaterialCatalog { Materials = materials });
    }

    private static void WriteCatalog(string path, MaterialCatalog catalog)
    {
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        new XmlSerializer(typeof(MaterialCatalog)).Serialize(stream, catalog);
    }

    public void SaveCatalog(List<Material> materials) => WriteCatalog(CatalogPath, new MaterialCatalog { Materials = materials });

    public List<string> GetSavedRequestFiles() =>
        Directory.GetFiles(RequestsPath, "*.xlsx").OrderByDescending(File.GetLastWriteTime).ToList();

    public string SaveRequest(RequestDocument request, string? existingPath = null)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var safeName = new string(request.Name.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray()).Trim();
        if (safeName.Length == 0)
            safeName = "Malzeme_Listesi";

        var path = existingPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            path = Path.Combine(RequestsPath, safeName + ".xlsx");
            if (File.Exists(path))
                throw new IOException("Bu isimde bir liste zaten var.");
        }

        int rowCount = Math.Max(request.Items.Count, 30);
        int lastRow = 12 + rowCount;   // 44 kalemde 56 eder

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sayfa1");
        sheet.Style.Font.FontName = "Calibri";
        sheet.Style.Font.Bold = true;
        sheet.PageSetup.PaperSize = (XLPaperSize)9;
        sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        sheet.PageSetup.Margins.Top = 0.35;
        sheet.PageSetup.Margins.Bottom = 0.35;
        sheet.PageSetup.Margins.Left = 0.25;
        sheet.PageSetup.Margins.Right = 0.25;
        sheet.PageSetup.PrintAreas.Clear();
        sheet.PageSetup.PrintAreas.Add($"A1:F{lastRow}");
        sheet.PageSetup.PagesWide = 1;
        sheet.PageSetup.PagesTall = 0;

        // Keep the original form's compact six-column proportions.
        sheet.Column("A").Width = 3.50;
        sheet.Column("B").Width = 14;
        sheet.Column("C").Width = 22.14;
        sheet.Column("D").Width = 36.86;
        sheet.Column("E").Width = 6.86;
        sheet.Column("F").Width = 8;

        sheet.Range("A2:F2").Merge();
        // Company logo in the upper-left area of the form.
        var logoPath = Path.Combine(Root, "Logo.png");
        if (File.Exists(logoPath))
        {
            sheet.AddPicture(logoPath)
                .MoveTo(sheet.Cell("A2"), 5, 5)
                .WithSize(105, 78);
        }

        sheet.Cell("A2").Value = "MALZEME TALEP FORMU";
        sheet.Cell("A2").Style.Font.FontName = "Times New Roman";
        sheet.Cell("A2").Style.Font.FontSize = 18;
        sheet.Cell("A2").Style.Font.Bold = true;
        sheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Cell("A2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Row(2).Height = 25;

        sheet.Cell("C4").Value = "LİSTE ADI :";
        sheet.Cell("C4").Style.Font.FontName = "Times New Roman";
        sheet.Cell("C4").Style.Font.Bold = true;
        sheet.Cell("C4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        sheet.Cell("D4").Value = request.Name;
        sheet.Cell("D4").Style.Font.FontName = "Times New Roman";
        sheet.Cell("D4").Style.Font.FontSize = 16;
        sheet.Cell("D4").Style.Font.Bold = true;
        sheet.Cell("D4").Style.Font.FontColor = XLColor.Red;
        sheet.Cell("D4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("D4:F4").Merge();
        sheet.Range("D4:F4").Style.Border.OutsideBorder = XLBorderStyleValues.Hair;

        // Company details occupy the full B:D area; keep the date in E:F.
        foreach (var range in new[] { "B6:D6", "B7:D7", "B8:D8", "B9:D9", "B10:D10" })
            sheet.Range(range).Merge();

        sheet.Cell("B6").Value = "DATA OTOMASYON ELEKTRİK SANAYİ VE TİCARET LİMİTED ŞİRKETİ";
        sheet.Cell("B7").Value = "ADRES: ORGANİZE SANAYİ BÖLGESİ MAH. NAFİ GÜRAL CAD. NO: 3/8D\nMERKEZ/ KÜTAHYA";
        sheet.Cell("B8").Value = "VERGİ DAİRESİ: 30 AĞUSTOS";
        sheet.Cell("B9").Value = "VERGİ NO: 2711937105";
        sheet.Cell("B10").Value = "MAİL: satinalma@dataotomasyon.com / muhasebe@dataotomasyon.com";
        foreach (var address in new[] { "B6", "B7", "B8", "B9", "B10" })
        {
            sheet.Cell(address).Style.Font.FontName = "Times New Roman";
            sheet.Cell(address).Style.Font.FontSize = 8;
            sheet.Cell(address).Style.Font.Bold = true;
            sheet.Cell(address).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Cell(address).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Cell(address).Style.Alignment.WrapText = true;
        }
        sheet.Row(6).Height = 18;
        sheet.Row(7).Height = 30;
        sheet.Row(8).Height = 18;
        sheet.Row(9).Height = 18;
        sheet.Row(10).Height = 22;

        sheet.Range("E9:F9").Merge();
        sheet.Cell("E9").Value = "TARİH:";
        sheet.Cell("E9").Style.Fill.BackgroundColor = XLColor.FromArgb(0, 0, 0);
        sheet.Cell("E9").Style.Font.FontColor = XLColor.FromArgb(255, 255, 255);
        sheet.Cell("E9").Style.Font.Bold = true;
        sheet.Cell("E9").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Cell("E9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Range("E10:F10").Merge();
        sheet.Cell("E10").Style.Fill.BackgroundColor = XLColor.FromArgb(0, 0, 0);
        sheet.Cell("E10").Style.Font.FontColor = XLColor.FromArgb(255, 255, 255);
        sheet.Cell("E10").Value = request.CreatedAt;
        sheet.Cell("E10").Style.DateFormat.Format = "dd.MM.yyyy";
        sheet.Cell("E10").Style.Font.Bold = true;
        sheet.Cell("E10").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] headers = ["SIRA NO", "MARKA", "ÜRÜN KODU", "AÇIKLAMA", "MİKTAR", "BİRİM"];
        for (int col = 1; col <= headers.Length; col++)
        {
            var cell = sheet.Cell(12, col);
            cell.Value = headers[col - 1];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 10;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Hair;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Hair;
        }
        sheet.Row(12).Height = 27;

        for (int i = 0; i < rowCount; i++)
        {
            int rowNumber = 13 + i;
            var row = sheet.Row(rowNumber);
            row.Height = 17;
            var item = i < request.Items.Count ? request.Items[i] : null;
            sheet.Cell(rowNumber, 1).Value = i + 1;

            if (item is not null)
            {
                sheet.Cell(rowNumber, 2).Value = item.Brand;
                sheet.Cell(rowNumber, 3).Value = item.Model;
                sheet.Cell(rowNumber, 4).Value = item.Specifications;
                sheet.Cell(rowNumber, 5).Value = item.Quantity;
                sheet.Cell(rowNumber, 6).Value = item.Unit;
            }

            int requiredLines = 1;
            for (int col = 1; col <= 6; col++)
            {
                var cell = sheet.Cell(rowNumber, col);
                cell.Style.Font.FontName = "Times New Roman";
                cell.Style.Font.FontSize = 9;
                cell.Style.Border.LeftBorder = XLBorderStyleValues.Hair;
                cell.Style.Border.RightBorder = XLBorderStyleValues.Hair;
                cell.Style.Border.TopBorder = XLBorderStyleValues.Hair;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Hair;

                // All list data is aligned to the top-left as requested.
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                // Reduce to 9 pt first; if text still exceeds the cell width, wrap to at most 2 lines.
                if (col == 2 || col == 3 || col == 4)
                {
                    var value = cell.GetString();
                    var maxCharsPerLine = col switch
                    {
                        2 => 18, // Brand
                        3 => 30, // Product code / model
                        _ => 52  // Description
                    };
                    int lines = Math.Max(1, (int)Math.Ceiling(value.Length / (double)maxCharsPerLine));
                    lines = Math.Min(2, lines);
                    if (lines > 1)
                    {
                        cell.Style.Alignment.WrapText = true;
                        requiredLines = Math.Max(requiredLines, lines);
                    }
                }
            }

            // Make rows tall enough for wrapped content, without changing single-line row heights.
            row.Height = requiredLines > 1 ? 30 : 17;
        }

        // Show normal Excel worksheet gridlines outside the explicitly bordered form.
        sheet.ShowGridLines = true;
        workbook.SaveAs(path);
        return path;
    }

    public List<RequestItem> LoadRequest(string path)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheet(1);
        var list = new List<RequestItem>();

        // Support the supplied company form as well as older app-generated spreadsheets.
        if (sheet.Cell("A12").GetString().Trim().Equals("SIRA NO", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var row in sheet.RowsUsed().Where(r => r.RowNumber() >= 13))
            {
                var brand = row.Cell(2).GetString().Trim();
                var model = row.Cell(3).GetString().Trim();
                var description = row.Cell(4).GetString().Trim();
                var quantityText = row.Cell(5).GetString().Trim();
                var unit = row.Cell(6).GetString().Trim();
                if (string.IsNullOrWhiteSpace(brand) && string.IsNullOrWhiteSpace(model) &&
                    string.IsNullOrWhiteSpace(description) && string.IsNullOrWhiteSpace(quantityText) &&
                    string.IsNullOrWhiteSpace(unit))
                    continue;

                decimal.TryParse(quantityText, out var quantity);
                var parts = description.Split(" • ", 2, StringSplitOptions.None);
                list.Add(new RequestItem
                {
                    Category = parts.Length > 1 ? parts[0] : "",
                    Brand = brand,
                    Model = model,
                    Specifications = parts.Length > 1 ? parts[1] : description,
                    Quantity = quantity,
                    Unit = unit
                });
            }
        }
        else
        {
            foreach (var row in sheet.RowsUsed().Where(r => r.RowNumber() >= 6))
            {
                if (row.Cell(1).IsEmpty() && row.Cell(2).IsEmpty() && row.Cell(3).IsEmpty() &&
                    row.Cell(4).IsEmpty() && row.Cell(5).IsEmpty() && row.Cell(6).IsEmpty())
                    continue;

                list.Add(new RequestItem
                {
                    Category = row.Cell(1).GetString(),
                    Brand = row.Cell(2).GetString(),
                    Model = row.Cell(3).GetString(),
                    Specifications = row.Cell(4).GetString(),
                    Quantity = row.Cell(5).GetValue<decimal>(),
                    Unit = row.Cell(6).GetString()
                });
            }
        }

        return list;
    }
}
