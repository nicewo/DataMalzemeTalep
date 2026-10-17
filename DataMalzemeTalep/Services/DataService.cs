using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Xml.Serialization;
using ClosedXML.Excel;
using DataMalzemeTalep.Models;


namespace DataMalzemeTalep.Services;

public sealed class DataService
{
    public string Root { get; } = AppContext.BaseDirectory; public string CatalogPath => Path.Combine(Root, "catalog.xml"); public string RequestsPath => Path.Combine(Root, "Listeler");
    public DataService() { Directory.CreateDirectory(Path.GetDirectoryName(CatalogPath)!); Directory.CreateDirectory(RequestsPath); if (!File.Exists(CatalogPath)) CreateCatalog(); }
    public List<Material> LoadCatalog() { using var s = File.OpenRead(CatalogPath); return ((MaterialCatalog?)new XmlSerializer(typeof(MaterialCatalog)).Deserialize(s))?.Materials ?? []; }
    private void CreateCatalog() { var c = new MaterialCatalog(); using var s = new FileStream(CatalogPath, FileMode.CreateNew, FileAccess.Write, FileShare.None); new XmlSerializer(typeof(MaterialCatalog)).Serialize(s, c); }
    public List<string> GetSavedRequestFiles() => Directory.GetFiles(RequestsPath, "*.xlsx").OrderByDescending(File.GetLastWriteTime).ToList();
    public string SaveRequest(RequestDocument r) { if (r.Items.Count == 0) throw new InvalidOperationException("Liste boş."); var bad = Path.GetInvalidFileNameChars(); var name = new string(r.Name.Select(c => bad.Contains(c) ? '_' : c).ToArray()).Trim(); if (name.Length == 0) name = "Malzeme_Talebi"; var path = Path.Combine(RequestsPath, name + ".xlsx"); if (File.Exists(path)) path = Path.Combine(RequestsPath, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"); using var wb = new XLWorkbook(); var sh = wb.Worksheets.Add("Malzeme Talebi"); sh.Cell(1, 1).Value = "DATA OTOMASYON"; sh.Range(1, 1, 1, 6).Merge(); sh.Cell(1, 1).Style.Font.Bold = true; sh.Cell(1, 1).Style.Font.FontSize = 18; sh.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#D9A514"); sh.Cell(2, 1).Value = r.Name; sh.Range(2, 1, 2, 6).Merge(); sh.Cell(3, 1).Value = "Oluşturulma tarihi"; sh.Cell(3, 2).Value = r.CreatedAt; sh.Cell(3, 2).Style.DateFormat.Format = "dd.MM.yyyy HH:mm"; string[] hs = ["Kategori", "Marka", "Model", "Teknik Özellikler", "Miktar", "Birim"]; for (int i = 0; i < hs.Length; i++) { var cell = sh.Cell(5, i + 1); cell.Value = hs[i]; cell.Style.Font.Bold = true; cell.Style.Font.FontColor = XLColor.White; cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#202B49"); } for (int i = 0; i < r.Items.Count; i++) { var it = r.Items[i]; int row = i + 6; sh.Cell(row, 1).Value = it.Category; sh.Cell(row, 2).Value = it.Brand; sh.Cell(row, 3).Value = it.Model; sh.Cell(row, 4).Value = it.Specifications; sh.Cell(row, 5).Value = it.Quantity; sh.Cell(row, 6).Value = it.Unit; } sh.Columns().AdjustToContents(); sh.SheetView.FreezeRows(5); wb.SaveAs(path); return path; }
    public List<RequestItem> LoadRequest(string path) { using var wb = new XLWorkbook(path); var sh = wb.Worksheet(1); var list = new List<RequestItem>(); foreach (var row in sh.RowsUsed().Where(r => r.RowNumber() >= 6)) list.Add(new() { Category = row.Cell(1).GetString(), Brand = row.Cell(2).GetString(), Model = row.Cell(3).GetString(), Specifications = row.Cell(4).GetString(), Quantity = row.Cell(5).GetValue<decimal>(), Unit = row.Cell(6).GetString() }); return list; }
}