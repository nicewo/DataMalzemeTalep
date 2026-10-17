using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using ClosedXML.Excel;
using DataMalzemeTalep.Models;
using DataMalzemeTalep.Services;

namespace DataMalzemeTalep;

public partial class MainWindow : Window
{
    private readonly DataService _service = new();
    private readonly ObservableCollection<RequestItem> _items = [];
    private List<Material> _materials = [];
    private string? _currentListPath;
    private DateTime _currentListCreatedAt = DateTime.Now;

    public MainWindow()
    {
        InitializeComponent();
        RequestGrid.ItemsSource = _items;
        CatalogGrid.AddHandler(Mouse.PreviewMouseWheelEvent,
            new MouseWheelEventHandler(CatalogGrid_PreviewMouseWheel), true);


        try
        {
            _materials = _service.LoadCatalog();
            CatalogGrid.ItemsSource = _materials;
            CatalogCountText.Text = _materials.Count.ToString();
            CategoryCombo.ItemsSource = _materials.Select(m => m.Category).Distinct().OrderBy(v => v).ToList();
            RefreshRequests();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Katalog hatası");
        }

        ShowPage("dashboard");
        UpdateCount();
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        // Base breakpoints on window size; the left navigation consumes a fixed
        // portion of the window, so keep the dashboard layout conservative.
        var compact = ActualWidth < 1050;
        var narrow = ActualWidth < 680;

        DashboardArtColumn.Width = compact ? new GridLength(0) : new GridLength(150);
        DashboardHeroArt.Visibility = compact ? Visibility.Collapsed : Visibility.Visible;

        if (narrow)
        {
            // Very narrow: one full-width statistic card per row.
            StatsColumn1.Width = new GridLength(1, GridUnitType.Star);
            StatsColumn2.Width = new GridLength(0);
            StatsColumn3.Width = new GridLength(0);
            Grid.SetRow(TotalListsCard, 0);
            Grid.SetColumn(TotalListsCard, 0);
            Grid.SetRow(LastCreatedCard, 1);
            Grid.SetColumn(LastCreatedCard, 0);
            Grid.SetRow(CatalogStatsCard, 2);
            Grid.SetColumn(CatalogStatsCard, 0);
            Grid.SetColumnSpan(TotalListsCard, 1);
            Grid.SetColumnSpan(LastCreatedCard, 1);
            Grid.SetColumnSpan(CatalogStatsCard, 1);
            TotalListsCard.Margin = new Thickness(0, 0, 0, 10);
            LastCreatedCard.Margin = new Thickness(0, 0, 0, 10);
            CatalogStatsCard.Margin = new Thickness(0);
            StatsRow1.Height = GridLength.Auto;
            StatsRow2.Height = GridLength.Auto;
        }
        else if (compact)
        {
            // Medium: two cards on the first row, catalog card spans the second.
            StatsColumn1.Width = new GridLength(1, GridUnitType.Star);
            StatsColumn2.Width = new GridLength(1, GridUnitType.Star);
            StatsColumn3.Width = new GridLength(0);
            Grid.SetRow(TotalListsCard, 0);
            Grid.SetColumn(TotalListsCard, 0);
            Grid.SetRow(LastCreatedCard, 0);
            Grid.SetColumn(LastCreatedCard, 1);
            Grid.SetRow(CatalogStatsCard, 1);
            Grid.SetColumn(CatalogStatsCard, 0);
            Grid.SetColumnSpan(TotalListsCard, 1);
            Grid.SetColumnSpan(LastCreatedCard, 1);
            Grid.SetColumnSpan(CatalogStatsCard, 2);
            TotalListsCard.Margin = new Thickness(0, 0, 8, 0);
            LastCreatedCard.Margin = new Thickness(8, 0, 0, 0);
            CatalogStatsCard.Margin = new Thickness(0, 10, 0, 0);
            StatsRow1.Height = GridLength.Auto;
            StatsRow2.Height = GridLength.Auto;
        }
        else
        {
            // Wide: three equal statistic cards in one row.
            StatsColumn1.Width = new GridLength(1, GridUnitType.Star);
            StatsColumn2.Width = new GridLength(1, GridUnitType.Star);
            StatsColumn3.Width = new GridLength(1, GridUnitType.Star);
            Grid.SetRow(TotalListsCard, 0);
            Grid.SetColumn(TotalListsCard, 0);
            Grid.SetRow(LastCreatedCard, 0);
            Grid.SetColumn(LastCreatedCard, 1);
            Grid.SetRow(CatalogStatsCard, 0);
            Grid.SetColumn(CatalogStatsCard, 2);
            Grid.SetColumnSpan(TotalListsCard, 1);
            Grid.SetColumnSpan(LastCreatedCard, 1);
            Grid.SetColumnSpan(CatalogStatsCard, 1);
            TotalListsCard.Margin = new Thickness(0, 0, 8, 0);
            CatalogStatsCard.Margin = new Thickness(8, 0, 0, 0);
            StatsRow1.Height = GridLength.Auto;
            StatsRow2.Height = new GridLength(0);
        }

        if (compact)
        {
            QuickActionsColumn.Width = new GridLength(1, GridUnitType.Star);
            GettingStartedColumn.Width = new GridLength(0);
            Grid.SetRow(QuickActionsCard, 0);
            Grid.SetColumn(QuickActionsCard, 0);
            Grid.SetRow(GettingStartedCard, 1);
            Grid.SetColumn(GettingStartedCard, 0);
            QuickActionsCard.Margin = new Thickness(0, 0, 0, 12);
            GettingStartedCard.Margin = new Thickness(0);
            QuickActionsRow.Height = GridLength.Auto;
            GettingStartedRow.Height = GridLength.Auto;
        }
        else
        {
            QuickActionsColumn.Width = new GridLength(1.25, GridUnitType.Star);
            GettingStartedColumn.Width = new GridLength(1, GridUnitType.Star);
            Grid.SetRow(QuickActionsCard, 0);
            Grid.SetColumn(QuickActionsCard, 0);
            Grid.SetRow(GettingStartedCard, 0);
            Grid.SetColumn(GettingStartedCard, 1);
            QuickActionsCard.Margin = new Thickness(0, 0, 12, 0);
            GettingStartedCard.Margin = new Thickness(12, 0, 0, 0);
            QuickActionsRow.Height = GridLength.Auto;
            GettingStartedRow.Height = GridLength.Auto;
        }

        if (DashboardPanel.Visibility == Visibility.Visible)
            MainContentScrollViewer.ScrollToTop();
    }

    private void MinimizeWindow_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void MaximizeWindow_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private void CloseWindow_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ShowPage(string page)
    {
        DashboardPanel.Visibility = page == "dashboard" ? Visibility.Visible : Visibility.Collapsed;
        NewRequestPanel.Visibility = page == "new" ? Visibility.Visible : Visibility.Collapsed;
        PreviousPanel.Visibility = page == "previous" ? Visibility.Visible : Visibility.Collapsed;
        CatalogPanel.Visibility = page == "catalog" ? Visibility.Visible : Visibility.Collapsed;

        SetNav(DashboardNav, page == "dashboard");
        SetNav(NewRequestNav, page == "new");
        SetNav(PreviousNav, page == "previous");
        SetNav(CatalogNav, page == "catalog");

        PageTitle.Text = page switch
        {
            "new" => "Liste Düzenleme",
            "previous" => "Kayıtlı Listeler",
            "catalog" => "Malzeme Kataloğu",
            _ => "Genel Bakış"
        };
        PageSubtitle.Text = page switch
        {
            "new" => "Malzeme ekleyin; değişiklikler açık Excel listesine kaydedilir.",
            "previous" => "Kaydedilmiş Excel listelerinizi açın.",
            "catalog" => $"XML kataloğunda {_materials.Count} malzeme tanımlı.",
            _ => "Malzeme listelerinizi tek yerden yönetin."
        };

        if (page == "previous")
            RefreshRequests();
    }

    private static void SetNav(Button button, bool active)
    {
        button.Background = active
            ? new SolidColorBrush(Color.FromRgb(32, 43, 73))
            : Brushes.Transparent;
        button.Foreground = active
            ? new SolidColorBrush(Color.FromRgb(199, 210, 254))
            : new SolidColorBrush(Color.FromRgb(169, 184, 204));
    }

    private void DashboardNav_Click(object sender, RoutedEventArgs e) => ShowPage("dashboard");

    private void NewRequestNav_Click(object sender, RoutedEventArgs e)
    {
        var name = AskListName();
        if (string.IsNullOrWhiteSpace(name))
            return;

        name = name.Trim();
        try
        {
            var path = _service.SaveRequest(new RequestDocument
            {
                Name = name,
                CreatedAt = DateTime.Now,
                Items = []
            });

            _currentListPath = path;
            _currentListCreatedAt = DateTime.Now;
            _items.Clear();
            RequestNameTextBox.Text = name;
            CurrentListTitle.Text = name;
            StatusText.Text = "Boş Excel listesi oluşturuldu. Malzeme ekleyebilirsiniz.";
            UpdateCount();
            RefreshRequests();
            ShowPage("new");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Liste oluşturulamadı");
        }
    }

    private string? AskListName()
    {
        var dialog = new Window
        {
            Title = "Yeni Liste",
            Width = 430,
            Height = 205,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            Background = new SolidColorBrush(Color.FromRgb(17, 28, 48)),
            Foreground = Brushes.White
        };

        var panel = new StackPanel { Margin = new Thickness(22) };
        panel.Children.Add(new TextBlock
        {
            Text = "Liste adı",
            FontSize = 17,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 12)
        });

        var input = new TextBox
        {
            Height = 36,
            Padding = new Thickness(9, 7, 9, 7),
            FontSize = 14,
            Background = new SolidColorBrush(Color.FromRgb(23, 36, 58)),
            Foreground = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(51, 68, 95))
        };
        panel.Children.Add(input);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        var cancel = new Button { Content = "İptal", Width = 85, Height = 32, Margin = new Thickness(0, 0, 8, 0) };
        cancel.Click += (_, _) => dialog.DialogResult = false;
        var create = new Button { Content = "Liste Oluştur", Width = 115, Height = 32, IsDefault = true };
        create.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(input.Text))
            {
                MessageBox.Show(dialog, "Bir liste adı girin.", "Eksik bilgi");
                input.Focus();
                return;
            }
            dialog.DialogResult = true;
        };
        buttons.Children.Add(cancel);
        buttons.Children.Add(create);
        panel.Children.Add(buttons);
        dialog.Content = panel;
        dialog.Loaded += (_, _) => input.Focus();

        return dialog.ShowDialog() == true ? input.Text.Trim() : null;
    }

    private void PreviousNav_Click(object sender, RoutedEventArgs e) => ShowPage("previous");
    private void CatalogNav_Click(object sender, RoutedEventArgs e) => ShowPage("catalog");

    private void CategoryCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ProductCombo is null || CategoryCombo.SelectedItem is not string category)
            return;

        var categoryMaterials = _materials.Where(m => m.Category == category).ToList();
        BrandCombo.ItemsSource = categoryMaterials.Select(m => m.Brand).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().OrderBy(v => v).ToList();
        BrandCombo.SelectedIndex = -1;
        ProductCombo.ItemsSource = null;
        ProductDescriptionText.Text = "Marka ve model seçin.";
    }

    private void BrandCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ProductCombo is null || CategoryCombo.SelectedItem is not string category ||
            BrandCombo.SelectedItem is not string brand)
            return;

        ProductCombo.ItemsSource = _materials
            .Where(m => m.Category == category && m.Brand == brand)
            .OrderBy(m => m.Model)
            .ToList();
        ProductCombo.SelectedIndex = -1;
        ProductDescriptionText.Text = "Model seçildiğinde açıklama ve birim burada görünür.";
    }

    private void ProductCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ProductCombo.SelectedItem is Material material)
            ProductDescriptionText.Text = $"Açıklama: {material.Specifications} {material.Description}    |    Birim: {material.Unit}";
    }

    private void AddItem_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentListPath))
        {
            MessageBox.Show("Önce yeni bir liste oluşturun.", "Liste seçilmedi");
            return;
        }

        if (ProductCombo.SelectedItem is not Material material)
        {
            MessageBox.Show("Önce bir malzeme seçin.");
            return;
        }

        if (!decimal.TryParse(QuantityTextBox.Text, out var quantity) || quantity <= 0)
        {
            MessageBox.Show("Miktar sıfırdan büyük bir sayı olmalı.");
            return;
        }

        var existing = _items.FirstOrDefault(i =>
            i.Category == material.Category &&
            i.Brand == material.Brand &&
            i.Model == material.Model);

        if (existing is not null)
            existing.Quantity += quantity;
        else
            _items.Add(new RequestItem
            {
                Category = material.Category,
                Brand = material.Brand,
                Model = material.Model,
                Specifications = material.Specifications,
                Unit = material.Unit,
                Quantity = quantity
            });

        RequestGrid.Items.Refresh();
        QuantityTextBox.Text = "1";
        UpdateCount();
        SaveCurrentListSilently();
    }

    private void RemoveItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: RequestItem item })
        {
            _items.Remove(item);
            UpdateCount();
            SaveCurrentListSilently();
        }
    }

    private void SaveRequest_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentListPath))
        {
            MessageBox.Show("Önce yeni bir liste oluşturun.");
            return;
        }

        try
        {
            SaveCurrentListSilently();
            StatusText.Text = "Liste kaydedildi: " + Path.GetFileName(_currentListPath);
            RefreshRequests();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Kayıt hatası");
        }
    }

    private void SaveCurrentListSilently()
    {
        if (string.IsNullOrWhiteSpace(_currentListPath))
            return;

        var name = Path.GetFileNameWithoutExtension(_currentListPath);
        _service.SaveRequest(new RequestDocument
        {
            Name = name,
            CreatedAt = _currentListCreatedAt,
            Items = new ObservableCollection<RequestItem>(_items)
        }, _currentListPath);

        StatusText.Text = "Değişiklikler Excel listesine kaydedildi.";
        RefreshRequests();
    }

    private void UpdateCount()
    {
        if (RequestCountText is not null)
            RequestCountText.Text = _items.Count == 0 ? "Henüz malzeme eklenmedi" : $"{_items.Count} malzeme satırı";
    }

    private void RefreshRequests_Click(object sender, RoutedEventArgs e) => RefreshRequests();

    private void RefreshRequests()
    {
        var files = _service.GetSavedRequestFiles();
        var savedRequests = files.Select(path => new SavedRequest(path)).ToList();

        if (PreviousList is not null)
            PreviousList.ItemsSource = savedRequests;

        if (SidebarSavedLists is not null)
            SidebarSavedLists.ItemsSource = savedRequests;

        if (TotalRequestsText is not null)
            TotalRequestsText.Text = files.Count.ToString();

        if (LastCreatedNameText is not null && LastCreatedDateText is not null)
        {
            if (files.Count > 0)
            {
                var latestPath = files[0];
                LastCreatedNameText.Text = Path.GetFileNameWithoutExtension(latestPath);
                LastCreatedDateText.Text = File.GetLastWriteTime(latestPath).ToString("dd.MM.yyyy HH:mm");
                LastCreatedCard.ToolTip = latestPath;
            }
            else
            {
                LastCreatedNameText.Text = "—";
                LastCreatedDateText.Text = "Henüz liste kaydı yok";
                LastCreatedCard.ToolTip = null;
            }
        }
    }

    private void PreviousList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (PreviousList.SelectedItem is SavedRequest request)
            OpenSavedRequest(request);
    }

    private void SidebarSavedLists_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (SidebarSavedLists.SelectedItem is SavedRequest request)
            OpenSavedRequest(request);
    }

    private void OpenSavedRequest(SavedRequest request)
    {
        try
        {
            _items.Clear();
            foreach (var item in _service.LoadRequest(request.Path))
                _items.Add(item);

            _currentListPath = request.Path;
            _currentListCreatedAt = File.GetCreationTime(request.Path);
            RequestNameTextBox.Text = request.Name;
            CurrentListTitle.Text = request.Name;
            StatusText.Text = "Açıldı: " + request.Name;
            UpdateCount();
            ShowPage("new");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Liste açılamadı");
        }
    }

    private void CatalogGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not DataGrid grid)
            return;

        // First scroll the DataGrid itself. When it reaches its boundary (or its
        // rows do not overflow), forward the wheel to the page ScrollViewer.
        var gridScrollViewer = FindVisualChild<ScrollViewer>(grid);
        if (gridScrollViewer is not null &&
            ((e.Delta < 0 && gridScrollViewer.VerticalOffset < gridScrollViewer.ScrollableHeight) ||
             (e.Delta > 0 && gridScrollViewer.VerticalOffset > 0)))
        {
            gridScrollViewer.ScrollToVerticalOffset(
                gridScrollViewer.VerticalOffset - e.Delta / 3.0);
            e.Handled = true;
            return;
        }

        if (MainContentScrollViewer is not null)
        {
            MainContentScrollViewer.ScrollToVerticalOffset(
                MainContentScrollViewer.VerticalOffset - e.Delta / 3.0);
            e.Handled = true;
        }
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T target)
                return target;

            var nested = FindVisualChild<T>(child);
            if (nested is not null)
                return nested;
        }

        return null;
    }

    private void AddCatalogItem_Click(object sender, RoutedEventArgs e)
    {
        var material = new Material
        {
            Category = "",
            Brand = "",
            Model = "",
            Specifications = "",
            Description = "",
            Unit = "Adet"
        };
        _materials.Add(material);
        CatalogGrid.ItemsSource = null;
        CatalogGrid.ItemsSource = _materials;
        CatalogGrid.SelectedItem = material;
        CatalogGrid.ScrollIntoView(material);
        CatalogGrid.BeginEdit();
        CatalogCountText.Text = _materials.Count.ToString();
    }

    private void DeleteCatalogItem_Click(object sender, RoutedEventArgs e)
    {
        var selectedMaterials = CatalogGrid.SelectedItems
            .OfType<Material>()
            .Distinct()
            .ToList();

        if (selectedMaterials.Count == 0)
        {
            MessageBox.Show("Silmek için bir veya daha fazla ürün seçin.\n\n" +
                "Bir aralık seçmek için Shift, tek tek ürün seçmek veya seçimi değiştirmek için Ctrl tuşunu kullanabilirsiniz.",
                "Ürün seçilmedi", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var preview = string.Join("\n", selectedMaterials.Take(8)
            .Select(m => $"• {m.Brand} {m.Model}".Trim()));
        if (selectedMaterials.Count > 8)
            preview += $"\n… ve {selectedMaterials.Count - 8} ürün daha";

        if (MessageBox.Show(
            $"Seçili {selectedMaterials.Count} ürün katalogdan silinsin mi?\n\n{preview}",
            "Seçili Ürünleri Sil", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        foreach (var material in selectedMaterials)
            _materials.Remove(material);

        CatalogGrid.ItemsSource = null;
        CatalogGrid.ItemsSource = _materials;
        CatalogCountText.Text = _materials.Count.ToString();
        RefreshCatalogSelectors();

        try
        {
            _service.SaveCatalog(_materials);
            MessageBox.Show($"{selectedMaterials.Count} ürün silindi.", "Katalog güncellendi",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Katalog kaydedilemedi:\n" + ex.Message,
                "Kayıt hatası", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ImportCatalogPdf_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Malzeme listesi PDF dosyasını seçin",
            Filter = "PDF dosyaları (*.pdf)|*.pdf",
            Multiselect = false,
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) != true)
            return;

        try
        {
            var importer = new PdfCatalogImportService();
            var extracted = importer.ReadMaterials(dialog.FileName);

            if (extracted.Count == 0)
            {
                MessageBox.Show(
                    "PDF içinde okunabilir malzeme satırı bulunamadı. PDF taranmış görüntü olabilir veya tablo biçimi desteklenmiyor. Metin seçilebilen bir PDF ile tekrar deneyin.",
                    "PDF okunamadı", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var added = 0;
            var updated = 0;

            foreach (var incoming in extracted)
            {
                // A model code is the stable key. If the PDF has no model/product
                // code, add it as a separate entry rather than guessing a match.
                var existing = string.IsNullOrWhiteSpace(incoming.Model)
                    ? null
                    : _materials.FirstOrDefault(m =>
                        !string.IsNullOrWhiteSpace(m.Model) &&
                        string.Equals(NormalizeModel(m.Model), NormalizeModel(incoming.Model),
                            StringComparison.OrdinalIgnoreCase));

                if (existing is not null)
                {
                    existing.Category = incoming.Category;
                    existing.Brand = incoming.Brand;
                    existing.Model = incoming.Model;
                    existing.Specifications = incoming.Specifications;
                    existing.Description = incoming.Description;
                    existing.Unit = incoming.Unit;
                    updated++;
                }
                else
                {
                    _materials.Add(incoming);
                    added++;
                }
            }

            _service.SaveCatalog(_materials);
            CatalogGrid.ItemsSource = null;
            CatalogGrid.ItemsSource = _materials;
            CatalogCountText.Text = _materials.Count.ToString();
            RefreshCatalogSelectors();

            MessageBox.Show(
                $"PDF okundu: {Path.GetFileName(dialog.FileName)}\n\n" +
                $"Yeni ürün: {added}\nMevcut modeli güncellenen: {updated}\n" +
                $"Katalogdaki toplam ürün: {_materials.Count}",
                "PDF içe aktarma tamamlandı", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("PDF içe aktarılırken hata oluştu:\n" + ex.Message,
                "PDF içe aktarma hatası", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string NormalizeModel(string model) =>
        string.Join("", model.Trim().Where(c => !char.IsWhiteSpace(c))).ToUpperInvariant();

    private void ExportCatalogExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            CatalogGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CatalogGrid.CommitEdit(DataGridEditingUnit.Row, true);

            var dialog = new SaveFileDialog
            {
                Title = "Kataloğu Excel olarak dışa aktar",
                Filter = "Excel çalışma kitabı (*.xlsx)|*.xlsx",
                FileName = $"MalzemeKatalogu_{DateTime.Now:yyyyMMdd}.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };

            if (dialog.ShowDialog(this) != true)
                return;

            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Katalog");
            string[] headers = ["Kategori", "Marka", "Ürün Kodu", "Teknik Özellikler", "Açıklama", "Birim"];
            for (var col = 1; col <= headers.Length; col++)
            {
                var cell = sheet.Cell(1, col);
                cell.Value = headers[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1B2A4B");
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            for (var i = 0; i < _materials.Count; i++)
            {
                var material = _materials[i];
                var row = i + 2;
                sheet.Cell(row, 1).Value = material.Category;
                sheet.Cell(row, 2).Value = material.Brand;
                sheet.Cell(row, 3).Value = material.Model;
                sheet.Cell(row, 4).Value = material.Specifications;
                sheet.Cell(row, 5).Value = material.Description;
                sheet.Cell(row, 6).Value = material.Unit;
            }

            sheet.SheetView.FreezeRows(1);
            sheet.RangeUsed()?.SetAutoFilter();
            sheet.Columns().AdjustToContents();
            sheet.Column(4).Width = Math.Min(Math.Max(sheet.Column(4).Width, 24), 55);
            sheet.Column(5).Width = Math.Min(Math.Max(sheet.Column(5).Width, 20), 45);
            sheet.Columns().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            workbook.SaveAs(dialog.FileName);

            MessageBox.Show($"Katalog Excel'e aktarıldı.\n{dialog.FileName}",
                "Dışa aktarma tamamlandı", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Katalog Excel'e aktarılamadı:\n" + ex.Message,
                "Dışa aktarma hatası", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ImportCatalogExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Düzenlenmiş katalog Excel dosyasını seçin",
            Filter = "Excel çalışma kitabı (*.xlsx)|*.xlsx",
            Multiselect = false,
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) != true)
            return;

        try
        {
            CatalogGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CatalogGrid.CommitEdit(DataGridEditingUnit.Row, true);

            using var workbook = new XLWorkbook(dialog.FileName);
            var sheet = workbook.Worksheets.FirstOrDefault();
            if (sheet is null || sheet.LastRowUsed() is null)
            {
                MessageBox.Show("Excel dosyasında okunabilir bir katalog tablosu bulunamadı.",
                    "İçe aktarma", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var lastHeaderColumn = sheet.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (var col = 1; col <= lastHeaderColumn; col++)
            {
                var header = NormalizeHeader(sheet.Cell(1, col).GetString());
                if (!string.IsNullOrWhiteSpace(header) && !headerMap.ContainsKey(header))
                    headerMap[header] = col;
            }

            int FindColumn(params string[] aliases)
            {
                foreach (var alias in aliases)
                    if (headerMap.TryGetValue(NormalizeHeader(alias), out var column))
                        return column;
                return 0;
            }

            var categoryCol = FindColumn("Kategori", "Category");
            var brandCol = FindColumn("Marka", "Brand");
            var modelCol = FindColumn("Ürün Kodu", "Urun Kodu", "Model", "Product Code", "ProductCode");
            var specsCol = FindColumn("Teknik Özellikler", "Teknik Ozellikler", "Specifications", "Özellikler");
            var descriptionCol = FindColumn("Açıklama", "Aciklama", "Description");
            var unitCol = FindColumn("Birim", "Unit");

            if (modelCol == 0)
            {
                MessageBox.Show("Excel başlıklarında 'Ürün Kodu' veya 'Model' sütunu bulunamadı. Dışa aktarılan şablonu kullan.",
                    "Geçersiz katalog dosyası", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var incomingMaterials = new List<Material>();
            for (var row = 2; row <= sheet.LastRowUsed()!.RowNumber(); row++)
            {
                string Read(int column) => column == 0 ? string.Empty : sheet.Cell(row, column).GetString().Trim();

                var incoming = new Material
                {
                    Category = Read(categoryCol),
                    Brand = Read(brandCol),
                    Model = Read(modelCol),
                    Specifications = Read(specsCol),
                    Description = Read(descriptionCol),
                    Unit = string.IsNullOrWhiteSpace(Read(unitCol)) ? "Adet" : Read(unitCol)
                };

                if (string.IsNullOrWhiteSpace(incoming.Category) &&
                    string.IsNullOrWhiteSpace(incoming.Brand) &&
                    string.IsNullOrWhiteSpace(incoming.Model) &&
                    string.IsNullOrWhiteSpace(incoming.Specifications) &&
                    string.IsNullOrWhiteSpace(incoming.Description))
                    continue;

                incomingMaterials.Add(incoming);
            }

            if (incomingMaterials.Count == 0)
            {
                MessageBox.Show("Excel dosyasında içe aktarılacak ürün satırı bulunamadı.",
                    "İçe aktarma", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var added = 0;
            var updated = 0;
            var noCodeAdded = 0;

            // Aynı Excel dosyası tekrar içe aktarıldığında kodsuz satırların
            // tekrar eklenmesini önlemek için her mevcut kayıt en fazla bir kez eşleşsin.
            var matchedWithoutCode = new HashSet<Material>();

            foreach (var incoming in incomingMaterials)
            {
                Material? existing = null;

                if (!string.IsNullOrWhiteSpace(incoming.Model))
                {
                    existing = _materials.FirstOrDefault(m =>
                        !string.IsNullOrWhiteSpace(m.Model) &&
                        string.Equals(NormalizeModel(m.Model), NormalizeModel(incoming.Model),
                            StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    // Ürün kodu yoksa kategori + marka + teknik özellik + açıklama + birim
                    // üzerinden eşleştir. Böylece dışa aktarılıp hiç değiştirilmemiş
                    // kodsuz ürünler yeniden eklenmez.
                    existing = _materials.FirstOrDefault(m =>
                        string.IsNullOrWhiteSpace(m.Model) &&
                        !matchedWithoutCode.Contains(m) &&
                        !string.IsNullOrWhiteSpace(m.Description) &&
                        SameCatalogText(m.Description, incoming.Description));

                    if (existing is not null)
                        matchedWithoutCode.Add(existing);
                }

                if (existing is not null)
                {
                    existing.Category = incoming.Category;
                    existing.Brand = incoming.Brand;
                    existing.Model = incoming.Model;
                    existing.Specifications = incoming.Specifications;
                    existing.Description = incoming.Description;
                    existing.Unit = incoming.Unit;
                    updated++;
                }
                else
                {
                    _materials.Add(incoming);
                    added++;
                    if (string.IsNullOrWhiteSpace(incoming.Model))
                        noCodeAdded++;
                }
            }

            _service.SaveCatalog(_materials);
            CatalogGrid.ItemsSource = null;
            CatalogGrid.ItemsSource = _materials;
            CatalogCountText.Text = _materials.Count.ToString();
            RefreshCatalogSelectors();

            MessageBox.Show(
                $"Excel içe aktarma tamamlandı.\n\nGüncellenen ürün: {updated}\nYeni eklenen ürün: {added}\n" +
                $"Ürün kodu olmayan yeni kayıt: {noCodeAdded}\nKatalog toplamı: {_materials.Count}",
                "Katalog içe aktarma", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Excel katalog içe aktarılırken hata oluştu:\n" + ex.Message,
                "İçe aktarma hatası", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static bool SameCatalogText(string? left, string? right) =>
        string.Equals(
            string.Join(" ", (left ?? string.Empty).Trim().ToUpperInvariant()
                .Replace('İ', 'I').Replace('Ş', 'S').Replace('Ğ', 'G')
                .Replace('Ü', 'U').Replace('Ö', 'O').Replace('Ç', 'C')
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)),
            string.Join(" ", (right ?? string.Empty).Trim().ToUpperInvariant()
                .Replace('İ', 'I').Replace('Ş', 'S').Replace('Ğ', 'G')
                .Replace('Ü', 'U').Replace('Ö', 'O').Replace('Ç', 'C')
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)),
            StringComparison.OrdinalIgnoreCase);

    private static string NormalizeHeader(string value) =>
        string.Join(" ", value.Trim().ToUpperInvariant().Replace('İ', 'I').Replace('Ş', 'S')
            .Replace('Ğ', 'G').Replace('Ü', 'U').Replace('Ö', 'O').Replace('Ç', 'C')
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private void SaveCatalog_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            CatalogGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CatalogGrid.CommitEdit(DataGridEditingUnit.Row, true);

            foreach (var material in _materials)
            {
                material.Category = material.Category?.Trim() ?? "";
                material.Brand = material.Brand?.Trim() ?? "";
                material.Model = material.Model?.Trim() ?? "";
                material.Specifications = material.Specifications?.Trim() ?? "";
                material.Description = material.Description?.Trim() ?? "";
                material.Unit = string.IsNullOrWhiteSpace(material.Unit) ? "Adet" : material.Unit.Trim();
            }

            _service.SaveCatalog(_materials);
            CatalogCountText.Text = _materials.Count.ToString();
            RefreshCatalogSelectors();
            MessageBox.Show("Katalog kaydedildi.", "Malzeme Kataloğu", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Katalog kaydedilemedi");
        }
    }

    private void RefreshCatalogSelectors()
    {
        CategoryCombo.ItemsSource = _materials.Select(m => m.Category)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct().OrderBy(v => v).ToList();
        BrandCombo.ItemsSource = null;
        ProductCombo.ItemsSource = null;
        ProductDescriptionText.Text = "Marka ve model seçin.";
    }

    private void ExportPdf_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("PDF çıktısı henüz uygulanmadı.", "PDF Çıkar");
    }

    private sealed class SavedRequest(string path)
    {
        public string Path { get; } = path;
        public string Name => System.IO.Path.GetFileNameWithoutExtension(Path);
        public string ModifiedText => File.GetLastWriteTime(Path).ToString("dd.MM.yyyy HH:mm");
        public override string ToString() => Name + " • " + ModifiedText;
    }
}
