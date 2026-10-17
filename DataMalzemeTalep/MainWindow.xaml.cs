using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
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

        var scrollViewer = FindVisualChild<ScrollViewer>(grid);
        if (scrollViewer is null)
            return;

        if (e.Delta < 0 && scrollViewer.VerticalOffset < scrollViewer.ScrollableHeight ||
            e.Delta > 0 && scrollViewer.VerticalOffset > 0)
        {
            scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
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
        if (CatalogGrid.SelectedItem is not Material material)
        {
            MessageBox.Show("Önce silinecek ürünü seçin.", "Ürün seçilmedi");
            return;
        }

        if (MessageBox.Show($"'{material.Brand} {material.Model}' katalogdan silinsin mi?",
            "Ürünü Sil", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        _materials.Remove(material);
        CatalogGrid.ItemsSource = null;
        CatalogGrid.ItemsSource = _materials;
        CatalogCountText.Text = _materials.Count.ToString();
        RefreshCatalogSelectors();
        try { _service.SaveCatalog(_materials); }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Katalog kaydedilemedi"); }
    }

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
