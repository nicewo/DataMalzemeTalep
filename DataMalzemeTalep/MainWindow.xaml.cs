using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DataMalzemeTalep
{
    public partial class MainWindow : Window
    {
        private readonly Brush _activeNavBackground = new SolidColorBrush(Color.FromRgb(32, 43, 73));
        private readonly Brush _activeNavForeground = new SolidColorBrush(Color.FromRgb(199, 210, 254));
        private readonly Brush _normalNavBackground = Brushes.Transparent;
        private readonly Brush _normalNavForeground = new SolidColorBrush(Color.FromRgb(169, 184, 204));

        public MainWindow()
        {
            InitializeComponent();
            ShowPage("dashboard");
        }

        private void ShowPage(string page)
        {
            DashboardPanel.Visibility = page == "dashboard" ? Visibility.Visible : Visibility.Collapsed;
            NewRequestPanel.Visibility = page == "new" ? Visibility.Visible : Visibility.Collapsed;
            PreviousPanel.Visibility = page == "previous" ? Visibility.Visible : Visibility.Collapsed;
            CatalogPanel.Visibility = page == "catalog" ? Visibility.Visible : Visibility.Collapsed;

            SetNavState(DashboardNav, page == "dashboard");
            SetNavState(NewRequestNav, page == "new");
            SetNavState(PreviousNav, page == "previous");
            SetNavState(CatalogNav, page == "catalog");

            switch (page)
            {
                case "new":
                    PageTitle.Text = "Yeni Talep";
                    PageSubtitle.Text = "Yeni bir malzeme listesi oluşturun.";
                    break;
                case "previous":
                    PageTitle.Text = "Önceki Talepler";
                    PageSubtitle.Text = "Kayıtlı taleplerinizi görüntüleyin ve düzenleyin.";
                    break;
                case "catalog":
                    PageTitle.Text = "Malzeme Kataloğu";
                    PageSubtitle.Text = "XML tabanlı ürün kataloğunu yönetin.";
                    break;
                default:
                    PageTitle.Text = "Genel Bakış";
                    PageSubtitle.Text = "Malzeme taleplerinizi tek yerden yönetin.";
                    break;
            }
        }

        private static void SetNavState(Button button, bool active)
        {
            button.Background = active
                ? new SolidColorBrush(Color.FromRgb(32, 43, 73))
                : Brushes.Transparent;
            button.Foreground = active
                ? new SolidColorBrush(Color.FromRgb(199, 210, 254))
                : new SolidColorBrush(Color.FromRgb(169, 184, 204));
        }

        private void DashboardNav_Click(object sender, RoutedEventArgs e) => ShowPage("dashboard");
        private void NewRequestNav_Click(object sender, RoutedEventArgs e) => ShowPage("new");
        private void PreviousNav_Click(object sender, RoutedEventArgs e) => ShowPage("previous");
        private void CatalogNav_Click(object sender, RoutedEventArgs e) => ShowPage("catalog");

        private void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "PDF çıktısı, talep ve Excel kayıt altyapısı eklendiğinde etkinleştirilecek.",
                "PDF Çıkar",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}