using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DataMalzemeTalep.Models;
using DataMalzemeTalep.Services;
namespace DataMalzemeTalep
{
 public partial class MainWindow : Window
 {
  private readonly DataService _service=new();
  private readonly ObservableCollection<RequestItem> _items=[];
  private List<Material> _materials=[];
  public MainWindow(){InitializeComponent();RequestGrid.ItemsSource=_items;try{_materials=_service.LoadCatalog();CatalogGrid.ItemsSource=_materials;CatalogCountText.Text=_materials.Count.ToString();CategoryCombo.ItemsSource=_materials.Select(m=>m.Category).Distinct().OrderBy(v=>v).ToList();RefreshRequests();}catch(Exception ex){MessageBox.Show(ex.Message,"Katalog hatası");}ShowPage("dashboard");UpdateCount();}
  private void ShowPage(string p){DashboardPanel.Visibility=p=="dashboard"?Visibility.Visible:Visibility.Collapsed;NewRequestPanel.Visibility=p=="new"?Visibility.Visible:Visibility.Collapsed;PreviousPanel.Visibility=p=="previous"?Visibility.Visible:Visibility.Collapsed;CatalogPanel.Visibility=p=="catalog"?Visibility.Visible:Visibility.Collapsed;SetNav(DashboardNav,p=="dashboard");SetNav(NewRequestNav,p=="new");SetNav(PreviousNav,p=="previous");SetNav(CatalogNav,p=="catalog");PageTitle.Text=p switch{"new"=>"Yeni Talep","previous"=>"Önceki Talepler","catalog"=>"Malzeme Kataloğu",_=>"Genel Bakış"};PageSubtitle.Text=p switch{"new"=>"Katalogdan ürün seçin ve talep listenizi oluşturun.","previous"=>"Kaydedilmiş Excel taleplerinizi açın.","catalog"=>$"XML kataloğunda {_materials.Count} malzeme tanımlı.",_=>"Malzeme taleplerinizi tek yerden yönetin."};if(p=="previous")RefreshRequests();}
  private static void SetNav(Button b,bool a){b.Background=a?new SolidColorBrush(Color.FromRgb(32,43,73)):Brushes.Transparent;b.Foreground=a?new SolidColorBrush(Color.FromRgb(199,210,254)):new SolidColorBrush(Color.FromRgb(169,184,204));}
  private void DashboardNav_Click(object s,RoutedEventArgs e)=>ShowPage("dashboard");private void NewRequestNav_Click(object s,RoutedEventArgs e)=>ShowPage("new");private void PreviousNav_Click(object s,RoutedEventArgs e)=>ShowPage("previous");private void CatalogNav_Click(object s,RoutedEventArgs e)=>ShowPage("catalog");
  private void CategoryCombo_SelectionChanged(object s,SelectionChangedEventArgs e){if(ProductCombo is null||CategoryCombo.SelectedItem is not string cat)return;ProductCombo.ItemsSource=_materials.Where(m=>m.Category==cat).OrderBy(m=>m.Brand).ThenBy(m=>m.Model).ToList();ProductCombo.SelectedIndex=-1;}
  private void ProductCombo_SelectionChanged(object s,SelectionChangedEventArgs e){if(ProductCombo.SelectedItem is Material m)ProductDescriptionText.Text=$"{m.Specifications} • {m.Description}";}
  private void AddItem_Click(object s,RoutedEventArgs e){if(ProductCombo.SelectedItem is not Material m){MessageBox.Show("Önce bir malzeme seçin.");return;}if(!decimal.TryParse(QuantityTextBox.Text,out var q)||q<=0){MessageBox.Show("Miktar sıfırdan büyük bir sayı olmalı.");return;}var old=_items.FirstOrDefault(i=>i.Category==m.Category&&i.Brand==m.Brand&&i.Model==m.Model);if(old!=null)old.Quantity+=q;else _items.Add(new(){Category=m.Category,Brand=m.Brand,Model=m.Model,Specifications=m.Specifications,Unit=m.Unit,Quantity=q});RequestGrid.Items.Refresh();QuantityTextBox.Text="1";StatusText.Text="Liste değişti; Excel'e kaydedin.";UpdateCount();}
  private void RemoveItem_Click(object s,RoutedEventArgs e){if(s is Button{Tag:RequestItem item}){_items.Remove(item);UpdateCount();StatusText.Text="Liste değişti; Excel'e kaydedin.";}}
  private void SaveRequest_Click(object s,RoutedEventArgs e){try{if(_items.Count==0){MessageBox.Show("Önce listeye malzeme ekleyin.");return;}var name=RequestNameTextBox.Text.Trim();if(name.Length==0){MessageBox.Show("Dosya adı girin.");return;}var path=_service.SaveRequest(new(){Name=name,Items=new ObservableCollection<RequestItem>(_items)});StatusText.Text="Kaydedildi: "+Path.GetFileName(path);RefreshRequests();MessageBox.Show("Excel dosyası kaydedildi:\n"+path,"Kayıt tamamlandı");}catch(Exception ex){MessageBox.Show(ex.Message,"Kayıt hatası");}}
  private void UpdateCount(){if(RequestCountText!=null)RequestCountText.Text=_items.Count==0?"Henüz ürün eklenmedi":$"{_items.Count} ürün satırı";if(TotalItemsText!=null)TotalItemsText.Text=_items.Count.ToString();}
  private void RefreshRequests_Click(object s,RoutedEventArgs e)=>RefreshRequests();private void RefreshRequests(){if(PreviousList==null)return;PreviousList.ItemsSource=_service.GetSavedRequestFiles().Select(p=>new SavedRequest(p)).ToList();if(TotalRequestsText!=null)TotalRequestsText.Text=_service.GetSavedRequestFiles().Count.ToString();}
  private void PreviousList_MouseDoubleClick(object s,MouseButtonEventArgs e){if(PreviousList.SelectedItem is not SavedRequest r)return;try{_items.Clear();foreach(var i in _service.LoadRequest(r.Path))_items.Add(i);RequestNameTextBox.Text=r.Name;StatusText.Text="Açıldı: "+r.Name;UpdateCount();ShowPage("new");}catch(Exception ex){MessageBox.Show(ex.Message,"Dosya açılamadı");}}
  private void ExportPdf_Click(object s,RoutedEventArgs e)=>MessageBox.Show("PDF çıktısını sonraki adımda ekleyeceğiz. Excel kaydı hazır.","PDF Çıkar");
  private sealed class SavedRequest(string path){public string Path{get;}=path;public string Name=>System.IO.Path.GetFileNameWithoutExtension(Path);public override string ToString()=>Name+" • "+File.GetLastWriteTime(Path).ToString("dd.MM.yyyy HH:mm");}
 }
}