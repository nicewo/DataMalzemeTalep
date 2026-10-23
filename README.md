# DataMalzemeTalep

[![.NET](https://img.shields.io/badge/.NET-10.0-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/)
[![WPF](https://img.shields.io/badge/UI-WPF-0078D4?logo=windows&logoColor=white)](https://learn.microsoft.com/dotnet/desktop/wpf/)
[![Platform](https://img.shields.io/badge/platform-Windows%20x64-0078D4?logo=windows&logoColor=white)](#)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE.txt)

**DataMalzemeTalep**, malzeme kataloğunu yönetmek, seçilen kalemlerden talep listeleri hazırlamak ve şirket formuna uygun Excel çıktıları üretmek için geliştirilmiş masaüstü uygulamasıdır.

## Özellikler

- Malzeme kataloğunu XML dosyasından yükleme ve kaydetme
- Marka, kategori, model/kod, açıklama, teknik özellik ve birim alanlarıyla katalog yönetimi
- Talep listesi oluşturma ve düzenleme
- Hazırlanan listeleri Excel çalışma kitabı olarak kaydetme
- Katalog verisini PDF'den içe aktarmaya yönelik servis
- Liste geçmişini yerel klasörde tutma
- Şirket logosu ve özel uygulama simgesi
- Koyu temalı, pencere boyutuna uyum sağlayan WPF arayüzü

## Teknoloji

- C# / .NET 10
- WPF
- ClosedXML — Excel dosyaları
- PdfPig — PDF içeriğini okuma
- XML — malzeme kataloğu

## Gereksinimler

- Windows x64
- .NET 10 SDK (kaynak koddan derlemek için)
- İnternet bağlantısı (NuGet paketlerini ilk kez geri yüklemek için)

## Derleme

Depoyu klonlayıp çözüm dosyasının bulunduğu dizinde çalıştırın:

```powershell
dotnet restore DataMalzemeTalep.slnx
dotnet build DataMalzemeTalep.slnx -c Release
```

## Bağımsız yayın

Proje dosyası Windows x64 için self-contained ve single-file yayın ayarları içerir:

```powershell
dotnet publish DataMalzemeTalep/DataMalzemeTalep.csproj -c Release -r win-x64 --self-contained true
```

Yayın sonrasında `Data/catalog.xml`, `Logo.png` ve `Listeler` klasörünün çalışma düzenini kontrol edin. Uygulama kataloğu ve oluşturulan listeleri çalıştırılabilir dosyanın bulunduğu dizin çevresinde tutacak şekilde tasarlanmıştır; klasörün yazılabilir olması gerekir.

## Veri dosyaları

- `Data/catalog.xml`: dağıtım içindeki başlangıç kataloğu
- `catalog.xml`: çalışma sırasında kullanılan katalog
- `Listeler/`: oluşturulan Excel talep listeleri
- `Logo.png`: form çıktılarında kullanılan logo

Dosyaları uygulama klasörü dışına taşımadan önce yedek alın. Güncellemelerden önce katalog ve liste dosyalarını kopyalamak önerilir.

## Lisans

Bu proje [GNU GPL v3](LICENSE.txt) kapsamında yayımlanmıştır.
