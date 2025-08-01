# JSON to Excel Converter

Bu proje, JSON dosyalarını Excel (.xlsx) formatına dönüştüren bir .NET Console uygulamasıdır.

## Özellikler

- JSON verilerini Excel dosyalarına dönüştürme
- RowIndex alanına göre satırları doğru sırada düzenleme
- Otomatik kolon boyutlandırma
- Tarih ve sayı formatlarını koruma
- Başlık satırlarını vurgulama

## Gereksinimler

- .NET 9.0 veya üstü
- ClosedXML (MIT lisanslı, ücretsiz)
- Newtonsoft.Json

## Kurulum

1. Projeyi klonlayın veya indirin
2. Proje klasörüne gidin
3. Bağımlılıkları yükleyin:

   ```bash
   dotnet restore
   ```

## Kullanım

### Temel Kullanım

```bash
dotnet run
```

Bu komut, varsayılan olarak `docs/FORD-KS-SingleSheet.json` dosyasını okur ve `output.xlsx` dosyasını oluşturur.

### Özel Dosya Yolları

```bash
dotnet run "girdi.json" "cikti.xlsx"
```

### Derleme ve Çalıştırma

```bash
# Projeyi derle
dotnet build

# Uygulamayı çalıştır
dotnet run [json-dosyası] [excel-dosyası]
```

## JSON Formatı

Uygulama aşağıdaki JSON formatını destekler:

```json
{
    "SHEET_NAME": [
        {
            "Column0": "Başlık1",
            "Column1": "Başlık2",
            "RowIndex": 0 // Başlık satırı, RowIndex=0 olmalı
        },
        {
            "Column0": "Veri1",
            "Column1": "Veri2", 
            "RowIndex": 1 // Veri satırı
        }
    ]
}
```

## Önemli Notlar


- JSON'daki `RowIndex` alanı, satırların Excel'deki sırasını belirler
- **Başlık satırı, RowIndex=0 olan satırdan alınır** (JSON dizisinde nerede olursa olsun)
- Veri satırları, RowIndex değerine göre sıralanır
- Excel'de geçersiz sheet isimleri otomatik olarak düzeltilir
- Tarih formatları otomatik olarak tanınır ve formatlanır

## Lisans

Bu proje MIT lisansı altında dağıtılmaktadır.
