# 📚 Data Kent - Etüt Listesi Oluşturucu

Öğretmenler için geliştirilmiş, modern ve kullanıcı dostu bir program. Data Kent sınav analiz dosyalarınızı otomatik olarak işleyerek etüt listesi oluşturur.

## ✨ Özellikler

- 🎨 **Modern Arayüz**: 2024 standartlarında, şık ve kullanıcı dostu tasarım
- 📁 **Kolay Dosya Yönetimi**: Birden fazla dosyayı tek seferde ekleyebilme ve kaldırabilme
- 📊 **Otomatik Analiz**: Sınıf başarısına göre akıllı etüt belirleme
- 📋 **Detaylı Rapor**: Excel formatında profesyonel çıktı
- ⚡ **Hızlı İşlem**: Toplu dosya işleme desteği
- 🎯 **Akıllı Gruplama**: Etüt grupları otomatik olarak maksimum 4 kişi olacak şekilde oluşturulur

## 📋 Sistem Gereksinimleri

- **İşletim Sistemi**: Windows, macOS veya Linux
- **Python**: 3.7 veya üzeri
- **Gerekli Kütüphaneler**: 
  - pandas
  - openpyxl
  - tkinter (genelde Python ile birlikte gelir)

## 🚀 Kurulum

### Adım 1: Python'u Kontrol Edin
Terminal/Command Prompt'ta şu komutu çalıştırın:
```bash
python --version
```
Eğer Python yüklü değilse, [python.org](https://www.python.org/downloads/) adresinden indirip yükleyin.

### Adım 2: Gerekli Kütüphaneleri Yükleyin
Proje klasöründe terminal açın ve şu komutu çalıştırın:
```bash
pip install -r requirements.txt
```

### Adım 3: Programı Başlatın
```bash
python etut_programi.py
```

## 📖 Kullanım Kılavuzu

### 1️⃣ Programı Başlatma
- `etut_programi.py` dosyasını çift tıklayarak veya terminalden çalıştırarak başlatın
- Modern grafik arayüz otomatik olarak açılacaktır

### 2️⃣ Dosya Ekleme
Programa dosya eklemenin iki yolu vardır:

**Yöntem 1: Tek Tek Dosya Ekleme**
- **"➕ Dosya Ekle"** butonuna tıklayın
- Açılan pencereden sınav analiz dosyalarınızı seçin (Excel veya CSV formatında)
- Birden fazla dosyayı aynı anda seçebilirsiniz (Ctrl tuşu ile)
- Dosya formatı için `input_format/input_format.xlsx` dosyasını referans alabilirsiniz

**Yöntem 2: Klasör Ekleme**
- **"📂 Klasör Ekle"** butonuna tıklayın
- Tüm klasörü seçin, program otomatik olarak uygun dosyaları bulacaktır

### 3️⃣ Dosya Yönetimi
- Eklenen dosyalar listede görünecektir
- Birden fazla dosyayı seçmek için: **Ctrl** tuşuna basılı tutarak dosyalara tıklayın
- Seçili dosyaları kaldırmak için: **"🗑️ Seçiliyi Kaldır"** butonuna tıklayın

### 4️⃣ Etüt Listesi Oluşturma
- Tüm dosyalarınızı ekledikten sonra **"🚀 Etüt Listesini Oluştur"** butonuna tıklayın
- Program dosyaları işlemeye başlayacak ve ilerleme çubuğu gösterecektir
- İşlem tamamlandığında, çıktı dosyasını nereye kaydetmek istediğiniz sorulacaktır
- Konum seçtikten sonra Excel dosyası oluşturulacaktır

### 5️⃣ Sonuçları İnceleme
Oluşturulan Excel dosyasında şu bilgiler yer alır:
- **Dosya**: İşlenen sınav analiz dosyasının adı
- **Soru**: Soru numarası
- **Kazanım**: Sorunun konusu/öğrenme kazanımı
- **Etüt Grubu**: Grup numarası
- **Öğrenciler**: Etüde katılacak öğrenciler (veya "TÜM SINIF")
- **Sebep**: Etüde kalma nedeni ve öğrenci puanları
- **Etüt Süresi**: Önerilen süre (20 veya 40 dakika)
- **Soru Sayısı**: Çalışılacak soru sayısı (5 soru)
- **Etüt Tipi**: Bireysel Etüt veya Sınıf Etütü

## 📊 Etüt Belirleme Kuralları

Program, her soru için şu kurallara göre etüt belirler:

### 📌 KURAL 1: Genel Başarı (Sınıf Etütü)
- Bir soruda **sınıf başarısı %35 ve altındaysa**
- **TÜM SINIF** o konudan etüde kalır
- **Etüt Süresi**: 40 dakika
- **Soru Sayısı**: 5 soru
- **Etüt Tipi**: Sınıf Etütü

### 📌 KURAL 2: Bireysel Başarı (Bireysel Etüt)
- Sınıf başarısı %35'in **üstündeyse**
- Sorunun tam puanının **%50'si ve altında** puan alan öğrenciler etüde kalır
- **Etüt Süresi**: 20 dakika
- **Soru Sayısı**: 5 soru
- **Etüt Tipi**: Bireysel Etüt
- Öğrenciler otomatik olarak **maksimum 4 kişilik gruplara** ayrılır

## ⚠️ Önemli Notlar

- Program belirli bir formattaki sınav analiz dosyalarını bekler
- Dosya formatı için `input_format/input_format.xlsx` dosyasını inceleyin
- Dosyalarda **"ADI VE SOYADI"** başlığı bulunmalıdır
- **"SORULARA GÖRE BAŞARI (%)"** satırından sınıf başarısı okunur
- **"Soruların ilgili olduğu konular, kazanımlar..."** başlığı altından soru-konu eşleştirmesi yapılır
- Eğer bir kazanımda sınıf başarısı **%0** görünüyorsa, Excel dosyasında hata olabilir (program bunu uyarı olarak gösterir)

## 🐛 Sorun Giderme

### Program açılmıyor
- Python'un doğru yüklendiğinden emin olun: `python --version`
- Gerekli kütüphaneleri yükleyin: `pip install -r requirements.txt`
- Python'un PATH'e eklendiğinden emin olun

### Dosya okunamıyor
- Dosya formatının doğru olduğundan emin olun (Excel veya CSV)
- Excel dosyası bozuk olabilir, başka bir dosya deneyin
- Dosya başka bir programda açık olmamalı

### Etüt kaydı bulunamıyor
- Dosya formatını kontrol edin (`input_format/input_format.xlsx` dosyasını referans alın)
- **"ADI VE SOYADI"** başlığının dosyada olduğundan emin olun
- Dosyanın tam formatını kontrol edin

### Excel dosyası kaydedilemiyor
- Dosya başka bir programda (ör. Excel) açık olmamalı
- Kayıt konumunda yazma izniniz olduğundan emin olun
- Farklı bir konum seçmeyi deneyin

## 📁 Proje Yapısı

```
etut/
├── etut_programi.py      # Ana GUI programı (kullanılacak)
├── main.py               # Komut satırı versiyonu (yedek)
├── requirements.txt      # Gerekli Python kütüphaneleri
├── README.md            # Bu kullanım kılavuzu
├── .gitignore           # Git yapılandırması
└── input_format/        # Örnek dosya formatı
    └── input_format.xlsx # Örnek sınav analiz dosyası formatı
```

## 📄 Dosya Formatı

Program, `input_format/input_format.xlsx` dosyasında gösterilen formattaki sınav analiz dosyalarını bekler. Bu dosyayı referans alarak kendi dosyalarınızı hazırlayabilirsiniz.

## 💡 İpuçları

- Birden fazla sınav analiz dosyasını aynı anda işleyebilirsiniz
- Etüt grupları otomatik olarak oluşturulur, manuel düzenleme gerekmez
- Excel çıktısı renkli ve okunabilir formatta hazırlanır
- Program, hatalı dosyaları atlayarak diğer dosyaları işlemeye devam eder

## 📞 Destek

Sorun yaşarsanız:
1. Bu kılavuzdaki "Sorun Giderme" bölümünü kontrol edin
2. Dosya formatınızın `input_format/input_format.xlsx` dosyasındaki formata uygun olduğundan emin olun
3. Python ve kütüphane versiyonlarınızı kontrol edin

---

**Başarılar dileriz! 📚✨**

