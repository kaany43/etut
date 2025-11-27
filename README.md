# 📚 Etüt Listesi Oluşturucu

Öğretmenler için geliştirilmiş, Data Kent sınav analiz dosyalarını işleyerek otomatik etüt listesi oluşturan kullanıcı dostu program.

## 🎯 Özellikler

- ✅ **Kolay Kullanım**: Grafik arayüz ile dosya seçimi ve işlem
- ✅ **Toplu İşleme**: Birden fazla dosyayı aynı anda işleyebilme
- ✅ **Otomatik Analiz**: Sınıf başarısına göre otomatik etüt belirleme
- ✅ **Excel Çıktı**: Sonuçları Excel formatında kaydetme

## 📋 Gereksinimler

- Python 3.7 veya üzeri
- Gerekli kütüphaneler:
  - pandas
  - openpyxl
  - tkinter (genelde Python ile birlikte gelir)

## 🚀 Kurulum

1. Gerekli kütüphaneleri yükleyin:
```bash
pip install -r requirements.txt
```

2. Programı çalıştırın:
```bash
python etut_programi.py
```

## 📖 Kullanım

### 1. Programı Başlatma
- `etut_programi.py` dosyasını çalıştırın
- Grafik arayüz açılacaktır

### 2. Dosya Ekleme
- **Dosya Ekle** butonuna tıklayın
- Data Kent sınav analiz dosyalarınızı seçin (Excel veya CSV)
- Veya **Klasör Ekle** butonu ile tüm klasörü ekleyebilirsiniz

### 3. Etüt Listesi Oluşturma
- **Etüt Listesini Oluştur** butonuna tıklayın
- Program dosyaları işleyecek ve ilerleme gösterecektir
- İşlem tamamlandığında kayıt konumu seçmeniz istenecektir

### 4. Sonuç
- Excel dosyası oluşturulacaktır
- Dosyada şu bilgiler yer alır:
  - **Dosya**: Analiz edilen dosya adı
  - **Öğrenci**: Öğrenci adı veya "TÜM SINIF"
  - **Konu**: Sorunun konusu (Kazanım)
  - **Sebep**: Etüde kalma nedeni

## 📊 Etüt Belirleme Kuralları

### KURAL 1: Genel Başarı
- Bir soruda sınıf başarısı **%35 ve altındaysa**
- **TÜM SINIF** o konudan etüde kalır

### KURAL 2: Bireysel Başarı
- Sınıf başarısı %35'in üstündeyse
- Sorunun tam puanının **%50'si ve altında** puan alan öğrenciler etüde kalır

## 📁 Dosya Yapısı

```
etut_analiz/
├── etut_programi.py      # Ana GUI programı
├── main.py               # Komut satırı versiyonu
├── requirements.txt      # Gerekli kütüphaneler
└── README.md            # Bu dosya
```

## ⚠️ Notlar

- Program Data Kent formatındaki dosyaları bekler
- Dosyalarda "ADI VE SOYADI" başlığı olmalıdır
- "SORULARA GÖRE BAŞARI (%)" satırından sınıf başarısı okunur
- "Kazanım:" satırlarından soru-konu eşleştirmesi yapılır

## 🐛 Sorun Giderme

**Program açılmıyor:**
- Python'un yüklü olduğundan emin olun
- `pip install -r requirements.txt` komutunu çalıştırın

**Dosya okunamıyor:**
- Dosya formatının doğru olduğundan emin olun
- Excel dosyası bozuk olabilir, başka bir dosya deneyin

**Etüt kaydı bulunamıyor:**
- Dosya formatını kontrol edin
- "ADI VE SOYADI" başlığının dosyada olduğundan emin olun

---

**Kolay gelsin! 📚✨**

