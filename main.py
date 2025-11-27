import pandas as pd
import glob
import os
import re

def data_kent_analiz(dosya_yolu):
    """
    Data Kent formatındaki sınav analiz dosyasını okuyup etüt listesi çıkarır.
    """
    try:
        # Önce CSV olarak dene
        df_raw = pd.read_csv(dosya_yolu, header=None, engine='python', encoding='utf-8-sig')
    except:
        try:
            # CSV değilse Excel olarak dene
            df_raw = pd.read_excel(dosya_yolu, header=None, engine='openpyxl')
        except Exception as e:
            print(f"HATA: {dosya_yolu} okunamadı. ({str(e)})")
            return []

    etut_listesi = []
    dosya_adi = os.path.basename(dosya_yolu)

    # 1. ADIM: "ADI VE SOYADI" başlık satırını bul
    baslik_index = -1
    for i, row in df_raw.iterrows():
        satir = " ".join([str(x) for x in row.values if pd.notna(x)])
        if "ADI VE SOYADI" in satir.upper():
            baslik_index = i
            break
            
    if baslik_index == -1:
        print(f"UYARI: {dosya_adi} dosyasında 'ADI VE SOYADI' başlığı bulunamadı.")
        return []

    # Öğrenci tablosunu oluştur
    df_students = df_raw.iloc[baslik_index+1:].copy()
    df_students.columns = df_raw.iloc[baslik_index]
    
    # "ADI VE SOYADI" sütununu bul
    adi_soyadi_col = None
    for col in df_students.columns:
        if 'ADI VE SOYADI' in str(col).upper():
            adi_soyadi_col = col
            break
    
    # Sıra numarası sütununu bul (genelde "SIRA NO" veya ilk sayısal sütun)
    sira_no_col = None
    for col in df_students.columns:
        col_str = str(col).upper()
        if 'SIRA' in col_str and 'NO' in col_str:
            sira_no_col = col
            break
    
    # Sadece gerçek öğrenci satırlarını al
    # 1. Sıra numarası sütununda sayı olmalı VEYA ilk sütunda sayı olmalı
    if sira_no_col:
        sira_sutun = df_students[sira_no_col]
        mask = pd.to_numeric(sira_sutun, errors='coerce').notnull()
    else:
        # Sıra numarası sütunu bulunamadıysa, ilk sütunu kontrol et
        ilk_sutun = df_students.iloc[:, 0]
        mask = pd.to_numeric(ilk_sutun, errors='coerce').notnull()
    
    # 2. "ADI VE SOYADI" sütununda gerçek bir ad olmalı (NaN değil, "Kazanım:" içermemeli)
    if adi_soyadi_col:
        # "ADI VE SOYADI" sütununda geçerli bir değer olmalı
        adi_mask = df_students[adi_soyadi_col].notna()
        # "Kazanım:" içermemeli
        try:
            kazanim_mask = ~df_students[adi_soyadi_col].astype(str).str.contains('Kazanım', case=False, na=False)
            mask = mask & adi_mask & kazanim_mask
        except:
            # Hata olursa sadece NaN kontrolü yap
            mask = mask & adi_mask
    
    df_students = df_students[mask].copy()
    
    # Boş satırları temizle
    df_students = df_students.dropna(how='all')
    
    if df_students.empty:
        print(f"UYARI: {dosya_adi} dosyasında öğrenci verisi bulunamadı.")
        return []

    # 2. ADIM: "SORULARA GÖRE BAŞARI (%)" satırını bul ve sınıf başarı yüzdelerini al
    sinif_basari_yuzdeleri = {}
    basari_satir_index = -1
    
    for i, row in df_raw.iterrows():
        satir = " ".join([str(x) for x in row.values if pd.notna(x)])
        if "SORULARA GÖRE BAŞARI" in satir.upper() or "BAŞARI (%)" in satir.upper():
            basari_satir_index = i
            # Başlık satırındaki soru numaralarını bul
            baslik_row = df_raw.iloc[baslik_index]
            
            # Bu satırdaki yüzde değerlerini al ve başlık satırındaki soru numaralarıyla eşleştir
            for col_idx, val in enumerate(row.values):
                if col_idx < len(baslik_row):
                    baslik_val = baslik_row.iloc[col_idx]
                    baslik_str = str(baslik_val).strip() if pd.notna(baslik_val) else ""
                    
                    # Başlık sütununda soru numarası var mı? (tam sayı veya "4.0" gibi float string'i)
                    baslik_clean = baslik_str.replace('.', '').replace(',', '')
                    if baslik_clean.isdigit():
                        # Soru numarasını tam sayı olarak al (örn: "4.0" -> "4")
                        soru_no = str(int(float(baslik_str))) if '.' in baslik_str else baslik_str
                        # Bu sütundaki yüzde değerini al
                        if pd.notna(val):
                            try:
                                # Direkt sayısal değer olabilir (örn: 68.57) veya string olabilir
                                if isinstance(val, (int, float)):
                                    yuzde_val = float(val)
                                else:
                                    val_str = str(val).strip().replace('%', '').replace(',', '.')
                                    yuzde_match = re.search(r'(\d+\.?\d*)', val_str)
                                    if yuzde_match:
                                        yuzde_val = float(yuzde_match.group(1))
                                    else:
                                        continue
                                
                                # Yüzde değeri 0-100 arasında olmalı
                                if 0 <= yuzde_val <= 100:
                                    sinif_basari_yuzdeleri[soru_no] = yuzde_val
                            except:
                                continue
            break

    # 3. ADIM: "Kazanım:" satırlarından soru-konu eşleştirmesi yap
    kazanim_map = {}
    soru_max_puan = {}
    
    for i, row in df_raw.iterrows():
        satir = " ".join([str(x) for x in row.values if pd.notna(x)])
        if "Kazanım:" in satir or "KAZANIM:" in satir:
            try:
                # Soru numarası genelde ilk sütunda (0. indeks)
                soru_no = None
                kazanim_metni = None
                max_puan = 25  # Varsayılan değer
                
                # İlk sütunda soru numarası var mı?
                if len(row.values) > 0 and pd.notna(row.values[0]):
                    ilk_deger = str(row.values[0]).strip()
                    if ilk_deger.isdigit():
                        soru_no = ilk_deger
                
                # Kazanım metnini bul (genelde 3. sütunda - indeks 2)
                if len(row.values) > 2 and pd.notna(row.values[2]):
                    kazanim_cell = str(row.values[2]).strip()
                    if "Kazanım:" in kazanim_cell or "KAZANIM:" in kazanim_cell:
                        kazanim_metni = kazanim_cell.replace("Kazanım:", "").replace("KAZANIM:", "").strip()
                
                # Maksimum puanı bul (genelde 5. sütunda - indeks 4)
                if len(row.values) > 4 and pd.notna(row.values[4]):
                    puan_cell = str(row.values[4]).strip()
                    if puan_cell.isdigit():
                        try:
                            max_puan = int(puan_cell)
                        except:
                            pass
                
                if soru_no and kazanim_metni:
                    kazanim_map[soru_no] = kazanim_metni
                    soru_max_puan[soru_no] = max_puan
            except Exception as e:
                continue

    # 4. ADIM: Soru sütunlarını bul (sayısal sütun adları)
    soru_sutunlari = []
    for col in df_students.columns:
        col_str = str(col).strip()
        # Sayısal değer veya "4.0" gibi float string'i olabilir
        col_clean = col_str.replace('.', '').replace(',', '')
        if col_clean.isdigit():
            soru_sutunlari.append(col)

    if not soru_sutunlari:
        print(f"UYARI: {dosya_adi} dosyasında soru sütunları bulunamadı.")
        return []

    # 5. ADIM: Kuralları uygula
    for soru_col in soru_sutunlari:
        # Sütun adını temizle (örn: "4.0" -> "4")
        soru_no = str(soru_col).strip().split('.')[0]
        
        # Konu bilgisini al
        konu = kazanim_map.get(soru_no, f"{soru_no}. Soru")
        
        # Maksimum puanı al
        max_p = soru_max_puan.get(soru_no, 25)
        
        # Sınıf başarı yüzdesini al (önce "SORULARA GÖRE BAŞARI (%)" satırından, yoksa hesapla)
        sinif_yuzde = sinif_basari_yuzdeleri.get(soru_no)
        
        if sinif_yuzde is None:
            # Manuel hesapla
            notlar = pd.to_numeric(df_students[soru_col], errors='coerce').dropna()
            if len(notlar) > 0:
                sinif_ort = notlar.mean()
                sinif_yuzde = (sinif_ort / max_p) * 100 if max_p > 0 else 0
            else:
                sinif_yuzde = 0

        # KURAL 1: Sınıf başarısı %35 ve altındaysa TÜM SINIF etüde kalır
        if sinif_yuzde <= 35:
            etut_listesi.append({
                "Dosya": dosya_adi,
                "Öğrenci": "TÜM SINIF",
                "Konu": konu,
                "Sebep": f"Sınıf Ortalaması: %{sinif_yuzde:.1f} (Kritik Altı)"
            })
            continue  # Tüm sınıf yazıldı, bireysel kontrol yapma

        # KURAL 2: Sınıf başarısı iyiyse, bireysel kontrol yap
        limit = max_p * 0.5  # Sorunun tam puanının %50'si
        
        for _, ogrenci in df_students.iterrows():
            try:
                puan = pd.to_numeric(ogrenci[soru_col], errors='coerce')
                if pd.isna(puan):
                    continue
                
                if puan <= limit:
                    # Öğrenci adını al
                    ogrenci_adi = None
                    if adi_soyadi_col:
                        ogrenci_adi_val = ogrenci.get(adi_soyadi_col)
                        if pd.notna(ogrenci_adi_val):
                            ogrenci_adi = str(ogrenci_adi_val).strip()
                            # "Kazanım:" içermemeli
                            if 'Kazanım' in ogrenci_adi or 'KAZANIM' in ogrenci_adi:
                                ogrenci_adi = None
                    
                    # Eğer hala bulamadıysak, tüm sütunları kontrol et
                    if not ogrenci_adi:
                        for col_name in df_students.columns:
                            col_str = str(col_name).upper()
                            if 'ADI' in col_str and 'SOYADI' in col_str:
                                ogrenci_adi_val = ogrenci.get(col_name)
                                if pd.notna(ogrenci_adi_val):
                                    ogrenci_adi_temp = str(ogrenci_adi_val).strip()
                                    if ogrenci_adi_temp and 'Kazanım' not in ogrenci_adi_temp:
                                        ogrenci_adi = ogrenci_adi_temp
                                        break
                    
                    if ogrenci_adi and ogrenci_adi.lower() != 'nan' and ogrenci_adi != '':
                        etut_listesi.append({
                            "Dosya": dosya_adi,
                            "Öğrenci": ogrenci_adi,
                            "Konu": konu,
                            "Sebep": f"Aldığı Puan: {puan:.1f}/{max_p} (≤%50)"
                        })
            except Exception as e:
                continue

    return etut_listesi


# --- ANA PROGRAM ---
if __name__ == "__main__":
    print("=" * 60)
    print("Data Kent Sınav Analizi - Etüt Listesi Oluşturucu")
    print("=" * 60)
    print()
    
    # 'inputs' klasöründeki tüm CSV ve Excel dosyalarını bul
    csv_files = glob.glob("inputs/*.csv")
    xlsx_files = glob.glob("inputs/*.xlsx")
    files = csv_files + xlsx_files
    
    if not files:
        print("[HATA] 'inputs' klasorunde CSV veya Excel dosyasi bulunamadi!")
        print("   Lutfen analiz dosyalarini 'inputs' klasorune koyun.")
    else:
        print(f"[BILGI] {len(files)} dosya bulundu. Isleniyor...\n")
        
        tum_data = []
        basarili_dosya_sayisi = 0
        
        for f in files:
            print(f"  [ISLEM] Isleniyor: {os.path.basename(f)}")
            try:
                sonuc = data_kent_analiz(f)
                if sonuc:
                    tum_data.extend(sonuc)
                    basarili_dosya_sayisi += 1
                    print(f"     [OK] {len(sonuc)} etut kaydi bulundu")
                else:
                    print(f"     [UYARI] Etut kaydi bulunamadi")
            except Exception as e:
                print(f"     [HATA] Hata: {str(e)}")
                import traceback
                traceback.print_exc()
        
        print()
        print("=" * 60)
        
        if tum_data:
            # Excel dosyasına yaz
            df_sonuc = pd.DataFrame(tum_data)
            cikti_dosyasi = "Etut_Listesi.xlsx"
            
            try:
                df_sonuc.to_excel(cikti_dosyasi, index=False, engine='openpyxl')
                print("[BASARILI] ISLEM TAMAM!")
                print(f"   Toplam {len(tum_data)} etut kaydi bulundu")
                print(f"   {basarili_dosya_sayisi} dosya basariyla islendi")
                print(f"   Cikti dosyasi: '{cikti_dosyasi}'")
            except Exception as e:
                print(f"[HATA] Excel dosyasi olusturulamadi: {str(e)}")
                import traceback
                traceback.print_exc()
        else:
            print("[UYARI] Hicbir dosyadan etut kaydi cikarilamadi.")
            print("   Dosya formatlarini kontrol edin.")
        
        print("=" * 60)
