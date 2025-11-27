import pandas as pd
import glob
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import threading

class EtutListesiProgrami:
    def __init__(self, root):
        self.root = root
        self.root.title("Etüt Listesi Oluşturucu")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # Modern renkler - Daha canlı ve profesyonel
        self.colors = {
            'primary': '#4A90E2',
            'success': '#50C878',
            'danger': '#FF6B6B',
            'warning': '#FFA500',
            'dark': '#2C3E50',
            'light': '#F5F7FA',
            'white': '#FFFFFF',
            'accent': '#6C5CE7',
            'info': '#00D2FF'
        }
        
        # Seçilen dosyalar
        self.secili_dosyalar = []
        
        # Arayüzü oluştur
        self.arayuz_olustur()
        
    def arayuz_olustur(self):
        # Başlık - Modern gradient efekti
        baslik_frame = tk.Frame(self.root, bg=self.colors['dark'], height=100)
        baslik_frame.pack(fill=tk.X)
        baslik_frame.pack_propagate(False)
        
        # Ana başlık
        baslik_label = tk.Label(
            baslik_frame,
            text="📚 Data Kent - Etüt Listesi Oluşturucu",
            font=("Segoe UI", 22, "bold"),
            bg=self.colors['dark'],
            fg="white"
        )
        baslik_label.pack(pady=20)
        
        # Alt başlık
        alt_baslik = tk.Label(
            baslik_frame,
            text="Sınav Analizlerinden Otomatik Etüt Planı Oluşturma Sistemi",
            font=("Segoe UI", 10),
            bg=self.colors['dark'],
            fg="#BDC3C7"
        )
        alt_baslik.pack(pady=(0, 20))
        
        # Ana içerik - Modern card tasarımı
        icerik_frame = tk.Frame(self.root, bg=self.colors['light'], padx=25, pady=25)
        icerik_frame.pack(fill=tk.BOTH, expand=True)
        
        # Açıklama kartı
        aciklama_frame = tk.Frame(
            icerik_frame,
            bg=self.colors['white'],
            relief=tk.FLAT,
            padx=20,
            pady=15
        )
        aciklama_frame.pack(fill=tk.X, pady=(0, 20))
        
        aciklama = tk.Label(
            aciklama_frame,
            text="📋 Sınav analiz dosyalarınızı seçin ve otomatik etüt listesi oluşturun.",
            font=("Segoe UI", 11),
            fg="#34495e",
            bg=self.colors['white']
        )
        aciklama.pack(pady=5)
        
        # Bilgi notu
        bilgi_notu = tk.Label(
            aciklama_frame,
            text="💡 İpucu: Birden fazla dosya seçebilir, klasör ekleyebilirsiniz. Etüt grupları maksimum 4 kişi olacak şekilde otomatik oluşturulur.",
            font=("Segoe UI", 9),
            fg="#7f8c8d",
            bg=self.colors['white'],
            wraplength=800,
            justify=tk.LEFT
        )
        bilgi_notu.pack(pady=(5, 0))
        
        # Dosya seçme bölümü - Modern card
        dosya_frame = tk.Frame(
            icerik_frame,
            bg=self.colors['white'],
            relief=tk.FLAT,
            padx=20,
            pady=20
        )
        dosya_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Başlık
        baslik_dosya = tk.Label(
            dosya_frame,
            text="📁 Dosya Yönetimi",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['white'],
            fg=self.colors['dark']
        )
        baslik_dosya.pack(anchor=tk.W, pady=(0, 15))
        
        # Dosya listesi - Modern liste kutusu
        liste_container = tk.Frame(dosya_frame, bg=self.colors['white'])
        liste_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Scrollbar
        scrollbar = tk.Scrollbar(liste_container, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Liste kutusu - Modern stil
        listbox_frame = tk.Frame(liste_container, bg="#f8f9fa", relief=tk.SOLID, bd=1)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.dosya_listbox = tk.Listbox(
            listbox_frame,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 10),
            height=10,
            bg="#f8f9fa",
            fg="#333",
            selectbackground=self.colors['primary'],
            selectforeground="white",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0
        )
        self.dosya_listbox.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        scrollbar.config(command=self.dosya_listbox.yview)
        
        # Butonlar - Modern buton tasarımı
        buton_frame = tk.Frame(dosya_frame, bg=self.colors['white'])
        buton_frame.pack(fill=tk.X)
        
        # Buton stili fonksiyonu
        def modern_buton(parent, text, command, bg_color, fg_color="white", width=15):
            btn = tk.Button(
                parent,
                text=text,
                command=command,
                font=("Segoe UI", 10, "bold"),
                bg=bg_color,
                fg=fg_color,
                padx=20,
                pady=10,
                cursor="hand2",
                relief=tk.FLAT,
                bd=0,
                activebackground=bg_color,
                activeforeground=fg_color,
                width=width
            )
            return btn
        
        dosya_ekle_btn = modern_buton(
            buton_frame,
            "➕ Dosya Ekle",
            self.dosya_ekle,
            self.colors['primary'],
            width=18
        )
        dosya_ekle_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        klasor_ekle_btn = modern_buton(
            buton_frame,
            "📂 Klasör Ekle",
            self.klasor_ekle,
            self.colors['success'],
            width=18
        )
        klasor_ekle_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        dosya_sil_btn = modern_buton(
            buton_frame,
            "🗑️ Seçiliyi Kaldır",
            self.dosya_sil,
            self.colors['danger'],
            width=18
        )
        dosya_sil_btn.pack(side=tk.LEFT)
        
        # İşlem butonu - Büyük ve dikkat çekici
        islem_frame = tk.Frame(icerik_frame, bg=self.colors['light'])
        islem_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.islem_btn = tk.Button(
            islem_frame,
            text="🚀 Etüt Listesini Oluştur",
            command=self.islem_baslat,
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['success'],
            fg="white",
            padx=50,
            pady=18,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#45B878",
            activeforeground="white"
        )
        self.islem_btn.pack()
        
        # Bilgi kutusu
        bilgi_kutusu = tk.Frame(
            islem_frame,
            bg="#E8F5E9",
            relief=tk.FLAT,
            padx=15,
            pady=10
        )
        bilgi_kutusu.pack(fill=tk.X, pady=(15, 0))
        
        bilgi_text = tk.Label(
            bilgi_kutusu,
            text="ℹ️ Etüt Süreleri: Bireysel Etüt → 20 dakika (5 soru) | Sınıf Etütü → 40 dakika (5 soru)",
            font=("Segoe UI", 9),
            fg="#2E7D32",
            bg="#E8F5E9"
        )
        bilgi_text.pack()
        
        # İlerleme çubuğu - Modern stil
        progress_frame = tk.Frame(islem_frame, bg=self.colors['light'])
        progress_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=500,
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress.pack()
        
        # Durum etiketi
        self.durum_label = tk.Label(
            progress_frame,
            text="Hazır",
            font=("Segoe UI", 9),
            fg="#7f8c8d",
            bg=self.colors['light']
        )
        self.durum_label.pack(pady=(10, 0))
        
        # Modern progressbar stili
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Modern.Horizontal.TProgressbar",
                       background=self.colors['success'],
                       troughcolor=self.colors['light'],
                       borderwidth=0,
                       lightcolor=self.colors['success'],
                       darkcolor=self.colors['success'])
        
    def dosya_ekle(self):
        dosyalar = filedialog.askopenfilenames(
            title="Sınav Analiz Dosyalarını Seçin",
            filetypes=[
                ("Excel/CSV Dosyaları", "*.xlsx *.xls *.csv"),
                ("Excel Dosyaları", "*.xlsx *.xls"),
                ("CSV Dosyaları", "*.csv"),
                ("Tüm Dosyalar", "*.*")
            ]
        )
        
        for dosya in dosyalar:
            if dosya not in self.secili_dosyalar:
                self.secili_dosyalar.append(dosya)
                self.dosya_listbox.insert(tk.END, f"📄 {os.path.basename(dosya)}")
        
        self.durum_guncelle(f"✅ {len(self.secili_dosyalar)} dosya seçildi")
        
    def klasor_ekle(self):
        klasor = filedialog.askdirectory(title="Klasör Seçin")
        if klasor:
            dosyalar = []
            for ext in ['*.xlsx', '*.xls', '*.csv']:
                dosyalar.extend(glob.glob(os.path.join(klasor, ext)))
            
            eklenen = 0
            for dosya in dosyalar:
                if dosya not in self.secili_dosyalar:
                    self.secili_dosyalar.append(dosya)
                    self.dosya_listbox.insert(tk.END, f"📄 {os.path.basename(dosya)}")
                    eklenen += 1
            
            if eklenen > 0:
                self.durum_guncelle(f"✅ {eklenen} dosya eklendi")
            else:
                messagebox.showinfo("Bilgi", "Klasörde uygun dosya bulunamadı.")
        
    def dosya_sil(self):
        secili_indeksler = self.dosya_listbox.curselection()
        if not secili_indeksler:
            messagebox.showwarning("Uyarı", "Lütfen silmek için bir dosya seçin.")
            return
        
        # Tersten sil
        for indeks in reversed(secili_indeksler):
            self.dosya_listbox.delete(indeks)
            del self.secili_dosyalar[indeks]
        
        self.durum_guncelle(f"📋 {len(self.secili_dosyalar)} dosya kaldı")
        
    def durum_guncelle(self, mesaj):
        self.durum_label.config(text=mesaj)
        self.root.update()
        
    def islem_baslat(self):
        if not self.secili_dosyalar:
            messagebox.showwarning("Uyarı", "Lütfen en az bir dosya seçin!")
            return
        
        # Butonu devre dışı bırak
        self.islem_btn.config(state=tk.DISABLED, text="⏳ İşleniyor...")
        self.progress.start()
        self.durum_guncelle("İşlem başlatılıyor...")
        
        # Arka planda çalıştır
        thread = threading.Thread(target=self.islem_yap)
        thread.daemon = True
        thread.start()
        
    def islem_yap(self):
        try:
            tum_data = []
            basarili_dosya_sayisi = 0
            
            for i, dosya_yolu in enumerate(self.secili_dosyalar):
                dosya_adi = os.path.basename(dosya_yolu)
                self.durum_guncelle(f"📊 İşleniyor: {dosya_adi} ({i+1}/{len(self.secili_dosyalar)})")
                
                try:
                    sonuc = self.data_kent_analiz(dosya_yolu)
                    if sonuc:
                        tum_data.extend(sonuc)
                        basarili_dosya_sayisi += 1
                        
                        # Dosyalar arası boş satır ekle (son dosya değilse)
                        if i < len(self.secili_dosyalar) - 1:
                            tum_data.append({
                                "Dosya": "",
                                "Soru": "",
                                "Kazanım": "",
                                "Etüt Grubu": "",
                                "Öğrenciler": "",
                                "Sebep": "",
                                "Etüt Süresi": "",
                                "Soru Sayısı": "",
                                "Etüt Tipi": ""
                            })
                except Exception as e:
                    print(f"Hata: {dosya_yolu} - {str(e)}")
            
            if tum_data:
                # Çıktı dosyasını kaydet
                cikti_dosyasi = filedialog.asksaveasfilename(
                    title="Etüt Listesini Kaydet",
                    defaultextension=".xlsx",
                    filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
                )
                
                if cikti_dosyasi:
                    df_sonuc = pd.DataFrame(tum_data)
                    
                    # Excel'e yazarken formatlamayı iyileştir
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter
                    
                    with pd.ExcelWriter(cikti_dosyasi, engine='openpyxl') as writer:
                        df_sonuc.to_excel(writer, index=False, sheet_name='Etüt Listesi')
                        
                        worksheet = writer.sheets['Etüt Listesi']
                        
                        # Stil tanımlamaları
                        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        header_font = Font(bold=True, color="FFFFFF", size=11)
                        border_style = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        
                        # Başlık satırını formatla
                        for cell in worksheet[1]:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = center_align
                            cell.border = border_style
                        
                        # Sütun genişliklerini ayarla
                        column_widths = {
                            'A': 35,  # Dosya
                            'B': 8,   # Soru
                            'C': 55,  # Kazanım
                            'D': 12,  # Etüt Grubu
                            'E': 50,  # Öğrenciler
                            'F': 80,  # Sebep (detaylı)
                            'G': 15,  # Etüt Süresi
                            'H': 12,  # Soru Sayısı
                            'I': 18   # Etüt Tipi
                        }
                        
                        for col, width in column_widths.items():
                            worksheet.column_dimensions[col].width = width
                        
                        # Satırları formatla
                        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
                            # Boş satır kontrolü
                            is_empty = all(cell.value is None or str(cell.value).strip() == '' for cell in row)
                            
                            for cell in row:
                                cell.border = border_style
                                
                                if is_empty:
                                    # Boş satırları farklı renkle işaretle
                                    cell.fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
                                else:
                                    # Normal satırlar
                                    if cell.column == 1:  # Dosya sütunu
                                        cell.alignment = left_align
                                        cell.font = Font(bold=True, size=10)
                                    elif cell.column == 2:  # Soru sütunu
                                        cell.alignment = center_align
                                        cell.font = Font(bold=True, size=10)
                                    elif cell.column == 3:  # Kazanım sütunu
                                        cell.alignment = left_align
                                        cell.font = Font(size=9)
                                    elif cell.column == 4:  # Etüt Grubu
                                        cell.alignment = center_align
                                        cell.font = Font(bold=True, size=10)
                                    elif cell.column == 5:  # Öğrenciler
                                        cell.alignment = left_align
                                        cell.font = Font(size=9)
                                    elif cell.column == 6:  # Sebep
                                        cell.alignment = left_align
                                        cell.font = Font(size=8)
                                    elif cell.column in [7, 8, 9]:  # Etüt Süresi, Soru Sayısı, Etüt Tipi
                                        cell.alignment = center_align
                                        cell.font = Font(size=9)
                                    
                                    # TÜM SINIF satırlarını vurgula
                                    if len(row) > 4 and row[4].value == "TÜM SINIF":
                                        for cell in row:
                                            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                                    
                                    # Bireysel etüt satırlarını hafif mavi yap
                                    elif len(row) > 8 and row[8].value == "Bireysel Etüt":
                                        for cell in row:
                                            try:
                                                # Eğer renk yoksa veya varsayılan renkse
                                                if not hasattr(cell.fill, 'start_color') or cell.fill.start_color.index == "00000000":
                                                    cell.fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
                                            except:
                                                cell.fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
                        
                        # Satır yüksekliklerini ayarla
                        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row):
                            worksheet.row_dimensions[row[0].row].height = 25
                    
                    self.progress.stop()
                    self.islem_btn.config(state=tk.NORMAL, text="🚀 Etüt Listesini Oluştur")
                    self.durum_guncelle("✅ İşlem tamamlandı!")
                    
                    # Toplam etüt grubu sayısını hesapla
                    toplam_grup = len([x for x in tum_data if x.get("Etüt Grubu") and str(x.get("Etüt Grubu")) != ""])
                    
                    messagebox.showinfo(
                        "Başarılı",
                        f"✅ İşlem tamamlandı!\n\n"
                        f"📊 Toplam {toplam_grup} etüt grubu oluşturuldu\n"
                        f"📁 {basarili_dosya_sayisi} dosya başarıyla işlendi\n"
                        f"💾 Dosya kaydedildi: {os.path.basename(cikti_dosyasi)}"
                    )
                else:
                    self.progress.stop()
                    self.islem_btn.config(state=tk.NORMAL, text="🚀 Etüt Listesini Oluştur")
                    self.durum_guncelle("İptal edildi")
            else:
                self.progress.stop()
                self.islem_btn.config(state=tk.NORMAL, text="🚀 Etüt Listesini Oluştur")
                self.durum_guncelle("Etüt kaydı bulunamadı")
                messagebox.showwarning(
                    "Uyarı",
                    "Hiçbir dosyadan etüt kaydı çıkarılamadı.\nDosya formatlarını kontrol edin."
                )
                
        except Exception as e:
            self.progress.stop()
            self.islem_btn.config(state=tk.NORMAL, text="🚀 Etüt Listesini Oluştur")
            self.durum_guncelle("❌ Hata oluştu!")
            messagebox.showerror("Hata", f"Bir hata oluştu:\n{str(e)}")
    
    def data_kent_analiz(self, dosya_yolu):
        """
        Data Kent formatındaki sınav analiz dosyasını okuyup etüt listesi çıkarır.
        Geliştirilmiş: Kazanım ve puan okuma daha esnek ve doğru.
        """
        try:
            df_raw = pd.read_csv(dosya_yolu, header=None, engine='python', encoding='utf-8-sig')
        except:
            try:
                df_raw = pd.read_excel(dosya_yolu, header=None, engine='openpyxl')
            except:
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
        
        # Sıra numarası sütununu bul
        sira_no_col = None
        for col in df_students.columns:
            col_str = str(col).upper()
            if 'SIRA' in col_str and 'NO' in col_str:
                sira_no_col = col
                break
        
        # Sadece gerçek öğrenci satırlarını al
        if sira_no_col:
            sira_sutun = df_students[sira_no_col]
            mask = pd.to_numeric(sira_sutun, errors='coerce').notnull()
        else:
            ilk_sutun = df_students.iloc[:, 0]
            mask = pd.to_numeric(ilk_sutun, errors='coerce').notnull()
        
        if adi_soyadi_col:
            adi_mask = df_students[adi_soyadi_col].notna()
            try:
                kazanim_mask = ~df_students[adi_soyadi_col].astype(str).str.contains('Kazanım', case=False, na=False)
                mask = mask & adi_mask & kazanim_mask
            except:
                mask = mask & adi_mask
        
        df_students = df_students[mask].copy()
        df_students = df_students.dropna(how='all')
        
        if df_students.empty:
            return []

        # 2. ADIM: "SORULARA GÖRE BAŞARI (%)" satırını bul
        sinif_basari_yuzdeleri = {}
        
        for i, row in df_raw.iterrows():
            satir = " ".join([str(x) for x in row.values if pd.notna(x)])
            if "SORULARA GÖRE BAŞARI" in satir.upper() or "BAŞARI (%)" in satir.upper():
                baslik_row = df_raw.iloc[baslik_index]
                
                for col_idx, val in enumerate(row.values):
                    if col_idx < len(baslik_row):
                        baslik_val = baslik_row.iloc[col_idx]
                        baslik_str = str(baslik_val).strip() if pd.notna(baslik_val) else ""
                        baslik_clean = baslik_str.replace('.', '').replace(',', '')
                        if baslik_clean.isdigit():
                            soru_no = str(int(float(baslik_str))) if '.' in baslik_str else baslik_str
                            if pd.notna(val):
                                try:
                                    if isinstance(val, (int, float)):
                                        yuzde_val = float(val)
                                    else:
                                        val_str = str(val).strip().replace('%', '').replace(',', '.')
                                        yuzde_match = re.search(r'(\d+\.?\d*)', val_str)
                                        if yuzde_match:
                                            yuzde_val = float(yuzde_match.group(1))
                                        else:
                                            continue
                                    if 0 <= yuzde_val <= 100:
                                        sinif_basari_yuzdeleri[soru_no] = yuzde_val
                                except:
                                    continue
                break

        # 3. ADIM: "Soruların ilgili olduğu konular, kazanımlar..." başlığını bul ve altındaki satırları işle
        kazanim_map = {}
        soru_max_puan = {}
        
        # Başlık satırını bul: "Soruların ilgili olduğu konular, kazanımlar veya alt öğrenme alanları"
        kazanim_baslik_index = None
        puan_sutun_index = None
        
        for i, row in df_raw.iterrows():
            satir = " ".join([str(x) for x in row.values if pd.notna(x)])
            satir_upper = satir.upper()
            
            # Başlık satırını bul
            if "SORULAR" in satir_upper and ("KAZANIM" in satir_upper or "KONU" in satir_upper or "ÖĞRENME ALANI" in satir_upper):
                kazanim_baslik_index = i
                
                # "Puan" sütununu bul
                for idx, val in enumerate(row.values):
                    if pd.notna(val):
                        val_str = str(val).upper()
                        if "PUAN" in val_str:
                            puan_sutun_index = idx
                            break
                break
        
        # Başlık bulunduysa, altındaki satırları işle
        if kazanim_baslik_index is not None:
            # Başlık satırındaki sütun yapısını anla
            baslik_row = df_raw.iloc[kazanim_baslik_index]
            
            # Başlık satırındaki soru numarası sütununu bul (genelde ilk sütun veya başlık satırındaki sayısal sütunlar)
            soru_no_sutun_index = None
            for idx, val in enumerate(baslik_row.values):
                if pd.notna(val):
                    val_str = str(val).strip()
                    # Başlık satırında sayısal değer varsa, bu soru numarası sütunu olabilir
                    if val_str.replace('.', '').isdigit():
                        soru_no_sutun_index = idx
                        break
            
            # Eğer bulunamadıysa, ilk sütunu varsay
            if soru_no_sutun_index is None:
                soru_no_sutun_index = 0
            
            # Kazanım metni sütununu bul (başlık satırında "Soruların ilgili olduğu konular, kazanımlar..." yazan sütun)
            kazanim_metin_sutun_index = None
            for idx, val in enumerate(baslik_row.values):
                if pd.notna(val):
                    val_str = str(val).upper()
                    # "SORULAR" ve ("KONU" veya "KAZANIM" veya "ÖĞRENME") içeren sütun
                    if "SORULAR" in val_str and ("KONU" in val_str or "KAZANIM" in val_str or "ÖĞRENME" in val_str):
                        kazanim_metin_sutun_index = idx
                        break
                    # Veya sadece "KONU", "KAZANIM", "ÖĞRENME" içeren sütun (başlık satırında)
                    elif ("KONU" in val_str or "KAZANIM" in val_str or "ÖĞRENME" in val_str) and "PUAN" not in val_str:
                        if kazanim_metin_sutun_index is None:  # İlk bulduğunu al
                            kazanim_metin_sutun_index = idx
            
            # Eğer bulunamadıysa, soru numarası sütununun yanındaki sütunu dene
            if kazanim_metin_sutun_index is None:
                # Soru numarası sütununun yanındaki 3 sütunu kontrol et
                for offset in [1, 2, 3]:
                    test_idx = soru_no_sutun_index + offset
                    if test_idx < len(baslik_row.values):
                        test_val = baslik_row.iloc[test_idx]
                        if pd.notna(test_val):
                            test_str = str(test_val).upper()
                            # "PUAN" içermiyorsa ve boş değilse
                            if "PUAN" not in test_str and len(test_str) > 3:
                                kazanim_metin_sutun_index = test_idx
                                break
            
            # Başlık satırının altındaki satırları işle
            for i in range(kazanim_baslik_index + 1, len(df_raw)):
                row = df_raw.iloc[i]
                
                # Boş satır kontrolü - eğer satır tamamen boşsa veya sadece NaN varsa dur
                if row.isna().all():
                    continue
                
                try:
                    soru_no = None
                    kazanim_metni = None
                    max_puan = None
                    
                    # 1. SORU NUMARASINI BUL
                    if soru_no_sutun_index < len(row.values):
                        soru_val = row.iloc[soru_no_sutun_index]
                        if pd.notna(soru_val):
                            soru_str = str(soru_val).strip()
                            # Sayısal değer mi kontrol et
                            if soru_str.replace('.', '').isdigit():
                                soru_no = str(int(float(soru_str))) if '.' in soru_str else soru_str
                                # Mantıklı bir soru numarası mı (1-100)
                                if not (1 <= int(soru_no) <= 100):
                                    soru_no = None
                    
                    # Eğer hala bulunamadıysa, ilk sütundaki sayısal değere bak
                    if not soru_no and len(row.values) > 0:
                        first_val = row.iloc[0]
                        if pd.notna(first_val):
                            first_str = str(first_val).strip()
                            if first_str.replace('.', '').isdigit():
                                sayi = int(float(first_str)) if '.' in first_str else int(first_str)
                                if 1 <= sayi <= 100:
                                    soru_no = str(sayi)
                    
                    # 2. KAZANIM METNİNİ BUL
                    if kazanim_metin_sutun_index < len(row.values):
                        kazanim_val = row.iloc[kazanim_metin_sutun_index]
                        if pd.notna(kazanim_val):
                            kazanim_metni = str(kazanim_val).strip()
                            # Eğer çok kısaysa, sonraki sütunlara bak
                            if len(kazanim_metni) < 5:
                                # Sonraki 3 sütuna bak
                                for next_idx in range(kazanim_metin_sutun_index + 1, min(kazanim_metin_sutun_index + 4, len(row.values))):
                                    next_val = row.iloc[next_idx]
                                    if pd.notna(next_val):
                                        next_str = str(next_val).strip()
                                        # Sayısal değilse ve yeterince uzunsa
                                        if not next_str.replace('.', '').isdigit() and len(next_str) > 5:
                                            kazanim_metni = next_str
                                            break
                    
                    # 3. PUANI BUL
                    if puan_sutun_index is not None and puan_sutun_index < len(row.values):
                        puan_val = row.iloc[puan_sutun_index]
                        if pd.notna(puan_val):
                            try:
                                if isinstance(puan_val, (int, float)):
                                    max_puan = int(puan_val)
                                else:
                                    puan_str = str(puan_val).strip()
                                    if puan_str.replace('.', '').isdigit():
                                        max_puan = int(float(puan_str)) if '.' in puan_str else int(puan_str)
                            except:
                                pass
                    
                    # Puan bulunamadıysa, satırdaki diğer sayısal değerlere bak
                    if max_puan is None:
                        for idx, val in enumerate(row.values):
                            if pd.notna(val):
                                val_str = str(val).strip()
                                if val_str.replace('.', '').isdigit():
                                    sayi = int(float(val_str)) if '.' in val_str else int(val_str)
                                    # Soru numarası değilse ve mantıklı bir puan değeriyse
                                    if soru_no and str(sayi) == soru_no:
                                        continue
                                    if 5 <= sayi <= 100:
                                        # Kazanım metninin sağındaki sayılar puan olma ihtimali daha yüksek
                                        if kazanim_metin_sutun_index is not None and idx > kazanim_metin_sutun_index:
                                            max_puan = sayi
                                            break
                    
                    # Son çare: Öğrenci notlarından tahmin et
                    if max_puan is None and soru_no:
                        soru_col = None
                        for col in df_students.columns:
                            col_str = str(col).strip().split('.')[0]
                            if col_str == soru_no:
                                soru_col = col
                                break
                        
                        if soru_col:
                            notlar = pd.to_numeric(df_students[soru_col], errors='coerce').dropna()
                            if len(notlar) > 0:
                                max_not = notlar.max()
                                # Olası puan değerlerine yuvarla
                                for puan in [5, 10, 15, 20, 25, 30, 40, 50, 100]:
                                    if max_not <= puan:
                                        max_puan = puan
                                        break
                    
                    # Varsayılan: 25
                    if max_puan is None:
                        max_puan = 25
                    
                    # Kaydet: Soru numarası ve kazanım metni varsa
                    if soru_no and kazanim_metni and len(kazanim_metni) > 3:
                        kazanim_map[soru_no] = kazanim_metni
                        soru_max_puan[soru_no] = max_puan
                        
                except Exception as e:
                    continue

        # 4. ADIM: Soru sütunlarını bul
        soru_sutunlari = []
        for col in df_students.columns:
            col_str = str(col).strip()
            col_clean = col_str.replace('.', '').replace(',', '')
            if col_clean.isdigit():
                soru_sutunlari.append(col)

        if not soru_sutunlari:
            return []

        # 5. ADIM: Kuralları uygula
        for soru_col in soru_sutunlari:
            soru_no = str(soru_col).strip().split('.')[0]
            
            konu = kazanim_map.get(soru_no, f"{soru_no}. Soru")
            max_p = soru_max_puan.get(soru_no, 25)
            sinif_yuzde = sinif_basari_yuzdeleri.get(soru_no)
            
            # Eğer sınıf başarı yüzdesi bulunamadıysa, öğrenci notlarından hesapla
            if sinif_yuzde is None:
                notlar = pd.to_numeric(df_students[soru_col], errors='coerce').dropna()
                if len(notlar) > 0:
                    sinif_ort = notlar.mean()
                    sinif_yuzde = (sinif_ort / max_p) * 100 if max_p > 0 else 0
                else:
                    sinif_yuzde = 0

            # KURAL 1: Sınıf başarısı %35 ve altındaysa TÜM SINIF
            if sinif_yuzde <= 35:
                etut_listesi.append({
                    "Dosya": dosya_adi,
                    "Soru": soru_no,
                    "Kazanım": konu,
                    "Etüt Grubu": 1,
                    "Öğrenciler": "TÜM SINIF",
                    "Sebep": f"Sınıf Başarısı: %{sinif_yuzde:.1f} (≤%35)",
                    "Etüt Süresi": "40 dakika",
                    "Soru Sayısı": "5 soru",
                    "Etüt Tipi": "Sınıf Etütü"
                })
                continue

            # KURAL 2: Bireysel kontrol
            limit = max_p * 0.5
            ogrenciler_detayli = []  # (isim, puan) tuple listesi
            
            for _, ogrenci in df_students.iterrows():
                try:
                    puan = pd.to_numeric(ogrenci[soru_col], errors='coerce')
                    if pd.isna(puan):
                        continue
                    
                    if puan <= limit:
                        ogrenci_adi = None
                        if adi_soyadi_col:
                            ogrenci_adi_val = ogrenci.get(adi_soyadi_col)
                            if pd.notna(ogrenci_adi_val):
                                ogrenci_adi = str(ogrenci_adi_val).strip()
                                if 'Kazanım' in ogrenci_adi or 'KAZANIM' in ogrenci_adi:
                                    ogrenci_adi = None
                        
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
                            ogrenciler_detayli.append((ogrenci_adi, puan))
                except:
                    continue
            
            # Öğrencileri etüt gruplarına böl (maksimum 4 kişi)
            if ogrenciler_detayli:
                max_ogrenci_per_etut = 4
                etut_grup_no = 1
                
                for i in range(0, len(ogrenciler_detayli), max_ogrenci_per_etut):
                    grup_ogrenciler = ogrenciler_detayli[i:i+max_ogrenci_per_etut]
                    
                    # Öğrenci isimlerini ve puanlarını formatla
                    ogrenci_isimleri = [og[0] for og in grup_ogrenciler]
                    ogrenci_puanlari = [f"{og[0]} ({og[1]:.1f}/{max_p})" for og in grup_ogrenciler]
                    
                    ogrenciler_str = ", ".join(ogrenci_isimleri)
                    sebep_detay = " | ".join(ogrenci_puanlari)
                    
                    etut_listesi.append({
                        "Dosya": dosya_adi,
                        "Soru": soru_no,
                        "Kazanım": konu,
                        "Etüt Grubu": etut_grup_no,
                        "Öğrenciler": ogrenciler_str,
                        "Sebep": sebep_detay,
                        "Etüt Süresi": "20 dakika",
                        "Soru Sayısı": "5 soru",
                        "Etüt Tipi": "Bireysel Etüt"
                    })
                    etut_grup_no += 1

        return etut_listesi


def main():
    root = tk.Tk()
    app = EtutListesiProgrami(root)
    root.mainloop()


if __name__ == "__main__":
    main()
