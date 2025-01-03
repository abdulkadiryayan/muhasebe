import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from viewmodels.muhasebe_viewmodel import MuhasebeViewModel
from tkcalendar import DateEntry
import os
from datetime import datetime

class AnaPencere(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.viewmodel = MuhasebeViewModel()
        
        self.title("Muhasebe Uygulaması")
        self.geometry("1024x768")
        
        # Ana menü oluşturma
        self.menu_olustur()
        
        # Ana notebook (sekmeli panel) oluşturma
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Sekmeleri oluştur
        self.cari_hesap_sekmesi_olustur()
        self.kasa_sekmesi_olustur()
        self.fatura_sekmesi_olustur()
        self.cek_senet_sekmesi_olustur()
        
        # Tablo seçim olaylarını bağla
        self.cari_tablo.bind('<<TreeviewSelect>>', 
            lambda e: self.tablo_secim_olayi(e, self.cari_tablo, self.cari_hesap_secildi))
        
        self.kasa_tablo.bind('<<TreeviewSelect>>', 
            lambda e: self.tablo_secim_olayi(e, self.kasa_tablo, self.kasa_secildi))
        
        self.fatura_tablo.bind('<<TreeviewSelect>>', 
            lambda e: self.tablo_secim_olayi(e, self.fatura_tablo, self.fatura_secildi))
        
        self.cek_senet_tablo.bind('<<TreeviewSelect>>', 
            lambda e: self.tablo_secim_olayi(e, self.cek_senet_tablo, self.cek_senet_secildi))

    def menu_olustur(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # Dosya menüsü
        dosya_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Dosya", menu=dosya_menu)
        dosya_menu.add_command(label="Excel'e Aktar", command=self.excel_export)
        dosya_menu.add_separator()
        dosya_menu.add_command(label="Çıkış", command=self.quit)
        
        # Rapor menüsü
        rapor_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Raporlar", menu=rapor_menu)
        rapor_menu.add_command(label="Cari Hesap Raporu")
        rapor_menu.add_command(label="Kasa Raporu")
        rapor_menu.add_command(label="Fatura Raporu")
        rapor_menu.add_command(label="Çek/Senet Raporu")

    def tablo_olustur(self, parent):
        # Tablo oluşturma fonksiyonu
        tablo = ttk.Treeview(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tablo.yview)
        tablo.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        return tablo

    def cari_hesap_sekmesi_olustur(self):
        cari_frame = ttk.Frame(self.notebook)
        self.notebook.add(cari_frame, text="Cari Hesaplar")
        
        # Sol frame - Form
        form_frame = ttk.LabelFrame(cari_frame, text="Cari Hesap Ekle/Düzenle", padding=10)
        form_frame.pack(side="left", fill="y", padx=5, pady=5)
        
        # Form elemanları
        ttk.Label(form_frame, text="Müşteri Adı:").grid(row=0, column=0, padx=5, pady=5)
        self.musteri_adi = ttk.Entry(form_frame)
        self.musteri_adi.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Borç:").grid(row=1, column=0, padx=5, pady=5)
        self.borc = ttk.Entry(form_frame)
        self.borc.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Alacak:").grid(row=2, column=0, padx=5, pady=5)
        self.alacak = ttk.Entry(form_frame)
        self.alacak.grid(row=2, column=1, padx=5, pady=5)
        
        # Butonlar
        ttk.Button(form_frame, text="Kaydet", command=self.cari_hesap_kaydet).grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(form_frame, text="Sil", command=self.cari_hesap_sil).grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(form_frame, text="Güncelle", command=self.cari_hesap_guncelle).grid(row=5, column=0, columnspan=2, pady=5)
        
        # Sağ frame - Tablo
        tablo_frame = ttk.LabelFrame(cari_frame, text="Cari Hesap Listesi", padding=10)
        tablo_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        # Tablo oluşturma
        self.cari_tablo = self.tablo_olustur(tablo_frame)
        self.cari_tablo["columns"] = ("id", "musteri_adi", "borc", "alacak", "bakiye")
        self.cari_tablo.column("#0", width=0, stretch=tk.NO)
        self.cari_tablo.column("id", width=50)
        self.cari_tablo.column("musteri_adi", width=150)
        self.cari_tablo.column("borc", width=100)
        self.cari_tablo.column("alacak", width=100)
        self.cari_tablo.column("bakiye", width=100)
        
        self.cari_tablo.heading("id", text="ID")
        self.cari_tablo.heading("musteri_adi", text="Müşteri Adı")
        self.cari_tablo.heading("borc", text="Borç")
        self.cari_tablo.heading("alacak", text="Alacak")
        self.cari_tablo.heading("bakiye", text="Bakiye")
        
        self.cari_tablo.pack(fill="both", expand=True)
        self.cari_hesaplari_listele()

    def kasa_sekmesi_olustur(self):
        kasa_frame = ttk.Frame(self.notebook)
        self.notebook.add(kasa_frame, text="Kasa")
        
        # Sol frame - Form
        form_frame = ttk.LabelFrame(kasa_frame, text="Kasa İşlemi Ekle/Düzenle", padding=10)
        form_frame.pack(side="left", fill="y", padx=5, pady=5)
        
        # Form elemanları
        ttk.Label(form_frame, text="Kasa Adı:").grid(row=0, column=0, padx=5, pady=5)
        self.kasa_adi = ttk.Entry(form_frame)
        self.kasa_adi.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Gelir:").grid(row=1, column=0, padx=5, pady=5)
        self.gelir = ttk.Entry(form_frame)
        self.gelir.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Gider:").grid(row=2, column=0, padx=5, pady=5)
        self.gider = ttk.Entry(form_frame)
        self.gider.grid(row=2, column=1, padx=5, pady=5)
        
        # Butonlar
        ttk.Button(form_frame, text="Kaydet", command=self.kasa_kaydet).grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(form_frame, text="Sil", command=self.kasa_sil).grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(form_frame, text="Güncelle", command=self.kasa_guncelle).grid(row=5, column=0, columnspan=2, pady=5)
        
        # Sağ frame - Tablo
        tablo_frame = ttk.LabelFrame(kasa_frame, text="Kasa Hareketleri", padding=10)
        tablo_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        # Tablo oluşturma
        self.kasa_tablo = self.tablo_olustur(tablo_frame)
        self.kasa_tablo["columns"] = ("id", "kasa_adi", "gelir", "gider", "bakiye")
        self.kasa_tablo.column("#0", width=0, stretch=tk.NO)
        self.kasa_tablo.column("id", width=50)
        self.kasa_tablo.column("kasa_adi", width=150)
        self.kasa_tablo.column("gelir", width=100)
        self.kasa_tablo.column("gider", width=100)
        self.kasa_tablo.column("bakiye", width=100)
        
        self.kasa_tablo.heading("id", text="ID")
        self.kasa_tablo.heading("kasa_adi", text="Kasa Adı")
        self.kasa_tablo.heading("gelir", text="Gelir")
        self.kasa_tablo.heading("gider", text="Gider")
        self.kasa_tablo.heading("bakiye", text="Bakiye")
        
        self.kasa_tablo.pack(fill="both", expand=True)

    def fatura_sekmesi_olustur(self):
        fatura_frame = ttk.Frame(self.notebook)
        self.notebook.add(fatura_frame, text="Faturalar")
        
        # Sol frame - Form
        form_frame = ttk.LabelFrame(fatura_frame, text="Fatura Ekle/Düzenle", padding=10)
        form_frame.pack(side="left", fill="y", padx=5, pady=5)
        
        # Form elemanları
        ttk.Label(form_frame, text="Fatura No:").grid(row=0, column=0, padx=5, pady=5)
        self.fatura_no = ttk.Entry(form_frame)
        self.fatura_no.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Tutar:").grid(row=1, column=0, padx=5, pady=5)
        self.fatura_tutar = ttk.Entry(form_frame)
        self.fatura_tutar.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Tür:").grid(row=2, column=0, padx=5, pady=5)
        self.fatura_tur = ttk.Combobox(form_frame, values=["Satış Faturası", "Alış Faturası"])
        self.fatura_tur.grid(row=2, column=1, padx=5, pady=5)
        
        # Butonlar
        ttk.Button(form_frame, text="Kaydet", command=self.fatura_kaydet).grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(form_frame, text="Sil", command=self.fatura_sil).grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(form_frame, text="Güncelle", command=self.fatura_guncelle).grid(row=5, column=0, columnspan=2, pady=5)
        
        # Sağ frame - Tablo
        tablo_frame = ttk.LabelFrame(fatura_frame, text="Fatura Listesi", padding=10)
        tablo_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        # Tablo oluşturma
        self.fatura_tablo = self.tablo_olustur(tablo_frame)
        self.fatura_tablo["columns"] = ("id", "fatura_no", "tutar", "tur")
        self.fatura_tablo.column("#0", width=0, stretch=tk.NO)
        self.fatura_tablo.column("id", width=50)
        self.fatura_tablo.column("fatura_no", width=150)
        self.fatura_tablo.column("tutar", width=100)
        self.fatura_tablo.column("tur", width=100)
        
        self.fatura_tablo.heading("id", text="ID")
        self.fatura_tablo.heading("fatura_no", text="Fatura No")
        self.fatura_tablo.heading("tutar", text="Tutar")
        self.fatura_tablo.heading("tur", text="Tür")
        
        self.fatura_tablo.pack(fill="both", expand=True)

    def cek_senet_sekmesi_olustur(self):
        cek_senet_frame = ttk.Frame(self.notebook)
        self.notebook.add(cek_senet_frame, text="Çek/Senet")
        
        # Sol frame - Form
        form_frame = ttk.LabelFrame(cek_senet_frame, text="Çek/Senet Ekle/Düzenle", padding=10)
        form_frame.pack(side="left", fill="y", padx=5, pady=5)
        
        # Form elemanları
        ttk.Label(form_frame, text="Evrak Türü:").grid(row=0, column=0, padx=5, pady=5)
        self.evrak_turu = ttk.Combobox(form_frame, values=["Çek", "Senet"])
        self.evrak_turu.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Vade Tarihi:").grid(row=1, column=0, padx=5, pady=5)
        self.vade_tarihi = DateEntry(form_frame, width=12, background='darkblue',
                                   foreground='white', borderwidth=2)
        self.vade_tarihi.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Tutar:").grid(row=2, column=0, padx=5, pady=5)
        self.evrak_tutar = ttk.Entry(form_frame)
        self.evrak_tutar.grid(row=2, column=1, padx=5, pady=5)
        
        # Butonlar
        ttk.Button(form_frame, text="Kaydet", command=self.cek_senet_kaydet).grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(form_frame, text="Sil", command=self.cek_senet_sil).grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(form_frame, text="Güncelle", command=self.cek_senet_guncelle).grid(row=5, column=0, columnspan=2, pady=5)
        
        # Sağ frame - Tablo
        tablo_frame = ttk.LabelFrame(cek_senet_frame, text="Çek/Senet Listesi", padding=10)
        tablo_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        # Tablo oluşturma
        self.cek_senet_tablo = self.tablo_olustur(tablo_frame)
        self.cek_senet_tablo["columns"] = ("id", "evrak_turu", "vade_tarihi", "tutar")
        self.cek_senet_tablo.column("#0", width=0, stretch=tk.NO)
        self.cek_senet_tablo.column("id", width=50)
        self.cek_senet_tablo.column("evrak_turu", width=100)
        self.cek_senet_tablo.column("vade_tarihi", width=100)
        self.cek_senet_tablo.column("tutar", width=100)
        
        self.cek_senet_tablo.heading("id", text="ID")
        self.cek_senet_tablo.heading("evrak_turu", text="Evrak Türü")
        self.cek_senet_tablo.heading("vade_tarihi", text="Vade Tarihi")
        self.cek_senet_tablo.heading("tutar", text="Tutar")
        
        self.cek_senet_tablo.pack(fill="both", expand=True)

    # Cari Hesap İşlemleri
    def cari_hesap_kaydet(self):
        try:
            if self.viewmodel.cari_hesap_ekle(
                self.musteri_adi.get(),
                float(self.borc.get() or 0),
                float(self.alacak.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Cari hesap kaydedildi")
                self.cari_hesaplari_listele()
                self.cari_hesap_formu_temizle()
            else:
                messagebox.showerror("Hata", "Kayıt sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    def cari_hesaplari_listele(self):
        for item in self.cari_tablo.get_children():
            self.cari_tablo.delete(item)
        
        for hesap in self.viewmodel.cari_hesap_listele():
            bakiye = self.viewmodel.cari_hesap_bakiye_hesapla(hesap[0])
            self.cari_tablo.insert("", "end", values=(hesap[0], hesap[1], hesap[2], hesap[3], bakiye))

    def cari_hesap_formu_temizle(self):
        self.musteri_adi.delete(0, tk.END)
        self.borc.delete(0, tk.END)
        self.alacak.delete(0, tk.END)

    # Kasa İşlemleri
    def kasa_kaydet(self):
        try:
            if self.viewmodel.kasa_ekle(
                self.kasa_adi.get(),
                float(self.gelir.get() or 0),
                float(self.gider.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Kasa kaydedildi")
                self.kasa_listele()
                self.kasa_formu_temizle()
            else:
                messagebox.showerror("Hata", "Kayıt sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    def kasa_listele(self):
        for item in self.kasa_tablo.get_children():
            self.kasa_tablo.delete(item)
        
        for kasa in self.viewmodel.kasa_listele():
            bakiye = self.viewmodel.kasa_bakiye_hesapla(kasa[0])
            self.kasa_tablo.insert("", "end", values=(kasa[0], kasa[1], kasa[2], kasa[3], bakiye))

    def kasa_formu_temizle(self):
        self.kasa_adi.delete(0, tk.END)
        self.gelir.delete(0, tk.END)
        self.gider.delete(0, tk.END)

    # Fatura İşlemleri
    def fatura_kaydet(self):
        try:
            if self.viewmodel.fatura_ekle(
                self.fatura_no.get(),
                float(self.fatura_tutar.get() or 0),
                self.fatura_tur.get()
            ):
                messagebox.showinfo("Başarılı", "Fatura kaydedildi")
                self.fatura_listele()
                self.fatura_formu_temizle()
            else:
                messagebox.showerror("Hata", "Kayıt sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    def fatura_listele(self):
        for item in self.fatura_tablo.get_children():
            self.fatura_tablo.delete(item)
        
        for fatura in self.viewmodel.fatura_listele():
            self.fatura_tablo.insert("", "end", values=fatura)

    def fatura_formu_temizle(self):
        self.fatura_no.delete(0, tk.END)
        self.fatura_tutar.delete(0, tk.END)
        self.fatura_tur.set('')

    # Çek/Senet İşlemleri
    def cek_senet_kaydet(self):
        try:
            if self.viewmodel.cek_senet_ekle(
                self.evrak_turu.get(),
                self.vade_tarihi.get(),
                float(self.evrak_tutar.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Çek/Senet kaydedildi")
                self.cek_senet_listele()
                self.cek_senet_formu_temizle()
            else:
                messagebox.showerror("Hata", "Kayıt sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    def cek_senet_listele(self):
        for item in self.cek_senet_tablo.get_children():
            self.cek_senet_tablo.delete(item)
        
        for evrak in self.viewmodel.vadesi_yaklasan_cek_senetler():
            self.cek_senet_tablo.insert("", "end", values=evrak)

    def cek_senet_formu_temizle(self):
        self.evrak_turu.set('')
        self.vade_tarihi.delete(0, tk.END)
        self.evrak_tutar.delete(0, tk.END)

    # Cari Hesap İşlemleri - Silme ve Güncelleme
    def cari_hesap_sil(self):
        selected = self.cari_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen silinecek kaydı seçin")
            return
            
        if messagebox.askyesno("Onay", "Seçili kayıt silinecek. Emin misiniz?"):
            item = self.cari_tablo.item(selected[0])
            id = item['values'][0]
            if self.viewmodel.cari_hesap_sil(id):
                messagebox.showinfo("Başarılı", "Kayıt silindi")
                self.cari_hesaplari_listele()
            else:
                messagebox.showerror("Hata", "Kayıt silinirken bir hata oluştu")

    def cari_hesap_guncelle(self):
        selected = self.cari_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen güncellenecek kaydı seçin")
            return
            
        item = self.cari_tablo.item(selected[0])
        id = item['values'][0]
        
        try:
            if self.viewmodel.cari_hesap_guncelle(
                id,
                self.musteri_adi.get(),
                float(self.borc.get() or 0),
                float(self.alacak.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Kayıt güncellendi")
                self.cari_hesaplari_listele()
                self.cari_hesap_formu_temizle()
            else:
                messagebox.showerror("Hata", "Güncelleme sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    # Kasa İşlemleri - Silme ve Güncelleme
    def kasa_sil(self):
        selected = self.kasa_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen silinecek kaydı seçin")
            return
            
        if messagebox.askyesno("Onay", "Seçili kayıt silinecek. Emin misiniz?"):
            item = self.kasa_tablo.item(selected[0])
            id = item['values'][0]
            if self.viewmodel.kasa_sil(id):
                messagebox.showinfo("Başarılı", "Kayıt silindi")
                self.kasa_listele()
            else:
                messagebox.showerror("Hata", "Kayıt silinirken bir hata oluştu")

    def kasa_guncelle(self):
        selected = self.kasa_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen güncellenecek kaydı seçin")
            return
            
        item = self.kasa_tablo.item(selected[0])
        id = item['values'][0]
        
        try:
            if self.viewmodel.kasa_guncelle(
                id,
                self.kasa_adi.get(),
                float(self.gelir.get() or 0),
                float(self.gider.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Kayıt güncellendi")
                self.kasa_listele()
                self.kasa_formu_temizle()
            else:
                messagebox.showerror("Hata", "Güncelleme sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    # Fatura İşlemleri - Silme ve Güncelleme
    def fatura_sil(self):
        selected = self.fatura_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen silinecek kaydı seçin")
            return
            
        if messagebox.askyesno("Onay", "Seçili kayıt silinecek. Emin misiniz?"):
            item = self.fatura_tablo.item(selected[0])
            id = item['values'][0]
            if self.viewmodel.fatura_sil(id):
                messagebox.showinfo("Başarılı", "Kayıt silindi")
                self.fatura_listele()
            else:
                messagebox.showerror("Hata", "Kayıt silinirken bir hata oluştu")

    def fatura_guncelle(self):
        selected = self.fatura_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen güncellenecek kaydı seçin")
            return
            
        item = self.fatura_tablo.item(selected[0])
        id = item['values'][0]
        
        try:
            if self.viewmodel.fatura_guncelle(
                id,
                self.fatura_no.get(),
                float(self.fatura_tutar.get() or 0),
                self.fatura_tur.get()
            ):
                messagebox.showinfo("Başarılı", "Kayıt güncellendi")
                self.fatura_listele()
                self.fatura_formu_temizle()
            else:
                messagebox.showerror("Hata", "Güncelleme sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    # Çek/Senet İşlemleri - Silme ve Güncelleme
    def cek_senet_sil(self):
        selected = self.cek_senet_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen silinecek kaydı seçin")
            return
            
        if messagebox.askyesno("Onay", "Seçili kayıt silinecek. Emin misiniz?"):
            item = self.cek_senet_tablo.item(selected[0])
            id = item['values'][0]
            if self.viewmodel.cek_senet_sil(id):
                messagebox.showinfo("Başarılı", "Kayıt silindi")
                self.cek_senet_listele()
            else:
                messagebox.showerror("Hata", "Kayıt silinirken bir hata oluştu")

    def cek_senet_guncelle(self):
        selected = self.cek_senet_tablo.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen güncellenecek kaydı seçin")
            return
            
        item = self.cek_senet_tablo.item(selected[0])
        id = item['values'][0]
        
        try:
            if self.viewmodel.cek_senet_guncelle(
                id,
                self.evrak_turu.get(),
                self.vade_tarihi.get(),
                float(self.evrak_tutar.get() or 0)
            ):
                messagebox.showinfo("Başarılı", "Kayıt güncellendi")
                self.cek_senet_listele()
                self.cek_senet_formu_temizle()
            else:
                messagebox.showerror("Hata", "Güncelleme sırasında bir hata oluştu")
        except ValueError:
            messagebox.showerror("Hata", "Lütfen sayısal değerleri doğru giriniz")

    # Tablo seçim olayları
    def tablo_secim_olayi(self, event, tablo, form_doldur_fonksiyonu):
        selected = tablo.selection()
        if selected:
            item = tablo.item(selected[0])
            form_doldur_fonksiyonu(item['values'])

    def cari_hesap_secildi(self, values):
        self.musteri_adi.delete(0, tk.END)
        self.borc.delete(0, tk.END)
        self.alacak.delete(0, tk.END)
        
        self.musteri_adi.insert(0, values[1])
        self.borc.insert(0, values[2])
        self.alacak.insert(0, values[3])

    def kasa_secildi(self, values):
        self.kasa_adi.delete(0, tk.END)
        self.gelir.delete(0, tk.END)
        self.gider.delete(0, tk.END)
        
        self.kasa_adi.insert(0, values[1])
        self.gelir.insert(0, values[2])
        self.gider.insert(0, values[3])

    def fatura_secildi(self, values):
        self.fatura_no.delete(0, tk.END)
        self.fatura_tutar.delete(0, tk.END)
        
        self.fatura_no.insert(0, values[1])
        self.fatura_tutar.insert(0, values[2])
        self.fatura_tur.set(values[3])

    def cek_senet_secildi(self, values):
        self.evrak_turu.set(values[1])
        self.vade_tarihi.delete(0, tk.END)
        self.evrak_tutar.delete(0, tk.END)
        
        self.vade_tarihi.insert(0, values[2])
        self.evrak_tutar.insert(0, values[3])

    def excel_export(self):
        try:
            # Varsayılan dosya adını oluştur
            tarih = datetime.now().strftime('%Y%m%d_%H%M%S')
            varsayilan_dosya = f"muhasebe_rapor_{tarih}.xlsx"
            
            # Dosya kaydetme dialogu
            dosya_yolu = filedialog.asksaveasfilename(
                initialfile=varsayilan_dosya,
                defaultextension=".xlsx",
                filetypes=[("Excel Dosyası", "*.xlsx")],
                title="Excel Dosyasını Kaydet"
            )
            
            if not dosya_yolu:  # Kullanıcı iptal ettiyse
                return
            
            # Excel'e aktar
            basarili, mesaj = self.viewmodel.excel_export(dosya_yolu)
            
            if basarili:
                cevap = messagebox.askquestion(
                    "Başarılı", 
                    f"Veriler Excel dosyasına aktarıldı:\n{dosya_yolu}\n\nDosyayı açmak ister misiniz?"
                )
                if cevap == 'yes':
                    try:
                        os.startfile(dosya_yolu)  # Windows'ta dosyayı aç
                    except:
                        # Windows dışı sistemler için alternatif açma yöntemi
                        import subprocess
                        subprocess.Popen(['start', dosya_yolu], shell=True)
            else:
                messagebox.showerror(
                    "Hata",
                    f"Excel dosyası oluşturulurken bir hata oluştu:\n{mesaj}"
                )
                
        except Exception as e:
            messagebox.showerror(
                "Hata",
                f"Beklenmeyen bir hata oluştu:\n{str(e)}"
            ) 