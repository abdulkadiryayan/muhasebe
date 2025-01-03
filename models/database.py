import sqlite3
from datetime import datetime

class MuhasebeDB:
    def __init__(self):
        self.conn = sqlite3.connect('muhasebe.db')
        self.cursor = self.conn.cursor()
        self.tablolari_olustur()

    def tablolari_olustur(self):
        # Cari Hesap tablosu
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS cari_hesap (
            id INTEGER PRIMARY KEY,
            musteri_adi TEXT NOT NULL,
            borc REAL DEFAULT 0,
            alacak REAL DEFAULT 0
        )
        ''')

        # Kasa Yönetimi tablosu
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS kasa_yonetimi (
            id INTEGER PRIMARY KEY,
            kasa_adi TEXT NOT NULL,
            gelir REAL DEFAULT 0,
            gider REAL DEFAULT 0
        )
        ''')

        # Faturalar ve İrsaliyeler tablosu
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS faturalar_irsaliyeler (
            id INTEGER PRIMARY KEY,
            fatura_no TEXT NOT NULL,
            tutar REAL NOT NULL,
            tur TEXT NOT NULL
        )
        ''')

        # Çek ve Senet tablosu
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS cek_senet (
            id INTEGER PRIMARY KEY,
            evrak_turu TEXT NOT NULL,
            vade_tarihi TEXT NOT NULL,
            tutar REAL NOT NULL
        )
        ''')
        
        self.conn.commit()

    # Cari Hesap İşlemleri
    def cari_hesap_ekle(self, musteri_adi, borc=0, alacak=0):
        self.cursor.execute('''
        INSERT INTO cari_hesap (musteri_adi, borc, alacak)
        VALUES (?, ?, ?)
        ''', (musteri_adi, borc, alacak))
        self.conn.commit()

    def cari_hesap_guncelle(self, id, musteri_adi=None, borc=None, alacak=None):
        mevcut = self.cursor.execute('SELECT * FROM cari_hesap WHERE id=?', (id,)).fetchone()
        if mevcut:
            yeni_musteri_adi = musteri_adi if musteri_adi is not None else mevcut[1]
            yeni_borc = borc if borc is not None else mevcut[2]
            yeni_alacak = alacak if alacak is not None else mevcut[3]
            
            self.cursor.execute('''
            UPDATE cari_hesap 
            SET musteri_adi=?, borc=?, alacak=?
            WHERE id=?
            ''', (yeni_musteri_adi, yeni_borc, yeni_alacak, id))
            self.conn.commit()

    def cari_hesap_sil(self, id):
        self.cursor.execute('DELETE FROM cari_hesap WHERE id=?', (id,))
        self.conn.commit()

    # Kasa Yönetimi İşlemleri
    def kasa_ekle(self, kasa_adi, gelir=0, gider=0):
        self.cursor.execute('''
        INSERT INTO kasa_yonetimi (kasa_adi, gelir, gider)
        VALUES (?, ?, ?)
        ''', (kasa_adi, gelir, gider))
        self.conn.commit()

    def kasa_guncelle(self, id, kasa_adi=None, gelir=None, gider=None):
        mevcut = self.cursor.execute('SELECT * FROM kasa_yonetimi WHERE id=?', (id,)).fetchone()
        if mevcut:
            yeni_kasa_adi = kasa_adi if kasa_adi is not None else mevcut[1]
            yeni_gelir = gelir if gelir is not None else mevcut[2]
            yeni_gider = gider if gider is not None else mevcut[3]
            
            self.cursor.execute('''
            UPDATE kasa_yonetimi 
            SET kasa_adi=?, gelir=?, gider=?
            WHERE id=?
            ''', (yeni_kasa_adi, yeni_gelir, yeni_gider, id))
            self.conn.commit()

    def kasa_sil(self, id):
        self.cursor.execute('DELETE FROM kasa_yonetimi WHERE id=?', (id,))
        self.conn.commit()

    # Fatura ve İrsaliye İşlemleri
    def fatura_ekle(self, fatura_no, tutar, tur):
        self.cursor.execute('''
        INSERT INTO faturalar_irsaliyeler (fatura_no, tutar, tur)
        VALUES (?, ?, ?)
        ''', (fatura_no, tutar, tur))
        self.conn.commit()

    def fatura_guncelle(self, id, fatura_no=None, tutar=None, tur=None):
        mevcut = self.cursor.execute('SELECT * FROM faturalar_irsaliyeler WHERE id=?', (id,)).fetchone()
        if mevcut:
            yeni_fatura_no = fatura_no if fatura_no is not None else mevcut[1]
            yeni_tutar = tutar if tutar is not None else mevcut[2]
            yeni_tur = tur if tur is not None else mevcut[3]
            
            self.cursor.execute('''
            UPDATE faturalar_irsaliyeler 
            SET fatura_no=?, tutar=?, tur=?
            WHERE id=?
            ''', (yeni_fatura_no, yeni_tutar, yeni_tur, id))
            self.conn.commit()

    def fatura_sil(self, id):
        self.cursor.execute('DELETE FROM faturalar_irsaliyeler WHERE id=?', (id,))
        self.conn.commit()

    # Çek ve Senet İşlemleri
    def cek_senet_ekle(self, evrak_turu, vade_tarihi, tutar):
        self.cursor.execute('''
        INSERT INTO cek_senet (evrak_turu, vade_tarihi, tutar)
        VALUES (?, ?, ?)
        ''', (evrak_turu, vade_tarihi, tutar))
        self.conn.commit()

    def cek_senet_guncelle(self, id, evrak_turu=None, vade_tarihi=None, tutar=None):
        mevcut = self.cursor.execute('SELECT * FROM cek_senet WHERE id=?', (id,)).fetchone()
        if mevcut:
            yeni_evrak_turu = evrak_turu if evrak_turu is not None else mevcut[1]
            yeni_vade_tarihi = vade_tarihi if vade_tarihi is not None else mevcut[2]
            yeni_tutar = tutar if tutar is not None else mevcut[3]
            
            self.cursor.execute('''
            UPDATE cek_senet 
            SET evrak_turu=?, vade_tarihi=?, tutar=?
            WHERE id=?
            ''', (yeni_evrak_turu, yeni_vade_tarihi, yeni_tutar, id))
            self.conn.commit()

    def cek_senet_sil(self, id):
        self.cursor.execute('DELETE FROM cek_senet WHERE id=?', (id,))
        self.conn.commit()

    def __del__(self):
        self.conn.close() 