# 📊 Tasarruf Planı ve Borç Azaltma Modeli (Finansal Projeksiyon)

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

Bu proje, kullanıcıların banka mevduat getirilerini optimize etmelerini ve kredi/borç ödeme planlarını stratejik olarak yönetmelerini sağlayan gelişmiş bir **masaüstü finansal simülasyon** uygulamasıdır. 

Uygulama, standart hesaplama araçlarının aksine, gerçek takvim günlerini (Act/365), yasal stopaj/vergi kesintilerini, BDDK limitlerini ve enflasyon etkisini (reel bakiye) dikkate alarak yüksek hassasiyetli (`Decimal` kütüphanesi ile) projeksiyonlar sunar.

## ✨ Temel Özellikler

### 🏦 Tasarruf Planı (Mevduat Modeli)
* **Gerçekçi Faiz ve Vergi Hesaplaması:** Vade gününe göre yasal stopaj oranlarını otomatik belirler ve net faizi hesaplar.
* **Erken Çekim ve Vade Bozma:** Dönem içi para çekme işlemlerinde bozulan vadeyi, yanan faizleri ve bakiye aşımı durumlarını simüle eder.
* **Enflasyon Etkisi (Reel Alım Gücü):** Girilen yıllık enflasyon beklentisine göre paranın gelecekteki reel alım gücünü nominal bakiye ile karşılaştırmalı olarak sunar.
* **Düzenli ve Özel İşlemler:** Her dönem düzenli para ekleme/çekme veya belirli dönemlere özel nakit akışları tanımlama imkanı.

### 📉 Borç Ekstresi ve Azaltma Modeli
* **Dinamik Kredi Tipleri:** İhtiyaç, Taşıt, Konut ve Özel kredi senaryoları için özelleştirilmiş altyapı.
* **Yasal Sınır Kontrolleri (BDDK):** Teminat (araç/konut değeri) üzerinden maksimum çekilebilir kredi tutarını (LTV) ve vade sınırlarını denetler.
* **Gelişmiş Ara Ödeme ve Yapılandırma:** Yapılan ara ödemelerde vadeyi sabit tutarak **taksit düşürme (yeniden yapılandırma)** veya taksiti sabit tutarak **erken bitirme** senaryolarını hesaplar.
* **Dinamik Asgari Ödeme:** Yasal sınırları (maksimum vadeyi) kurtaracak minimum ödeme tutarını tahakkuk eden faiz üzerinden anlık hesaplar.

### 🖥️ Kullanıcı Deneyimi (UI/UX) ve Teknik Altyapı
* **Görselleştirme:** Etkileşimli (Hover destekli) ve logaritmik ölçekleme yapabilen dinamik Tkinter Canvas grafikleri.
* **Tema Desteği:** Tek tıkla değiştirilebilen, göz yormayan modern Aydınlık (Light) ve Karanlık (Dark) tema.
* **Gelişmiş Dışa Aktarma:** Oluşturulan ekstreleri Zebra formatlı olarak **Excel (XLSX)**, **CSV** ve stilize edilmiş **HTML** olarak dışa aktarma.
* **Akıllı Veri Girişi:** Otomatik para formatlama, interaktif takvim modülü ve hatalı girişi engelleyen validasyon sistemi.

## 🚀 Kurulum ve Çalıştırma

Proje standart Python kütüphaneleri ağırlıklı olarak geliştirilmiştir. Dışa aktarma özellikleri için sadece `openpyxl` kütüphanesine ihtiyaç duyar.

**1. Depoyu Klonlayın:**
```bash
git clone [https://github.com/KULLANICI_ADIN/finansal-analiz-simulatoru.git](https://github.com//finansal-analiz-simulatoru.git)
cd finansal-analiz-simulatoru
