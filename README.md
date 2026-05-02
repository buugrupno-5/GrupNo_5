# 🏦 Tasarruf Planı ve Borç Azaltma Modeli: Dizilerin Yakınsaklığı

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange.svg)
![Precision](https://img.shields.io/badge/Precision-28_digits_Decimal-success.svg)

Bu proje, Bursa Uludağ Üniversitesi Matematik Bölümü öğrencileri tarafından finansal fark denklemleri ve dizilerin yakınsaklığı prensipleri kullanılarak geliştirilmiş bir masaüstü uygulamasıdır. Birikim hedeflerini ve kredi/borç ödeme planlarını yüksek matematiksel hassasiyetle simüle eden, grafik destekli, gelişmiş bir masaüstü finansal analiz uygulamasıdır. Standart hesaplama araçlarının aksine, finansal serilerin yakınsaklık ve ıraksaklık durumlarını analiz eder.

## 📑 İçindekiler
- [Projenin Amacı ve Kullanım Alanları](#-projenin-amacı-ve-kullanım-alanları)
- [Öne Çıkan Teknik Özellikler](#-öne-çıkan-teknik-özellikler)
- [Modüller ve İşlevler](#-modüller-ve-işlevler)
- [Matematiksel Altyapı](#-matematiksel-altyapı)
- [Kurulum ve Çalıştırma](#-kurulum-ve-çalıştırma)
- [Geliştirici](#-Geliştirici)

---

## 🎯 Projenin Amacı ve Kullanım Alanları
Finansal hesaplamalarda sıkça karşılaşılan kuruşluk sapmaları (float hataları) ortadan kaldırarak, sıfır hata payıyla çalışmak üzere tasarlanmıştır. Aşağıdaki kullanım senaryoları için idealdir:

*>* **Akademik Analizler:** Faiz hesaplamalarını standart 30 gün yerine, gerçek takvim günlerine (Act/365) göre yapmak ve enflasyonun paranın değerine olan etkisini test etmek.

*>* **Bireysel Finans Planlaması:** Düzenli yatırımları, vadeden önce para çekme (vade bozma) durumlarını ve ara ödemeli kredi planlarını grafikler üzerinden kolayca izlemek.

*>* **Banka Simülasyonları:** İhtiyaç, taşıt ve konut kredilerindeki yasal vergi (KKDF ve BSMV) kesintilerinin, cebinizden çıkacak net tutarlara nasıl yansıdığını hesaplamak.

---

## 🚀 Öne Çıkan Teknik Özellikler

Proje, yazılım mühendisliği standartlarına uygun olarak çeşitli mimari geliştirmeler içerir:

* **Kusursuz Hassasiyet (Decimal Koruması):** Standart float tiplerinin neden olduğu kuruşluk sapmaları önlemek için `decimal` kütüphanesi kullanılmış ve hassasiyet 28 basamağa (`getcontext().prec = 28`) sabitlenmiştir.
* **Dinamik Logaritmik Grafik Ölçeği:** Borcun kapanmayıp sürekli büyüdüğü "Iraksak " durumlarda grafiklerin okunabilirliğini korumak için sistem otomatik olarak lineer ölçekten logaritmik ölçeğe geçiş yapar.
* **Inline Error Handling (Satır İçi Hata Yönetimi):** Kullanıcı girişleri (`tkinter` validate komutlarıyla) anlık olarak doğrulanır. Hatalı veri girişlerinde programın çökmesi engellenir ve dinamik uyarı etiketleri ile kullanıcıya yönlendirme yapılır.
* **Karanlık ve Aydınlık Tema Desteği:** `TEMA_KARANLIK` ve `TEMA_AYDINLIK` sözlükleri ile arayüz renkleri anlık olarak değiştirilebilir, göz yormayan modern bir UI sunulur.
* **Özel Takvim Bileşeni:** İşletim sisteminden bağımsız, baştan yazılmış özel bir takvim açılır penceresi içerir.

---

## 🧰 Modüller ve İşlevler

### 1. Tasarruf Planı (Mevduat Simülasyonu)
* **Reel Alım Gücü (Enflasyonun alım gücüne etkisi):** Nominal bakiyeyi enflasyon oranına göre iskontolayarak gelecekteki reel alım gücünü hesaplar.
* **Vade Bozma (Erken Çekim):** Belirlenen dönemlerde anaparadan para çekilmesi durumunda "vade bozuldu" mantığını işletir ve yanan faizi hesaplar.
* **Özel İşlemler & Düzenli Nakit Akışı:** Her döneme özel farklı nakit giriş/çıkışları tanımlanabilir.

### 2. Borç Ekstresi (Kredi Azaltma Modeli)
* **Dinamik Kredi Tipleri:** İhtiyaç, Taşıt, Konut ve Özel kredi seçenekleriyle yasal vergi oranlarını (örn: %30 veya %0) otomatik ayarlar.
* **Ara Ödeme Sistemi:** Standart aylık taksitlerin dışına çıkılarak istenilen aylara özel "Ara Ödeme" eklenebilir.
* **Asgari Ödeme Kontrolü:** Ödenen taksitin, tahakkuk eden faizi karşılayıp karşılamadığını denetler ve "Borç Asla Kapanmaz, Büyür" uyarısı verebilir.

---

## 🧮 Matematiksel Altyapı

Program, döngüsel hesaplamalarında aşağıdaki prensipleri kullanır:
* **Günlük Faiz (Act/365):** Dönemlik tahakkuklar `(F * gün / 365)` formülüyle gerçek dünya bankacılık sistemine (Act/365 standardı) uygun hesaplanır.
* **Dizi Karakteristikleri:** 
  * *Yakınsak:* Borcun sıfırlandığı veya tasarrufun tükendiği senaryolar.
  * *Iraksak:* Faiz getirisinin/yükünün nakit akışını aştığı ve sonsuza giden seriler.

---

## ⚙️ Kurulum ve Çalıştırma

Bu proje **"Pure Python" (Saf Python)** ile geliştirilmiştir. Pandas, NumPy veya Matplotlib gibi ağır dış bağımlılıklara ihtiyaç duymaz. Tüm grafikler ve arayüz yerleşik `tkinter` modülü ile sıfırdan çizilmiştir.
1. Repoyu bilgisayarınıza klonlayın:
   ```bash
  git clone https://github.com/buugrupno-5/GrupNo_5.git
cd GrupNo_5

2.Sisteminizde Python 3.8 veya daha üstü bir sürümün kurulu olduğundan emin olun.
3.Uygulamayı başlatın:

python GrupNo_5.py



## 👥 Geliştiriciler / Proje Ekibi

Bu proje, **Bursa Uludağ Üniversitesi Matematik Bölümü** bünyesinde, "Tasarruf Planı ve Borç Azaltma Modeli: Dizilerin Yakınsaklığı" araştırması kapsamında **Grup 5** tarafından ortak bir akademik çalışma olarak geliştirilmiştir.

**Ekip Üyeleri:**
* 🎓 **Meltem Nur Yılmaz** (Öğrenci No: 082040014)
* 🎓 **Öznur Çağdaş** (Öğrenci No: 082140019)
* 🎓 **Ahmet İşbilen** (Öğrenci No: 082240067)

**Projenin matematiksel altyapısı, algoritmik tasarımı veya finansal simülasyon detayları hakkında geri bildirim vermek için proje ekibimizle iletişime geçebilirsiniz.**
