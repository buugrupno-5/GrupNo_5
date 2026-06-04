# 📊 Tasarruf Planı ve Borç Azaltma Modeli Simülatörü (Grup_No5)

Bu proje, hem günlük kullanıcıların kişisel finansal hedeflerini planlamasını sağlayan hem de finans profesyonellerinin ihtiyaç duyduğu yüksek hassasiyetli matematiksel modelleri barındıran gelişmiş bir finansal simülasyon uygulamasıdır. 

Python ve `tkinter` kullanılarak geliştirilen bu masaüstü uygulaması; birikimlerin enflasyon karşısındaki reel değerini hesaplamaktan, karmaşık BDDK kurallarına tabi kredi amortisman tablolamalarına kadar geniş bir yelpazede analiz imkanı sunar.

---

## 🎯 Projenin Amacı ve Hedef Kitlesi
* **Günlük Kullanıcılar İçin:** Kredi çekerken ne kadar faiz ödeyeceğinizi, ara ödemelerin taksitlerinizi nasıl düşüreceğini veya aylık düzenli birikimlerinizin yıllar içinde ne kadara ulaşacağını kolayca ve görsel olarak görebilirsiniz.
* **Finans Profesyonelleri ve Geliştiriciler İçin:** Proje; `Act/365` gün sayma konvansiyonu, `Decimal` modülü ile 28 basamak hassasiyet (kayan nokta hatalarından arındırılmış tam doğruluk) ve Türkiye Cumhuriyet Merkez Bankası / BDDK regülasyon mantıklarını (LTV oranları, stopaj dilimleri) içeren algoritmik bir mimariye sahiptir.

---

## 🌟 Detaylı Özellikler ve Algoritmik Yetenekler

### 🏦 1. Tasarruf Planı (Mevduat Simülasyonu)
Mevduat hesaplamaları standart bir bileşik faiz formülünün çok ötesindedir. Gerçek dünya senaryolarını simüle eder:
* **Hassas Günlük Tahakkuk ve Yasal Stopaj:** Faiz getirisi vade gününe göre (Örn: 32 gün) `Act/365` mantığıyla hesaplanır. Elde tutma süresine göre değişen yasal stopaj (vergi) kesinti oranları otomatik uygulanır.
* **Enflasyon İskontosu (Reel Alım Gücü):** Uygulama, belirtilen enflasyon oranını kullanarak paranın zaman değerini (Time Value of Money) formülize eder. Nominal olarak büyüyen kasanın, enflasyon karşısındaki **reel alım gücünü** ikinci bir grafik eğrisi olarak çizer.
* **Dinamik Dönem İçi Nakit Akışları (Vade Bozulması):** Kullanıcı, vade dolmadan para çekme işlemi tanımlayabilir. Algoritma bu durumu "vade bozulması" olarak algılar, o döneme ait birikmiş faizi yakar ve anaparadan düşerek simülasyonu yeni baştan yapılandırır.
* **Yakınsak/Iraksak Seri Analizi:** Çekilen tutarların, tahakkuk eden faizi aşıp aşmadığı matematiksel olarak analiz edilir ve kullanıcının kasasının "sıfıra eriyen (yakınsak)" veya "sonsuza büyüyen (ıraksak)" olduğu raporlanır.

### 📉 2. Borç Ekstresi (Kredi Amortisman Modeli)
Kredi hesaplamaları yasal sınırlamalar ve esnek ödeme planları ile donatılmıştır:
* **BDDK Limitasyon Motoru:** 
  * *İhtiyaç Kredisi:* Talep edilen tutara göre yasal maksimum vade (12, 24 veya 36 ay) kısıtlamalarını otomatik uygular.
  * *Taşıt ve Konut Kredisi:* Girilen teminat (fatura/ekspertiz) değerine göre yasal **LTV (Loan-to-Value)** oranlarını hesaplar. Yasal olarak çekilebilecek maksimum kredi tutarını ve vadeyi kilitler.
* **Erken / Ara Ödeme Optimizasyonu:** Kullanıcı istediği aya özel ekstra ara ödeme tanımlayabilir.
* **Yeniden Yapılandırma Algoritması:** Ara ödeme yapıldığında, kullanıcıya iki seçenek sunulur: Vadeyi erkene çekmek (taksit sabit kalır, borç erken biter) veya anüite formülünü kalan anapara üzerinden yeniden hesaplayarak aylık taksitleri düşürmek (vade sabit kalır).

### 💻 3. Arayüz (UI) ve Raporlama Altyapısı
* **Etkileşimli ve Logaritmik Grafikler:** Native Canvas üzerine inşa edilen grafik motoru, veriler arasında çok büyük uçurumlar olduğunda (örneğin patlayarak büyüyen bir borç) eksenleri otomatik olarak **Logaritmik Ölçeğe (Log-Scale)** çevirir. 
* **Custom Tkinter Bileşenleri:** Özel olarak yazılmış Scrollable Frame'ler, üzerine gelince detay gösteren (Hover) Tooltip'ler ve sayı girerken anlık para birimi formatına (1.000.000,00 ₺) dönüşen akıllı TextBox'lar (CustomEntry).
* **Profesyonel Dışa Aktarım (Zebra Desenli Raporlar):** Oluşturulan ekstreler; veri analizi için **CSV**, profesyonel sunumlar için `openpyxl` ile renkli/zebra desenli **XLSX (Excel)** ve `reportlab` kütüphanesi ile **PDF** formatlarında dışa aktarılabilir.

---

## 🛠️ Kurulum Adımları (Hayatidir)

Projeyi yerel ortamınızda, tüm kütüphane bağımlılıkları çözülmüş şekilde çalıştırmak için aşağıdaki adımları sırasıyla uygulayınız.

### 1. Sistem Gereksinimleri
Sisteminizde **Python 3.7 veya üzeri** bir sürümün yüklü olduğundan emin olun.
Kontrol etmek için terminalinize (macOS/Linux) veya Komut İstemi'ne (Windows) şu komutu yazın:
```bash
python --version
```

### 2. Projeyi Klonlama (Grup_No5)
Projeyi GitHub üzerinden bilgisayarınıza indirin ve proje dizinine gidin:
```bash
git clone [https://github.com/KULLANICI_ADINIZ/Grup_No5.git](https://github.com/KULLANICI_ADINIZ/Grup_No5.git)
cd Grup_No5
```
*(Alternatif olarak yeşil "Code" butonuna tıklayıp "Download ZIP" seçeneğiyle indirebilir ve dosyaları arşivden çıkarabilirsiniz.)*

### 3. Gerekli Kütüphanelerin Yüklenmesi
Simülasyonun çekirdek hesaplama motoru Python standart kütüphaneleriyle çalışır, ancak **Excel ve PDF dışa aktarma (raporlama) özelliklerinin çalışabilmesi için** aşağıdaki paketlerin kurulması **zorunludur**.

Terminal ekranında şu komutu çalıştırın:
```bash
pip install openpyxl reportlab
```

---

## 🚀 Çalıştırma ve Kullanım

Kurulum tamamlandıktan sonra uygulamayı başlatmak için proje dizinindeyken aşağıdaki komutu çalıştırın:

```bash
python "Grup_5_Finalize (1).py"
```

1. **Arayüz Gezinimi:** Üst kısımdaki sekmeleri kullanarak "Tasarruf Planı" veya "Borç Ekstresi" modüllerine geçiş yapabilirsiniz.
2. **Parametre Girişi:** Sol paneldeki alanları doldurun. Hatalı veya yasal sınırı aşan girişlerde sistem sizi (BDDK limit uyarıları vb.) anlık olarak bilgilendirecektir.
3. **Hesapla ve İncele:** "Hesapla" butonuna tıkladığınızda sağ panelde detaylı amortisman tablosunu ve etkileşimli grafiği görebilirsiniz.
4. **Tema Değişimi:** Gece çalışmalarında göz yorgunluğunu önlemek için sağ üstteki **"☀️/🌙 Temayı Değiştir"** butonu ile Karanlık (Dark) moda geçebilirsiniz.
5. **Dışa Aktarma:** "İndir" butonuna basarak sonuçları raporlayabilirsiniz.

---

> **Geliştirici Notu:** Bu proje, Grup No:5 tarafından akademik proje isterlerine ve gerçek hayat senaryolarına uygun olarak geliştirilmiştir. Kod içerisinde yer alan `getcontext().prec = 28` ayarı sayesinde finansal kurumlarda kullanılan hesaplama standartlarına tam uyum sağlanmıştır.
