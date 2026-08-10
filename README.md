# GoPhish Reporting Tool 🎣

Modern, tam donanımlı bir **GoPhish phishing simülasyonu raporlama motoru**. Güvenlik ekibi ve kırmızı takım ekipleri için GoPhish REST API'sine bağlanarak kampanya verilerini otomatik çeker; analiz, risk hesaplama ve bölümlemeyi yapar; müşteriye / yönetime sunulabilecek profesyonel **Microsoft Word (`.docx`)** ve **Excel (`.xlsx`)** raporlarını saniyeler içinde üretir.

Tamamen **GUI tabanlı**, modern CustomTkinter arayüzü ile çalışır; komut satırı bilgisi gerektirmez. Türkçe ve İngilizce dil desteği içerir.

---

## 📌 Executive Summary (Neden Gerekli?)

Kurumsal phishing değerlendirmelerinde kampanya istatistiklerini el ile çekmek, süreleri ve olayları sınıflandırmak ve müşteriye sunulabilir dokümanlar hâline getirmek **yavaş** ve **hataya açık** bir süreçtir.

Bu araç; GoPhish API'sinden ham olay verisini programatik olarak çeker, hedef telemetrisini işler (gönderilen, açılan, bağlantıya tıklanan, kimlik bilgisi girilen) ve **sunum kalitesinde** Word + Excel raporları üretir.

---

## 🚀 Öne Çıkan Özellikler

### 🔌 Otomatik API Veri Çekme
- Aktif veya arşivlenmiş GoPhish sunucuna **REST API** üzerinden güvenli bağlantı.
- **Timeout + Retry katmanı** (30 sn, 3 deneme, exponential backoff).
- Sunucu/API anahtarı ilk kullanımda sorulur, yerelde `.gophish_config.json` içinde güvenle saklanır.
- **API Health** kontrolü: ping, gecikme ölçümü ve bağlantı durumu canlı takip.
- **Akıllı cache**: Kampanya verisi JSON cache (TTL 5 dk) ile tekrar eden isteklerde hız kazanır.

### 🖥️ Modern GUI (CustomTkinter)
- 4 ana sekme: **Dashboard**, **Kampanyalar**, **Rapor Tasarımı**, **Ayarlar**.
- **TR / EN** dil desteği (anlık çeviri).
- **System / Dark / Light** tema seçenekleri; otomatik font seçimi (Inter, Roboto, Segoe UI vb.).
- Windows'ta **sürükle-bırak** çıktı klasörü seçimi.
- Toast bildirimleri, yükleme spinner'ı, iptal edilebilir progress takibi.
- Bağımlılıklar **otomatik kurulur** (ilk çalıştırmada).

### 📊 Ölçülen Metrikler & Telemetri
Kampanyanın tüm yaşam döngüsü yakalanır ve kategorize edilir:

| Metric Stage | Description | Risk Etkisi |
| :--- | :--- | :--- |
| **Gönderilen E-posta** | Sunucu tarafından gönderilen toplam e-posta sayısı. | Başlangıç |
| **Açılan E-posta** | Takipçi pikselini yükleyen hedefler. | Düşük |
| **Tıklanan Bağlantı** | E-postadaki phishing URL'sini tıklayan hedefler. | Orta |
| **Veri Giren** | Kimlik bilgisi / form dolduran hedefler. | **Kritik** |
| **Hata / Geri Dönme** | SMTP filtresi veya geçersiz adres nedeniyle başarısız dispatch. | Bilgi |

### 📈 Analiz & Akıllı İstatistik
- **Risk Skoru** otomatik hesaplanır: `Submit × 3 + Click × 2 + Open × 1` → yüzdelik risk oranı.
- **Çoklu kampanya** seçimi ve birleştirme, **trend analizi**, kampanya kıyaslaması.
- **Dashboard** kartları: toplam kampanya, hedef, açan, tıklayan, veri giren ve % risk.
- **Domain dağılımı** sayfası (Excel) — hangi kurum-domaininden ne kadar etkilendiği.

### 📄 Rapor Üretimi
- **Microsoft Word (`.docx`)** — yönetici özeti, pasta grafiği (matplotlib), istatistik tablosu, kritik/orta risk detay tabloları.
- **Excel (`.xlsx`)** — çoklu liste sayfası (gönderilen, açan, tıklayan, veri giren), eylem saatleri, payload, **dinamik form alanları**, durum (Geçerli/Şüpheli/Troll), domain dağılımı.
- **Cihaz / Tarayıcı ayrımı** — işletim sistemi sürümü (Windows/iOS/Android/Mac/Linux) ve tarayıcı + sürümü (`Chrome 150`, `Safari 26` vb.) ayrı sütunlarda raporlanır.
- Otomatik **dosya adı çakışma** yönetimi (`_1`, `_2` eklenir).

### 🛡️ OPSEC & Veri Gizliliği
- **Payload maskeleme**: parola gibi hassas alanlar raporda `***` olarak gösterilir.
- **Şüpheli paylaşım tespiti**: kullanıcının yazdığı veride hassas ifadeler denetlenir.
- Kapsamlı **`.gitignore`** ile API token, hedef e-posta adresleri, config dosyaları, loglar ve üretilen rapor dosyalarının repo'ya sızması engellenir.

---

## 📁 Uygulama Yapısı (Mimari)

> Tek dosyalık, modüler ve kolay genişletilebilir yapı.

| Bileşen | Açıklama |
| :--- | :--- |
| `GoPhishApp` | Ana CustomTkinter uygulaması, sekmeler ve arayüz. |
| `GophishApiClient` | API çağrıları, timeout/retry & hata yönetimi. |
| `ApiHealth` | Sağlık kontrolü & gecikme ölçümü. |
| `AppCache` | Kampanya verisi TTL'li JSON önbelleği. |
| `ReportOptions` | Word/Excel rapor seçeneklerini yapılandıran veri modeli. |
| `ProgressReporter` | İlerleme / iptal bildirimi. |
| `ToastManager`, `LoadingSpinner`, `AppFonts`, `IconManager` | UI yardımcıları. |
| `_MergedCampaign` | Çoklu kampanya birleştirme modeli. |
| `create_table`, `generate_reports` | Word/Excel rapor üretim motoru. |

---

## 📋 Gereksinimler

### Sistem
- **Windows 10/11** (tam destek, sürükle-bırak + windnd)
- macOS / Linux (GUI desteklenir; sürükle-bırak isteğe bağlı)
- **Python 3.9+**

### Gerekli Kütüphaneler
Kurulum **otomatik** yapılır (ilk çalıştırmada). Elle kurmak isterseniz:

```bash
pip install gophish python-docx matplotlib openpyxl urllib3 customtkinter Pillow
# Windows sürükle-bırak desteği için:
pip install windnd
```

---

## 🚀 Kullanım

1. **Çalıştır**:
   ```bash
   python gophish_report_tool_.py
   ```
   İlk açılışta eksik paketler otomatik kurulur.
2. **Kampanyalar** sekmesinde GoPhish sunucu adresini ve **API key**'i gir, `Bağlan`'a bas. (URL ve key bir kez girilir, sonra `.gophish_config.json`'dan hatırlanır.)
3. İstediğin kampanyaları seç (çoklu seçim mümkün), **Rapor Tasarımı** sekmesine geç.
4. Çıktı klasörünü seç (sürükle-bırak), Word / Excel seçeneklerini işaretle, oluştur'a bas.
5. Raporlar seçtiğin klasöre kaydedilir.

Bağlantı sorunlarında **Dashboard** sekmesindeki API health ve gecikme göstergesinden durumu kontrol edebilirsin.

---

## ⚙️ Konfigürasyon & Saklama

| Dosya | İçerik | Git'ten gizli mi? |
| :--- | :--- | :---: |
| `.gophish_config.json` | Sunucu URL + API anahtarı | ✅ `.gitignore` |
| `cache/` | Kampanya önbelleği | ✅ |
| `*.log`, `GOPHISH_HATA_LOG.txt` | Çalışma / hata logları | ✅ |
| `*_Rapor` , `Trend_Analizi*`, `*.docx`, `*.xlsx` | Üretilen raporlar | ✅ |

> **Güvenlik:** API anahtarı ve konfigürasyon dosyaları `.gitignore` ile korunur; repo'ya asla sızmasın.

---

## 🔧 Sorun Giderme

- **Paket kurulumu başarısız**: Python'un PATH'te olduğundan emin ol, terminal'i yönetici olarak çalıştır.
- **API bağlantı hatası**: Sunucu URL'nin `/api/v3` içerdiğini ve API key'in geçerli olduğunu doğrula; güvenlik duvarını kontrol et.
- **Grafik boş çıkıyor**: Kampanya için `Submitted` verisi yoksa grafik boş kalabilir (bu normal).
- **Hata logları**: `GOPHISH_HATA_LOG.txt` içinde hata detaylarını bulabilirsin.

---

## 🧩 Genişletilebilirlik

Kod; marka/kurumsal rapor şablonları, özel istatistik fonksiyonları veya ek veri görselleştirme kütüphaneleri ekleyecek şekilde modüler tasarlanmıştır. İstatistik mantığı `generate_reports` içinde, arayüz `GoPhishApp` içinde izole edilmiştir.

---

## ⚖️ Yasal Uyarı

Bu araç yalnızca **yetki verilmiş** phishing simülasyonları ve güvenlik değerlendirmeleri için tasarlanmıştır. İzinsiz kullanım ilgili kanunlara aykırı olabilir. Kullanıcı, aracı yalnızca sahibi olduğu / yetkilendirildiği sistemlerde kullanmalıdır.
