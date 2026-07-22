# Reyart Lead Scraper — CLAUDE.md

## Görev
KOMATEK 2026 ve WIN EURASIA 2026 katılımcı firmalarının iletişim bilgilerini (telefon, e-posta, website, LinkedIn) topla ve Excel'e dökmek.

## Kurulum
```bash
pip install playwright beautifulsoup4 openpyxl requests lxml
playwright install chromium
```

## Proje Yapısı
```
reyart-scraper/
├── CLAUDE.md
├── scraper.py         # Ana scraper
├── enricher.py        # Her firma için iletişim bilgisi bulucu
├── output/
│   └── leads.xlsx     # Çıktı
```

## Adım 1 — scraper.py yaz

KOMATEK listesini çek:
- URL: https://komatekfuar.com/en/list-of-komatek-2026-participants/
- Tüm firma isimlerini parse et (li elementleri)
- Her firma için: isim, menşei tahmini, sektör

WIN EURASIA listesi için:
- URL: https://platform.win-eurasia.com/participants?new
- Sayfada JavaScript var, Playwright kullan
- Firma adı, ülke, ürün grubu bilgilerini çek

## Adım 2 — enricher.py yaz

Her firma için şunları bul:
1. Google'da "{firma adı} Turkey telefon" ara
2. Firma web sitesine gir
3. /contact, /iletisim, /tr/iletisim sayfalarını dene
4. Regex ile telefon numarası çek: +90, 0212, 0216, 0850 formatları
5. Regex ile e-posta çek: @domain.com formatı
6. LinkedIn şirket sayfası URL'si bul

Rate limiting: Her istek arası 1-2 saniye bekle, ban yeme.
User-Agent: Gerçekçi browser header kullan.

## Adım 3 — Excel çıktısı (openpyxl)

Kolonlar:
| Firma Adı | Menşei | Sektör | Fuar | Website | Telefon | E-posta | LinkedIn | Öncelik | Durum | Not |

Öncelik otomatik ata:
- "⭐⭐⭐" = CAT, Liebherr, Komatsu, Hyundai, Bobcat, HİDROMEK, Sandvik, Metso, DEVELON, NEW HOLLAND, CASE, Manitou, Kobelco, Jungheinrich, TRUMPF gibi global markalar
- "⭐⭐" = Orta büyüklükte uluslararası markalar
- "⭐" = Yerli veya küçük firmalar

Renk kodlaması:
- ⭐⭐⭐ satırlar: altın sarısı arka plan (#FFD700)
- ⭐⭐ satırlar: açık mavi (#D9E8F6)
- ⭐ satırlar: beyaz

Durum kolonu default: "⬜ Aranmadı"

## Adım 4 — İkinci kaynak: Geçmiş fuarlar

E�er WIN EURASIA 2026 listesi çekilemezse:
- https://www.expointurkey.org/win-eurasia-2026 sayfasını dene
- https://www.tradefairdates.com/WIN-EURASIA sayfasını dene
- 2024 WIN EURASIA katılımcılarını bul (büyük ihtimalle 2026'da da varlar)

## Hata yönetimi
- Her firma için try/except kullan, hata olursa atla devam et
- Hangi firmadan veri alınamadığını logla
- Çalışma sonunda özet: "X firmadan Y'sinin iletişim bilgisi bulundu"

## Çalıştırma
```bash
python scraper.py      # Önce listeyi çek
python enricher.py     # Sonra iletişim bilgilerini zenginleştir
python webapp.py       # Web arayüz → http://127.0.0.1:5000
```

Çıktı: output/leads.xlsx — fuar başına ayrı sheet + Özet sekmesi.

## Web arayüz (webapp.py + datastore.py + templates/index.html)
- Dashboard kartları (fuar bazında sayılar, güven seviyeleri), filtreler
  (fuar/öncelik/güven/durum/eksik-veri), TR-katlamalı arama, sayfalama.
- Satıra tıkla → detay paneli: telefon/mail/website/LinkedIn/öncelik/durum/not
  düzenlenebilir; kayıtlar atomik yazma + otomatik yedek rotasyonuyla saklanır.
- "Excel İndir" → zaman damgalı export (leads.xlsx kilitliyken de çalışır).

## Mail sistemi (mailer.py + mail_templates.json)
- Gönderim **classic Outlook** üzerinden COM otomasyonuyla (pywin32) yapılır —
  SMTP şifresi gerekmez. Outlook açık ve `demir@reyartfuar.com` hesabı ekli
  olmalı (webapp.py'nin çalıştığı aynı Windows makinesinde).
- Hesap/ayarlar opsiyonel `mail_config.json`'dan gelir (`outlook_account`,
  `sender_name`, `daily_limit`); yoksa demir@reyartfuar.com varsayılan.
- Aynı firma+şablon ikinci kez gönderilmez (output/mail_log.jsonl'dan kontrol).
- Arayüzde: firma seç → "Seçilenlere Mail" → şablon seç → önizle → Kuru
  Çalıştırma veya Gerçek Gönder. Tekli mail için detay panelinde "Mail Yaz
  (Outlook)" var — bilerek `mailto:` KULLANMIYOR (Windows'un varsayılan mail
  uygulamasına, ör. yanlış hesaplı yeni Outlook'a gidiyordu); bunun yerine
  `/api/mail/compose` COM ile doğrudan classic Outlook'ta doğru hesapla
  taslak açar (göndermez, kullanıcı gözden geçirip kendi gönderir).

## Menşei (Yerli/Yabancı) ayrımı
- Her kayıt `country` alanından otomatik `mensei` hesaplar (`datastore.is_foreign()` /
  `compute_mensei()`): Türkiye/Turkey/TR/boş → 🇹🇷 Yerli, diğerleri → 🌍 Yabancı.
- Arayüzde filtre (`Tüm Menşeiler / Yerli / Yabancı`) + dashboard kartı + tablo rozeti.
- Tekli "Mail Yaz (Outlook)" butonu yabancı firmalarda otomatik `tanitim_en`, yerlilerde
  `tanitim_tr` şablonunu seçer.

## Gelecek Fuarlar sekmesi (tobb_takvim.py)
- Kaynak: https://fuarlar.tobb.org.tr/FuarTakvimi — Blazor Server ama tablo ilk HTML
  yanıtında dolu geliyor (prerender), Playwright gerekmiyor, düz requests+BS4 yeterli.
- `output/tobb_fuar_takvimi.json`'a cache'lenir; `/api/calendar` cache'i döner (yoksa
  otomatik çeker), `/api/calendar/refresh` TOBB'dan yeniden çeker.
- Arayüzde "📅 Gelecek Fuarlar" sekmesi: şehir filtresi, arama, "sadece gelecek
  fuarlar" (bitiş tarihi bugünden ileri) checkbox'ı, "TOBB'dan Güncelle" butonu.
- Her satırda "➕ Ekle" butonu — fuar adı + organizatör websitesiyle "Fuar Ekle"
  modalını önceden doldurur, oradan direkt kazımaya bağlanır (aşağıya bkz.).

## generic_scrape() — herhangi bir fuar sitesini tarama (scraper.py)
- Artık tek sayfa/tek strateji ile sınırlı değil: UL/class-selector/tablo'ya ek
  olarak tekrar eden div/article "kart" bloklarını da tanıyor (WordPress/React
  grid düzenleri), her adayı unique-ratio kalite kapısından geçiriyor (yanlış
  seçici → aynı metin defalarca sorunu böyle elenir).
- `?page=N` tarzı sayfalamayı otomatik takip eder (2 sayfa üst üste yeni kayıt
  gelmeyince durur) — tek sayfa sanıp yarım veri toplama riski kalktı.
- Tüyap-CMS'i (`div.brand-item`) otomatik tanır ve `tuyap_platform_scrape.py`'nin
  doğrulanmış parser'ını kullanır (bu platformda onlarca fuar sitesi var).
- Ağ hatalarına karşı retry+backoff, charset otomatik düzeltme (apparent_encoding
  — server charset bildirmeyince Türkçe karakter bozulmasın diye), login-duvarı
  tespiti (sahte/boş veri üretmek yerine net uyarı loglar).
- `discover_participant_url(homepage_url)`: verilen URL doğrudan katılımcı listesi
  değilse (ör. organizatörün ana sayfası), sayfadaki linklerden "katılımcı listesi
  / exhibitor list" gibi anahtar kelimelerle eşleşeni otomatik bulur. "Fuar Ekle"
  akışı (webapp.py `_run_scrape_job`) önce verilen URL'i dener, boş dönerse bunu
  otomatik devreye sokar — kullanıcı çoğu zaman sadece ana sayfa yapıştırsa da çalışır.
- Strateji 5 ("p>strong lead"): bazı fuar siteleri (ör. f-istanbul.com) katılımcıları
  ul/table/kart değil, düz tanıtım metni olarak yazıyor — her firma kendi `<p>`'sinde,
  adı paragrafın en başında `<strong>`/`<b>` ile geçiyor. Bu deseni de tanıyor.
- Bilinen bir yanlış-pozitif kaynağı düzeltildi: "ziyaretçi geldiği ülkeler" bayrak
  grid'i (`div.flag` gibi tekrar eden bloklar) firma listesi sanılıp F İstanbul'da
  18 ülke adını companies.json'a yanlışlıkla eklemişti — hem `class` adında
  flag/country/ulke geçen blokları hariç tutan bir filtre hem de `COUNTRY_NAMES`
  junk seti eklendi (defense-in-depth, ikisi birden).

## Satış motoru (fuar tarihi → sıcaklık → skor)
- `build_fair_dates.py` → `output/fair_dates.json`: her fuar adına başlangıç
  tarihi eşler (TOBB takvimi cache'inden otomatik + elle eklenenler için
  MANUAL_DATES sözlüğü). TOBB takvimi yenilendiğinde tekrar çalıştır.
- `datastore.load()` her kayda hesaplar: `fuar_tarihi`, `sicaklik`
  (🔥 Sıcak = fuara ≤150 gün — stand kararı bu pencerede verilir /
  🌡 Yaklaşıyor / ❄ Soğuk-geçmiş-bilinmiyor) ve `satis_skoru`
  (sıcaklık tabanı + ⭐'lar + yabancı +20 + tel/mail +5'er).
- `takip_tarihi` düzenlenebilir alan (YYYY-MM-DD): "bu gün tekrar ara".
  Arayüzde "📞 Bugün Aranacaklar" butonu takip tarihi gelmiş kayıtları süzer.
- API: `/api/companies?sicaklik=🔥&sort=skor&takip=bugun`; stats'ta `sicak` sayısı.
- Arayüz: 🔥 SICAK LEAD kartı (tıklayınca filtre+skor sıralaması), sıcaklık
  filtresi/rozeti, panelde satış durumu + takip tarihi. Dashboard fuar kartları
  en kalabalık 11 fuarla sınırlı (164 fuar kartı sayfayı boğuyordu).
- `hot_leads_excel.py` → `output/SICAK_ARAMA_LISTESI.xlsx`: "🔥 Bugün Ara"
  (telefonlu sıcaklar, skor sıralı), "🌍 Yabancı Sıcak", "Fuar Özeti" sheet'leri.

## Toplu veri kaynakları
- `bulk_scrape_calendar.py`: TOBB takvimindeki tüm gelecek fuarları dener
  (generic_scrape + discover), ilerleme `output/bulk_scrape_progress.json`.
  Temmuz-Aralık 2026 koşusu: 234 fuar → 165'inde liste bulundu, 3836 firma.
- `rescrape_upcoming.py`: önümüzdeki N gün içinde başlayan fuarları YENİDEN
  tarar (`--days 150` varsayılan, `--dry-run`, `--refresh` TOBB güncelle,
  `--fair maktek` tek fuar). bulk'tan farkı: taranmışları atlamaz — fuar
  yaklaştıkça büyüyen listelerden yeni firmaları yakalar (isim+fuar dedupe
  mevcut kayıtları korur). TOBB uzun adını mevcut bucket'a token eşleşmesiyle
  bağlar (yıl + şehir uyumu şart; tek kelimelik ad ilk kelime olmalı;
  MANUAL_ALIASES: "f istanbul"). İlerleme: output/rescrape_progress.json.
  Bitince build_fair_dates otomatik çalışır. Fuar sezonu öncesi ayda bir koş.
- `import_listeler.py`: vault'un `Listeler/` klasöründeki hazır katılımcı
  Excel'lerini aktarır (12 dosya işlendi; Intermob/KOMATEK/F İstanbul için
  merge modu — eksik alan tamamlar, AÇIKLAMA sütunundaki arama notlarını
  duruma çevirir: ULAŞAMADIM→Arandı, ANLAŞMIŞLAR→Anlaşıldı).

## Bilinen durumlar
- WIN EURASIA 2026 listesi login duvarı arkasında (platform.win-eurasia.com,
  B2B girişi gerekiyor) — eski fallback junk üretiyordu, kaldırıldı. Liste için
  platforma üye girişiyle manuel export gerekir.
- Doğrulama sistemi: her kayıtta site_dogrulama/tel_dogrulama + güven seviyesi
  (🟢 site+tel iki kaynaktan doğrulandı / 🟡 tek kaynak / ⚪ belirsiz).
