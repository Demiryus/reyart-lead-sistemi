# Lead Manager — Ekiple Paylaşım (Cloudflare Tunnel)

## Neden Netlify değil
Netlify statik site barındırır; bu sistem Flask (Python sunucu) + diske yazan
companies.json + bu PC'deki Outlook COM mail entegrasyonu. Uygulama bu
makinede kalmalı → çözüm: Cloudflare Tunnel ile linke açmak.

## Şu anki durum: Quick Tunnel (geçici, DEMO)
```powershell
python webapp.py                          # önce webapp
& "C:\Program Files (x86)\cloudflared\cloudflared.exe" tunnel --url http://localhost:5000
```
- Çıktıdaki `https://<rastgele>.trycloudflare.com` linki ekiple paylaşılır.
- ⚠️ Her başlatmada URL DEĞİŞİR ve GİRİŞ KORUMASI YOK (linki bilen girer).
  Uzun süreli ekip kullanımı için aşağıdaki kalıcı kuruluma geç.

## Kalıcı kurulum: Named Tunnel + Access (giriş koruması)
Gereken: ücretsiz Cloudflare hesabı + reyartfuar.com DNS'inin Cloudflare'e
taşınması (nameserver değişikliği, domain sağlayıcı panelinden).

```powershell
cloudflared tunnel login                  # tarayıcıda Cloudflare girişi
cloudflared tunnel create leadmanager
cloudflared tunnel route dns leadmanager leads.reyartfuar.com
cloudflared tunnel run --url http://localhost:5000 leadmanager
```
Sonra Cloudflare dashboard → Zero Trust → Access → Application ekle:
`leads.reyartfuar.com`, policy = e-posta listesi (ekibin mailleri).
Böylece linke giren herkes önce mail doğrulamasından geçer (kodsuz auth).

Windows servisi olarak (PC her açıldığında otomatik):
```powershell
cloudflared service install
```

## Bilinen sınırlar
- PC kapalıyken link çalışmaz (7/24 istenirse VPS planı: bkz. hafıza notu —
  auth kodu + waitress + SQLite gerekir, Outlook mail sunucuda çalışmaz).
- Eşzamanlı yoğun kullanımda JSON yerine SQLite'a geçiş planlanmalı.
