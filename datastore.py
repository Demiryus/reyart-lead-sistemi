"""
Veri katmanı — companies.json'a güvenli erişim.

- Atomik yazma (geçici dosya + rename): yarıda kesilen yazma veriyi bozamaz.
- Yedek rotasyonu: her save() öncesi otomatik yedek, en yeni BACKUP_KEEP tanesi tutulur.
- Kalıcı ID: her kayda name+fair'den türetilen deterministik kimlik eklenir
  (web arayüzünden güncelleme için).
- Web uygulaması scraping bağımlılıklarını (requests/bs4/playwright) çekmesin
  diye güven seviyesi hesabı burada bağımsız olarak tekrarlanır.
"""

import hashlib
import json
import os
import tempfile
import threading
from datetime import date, datetime
from pathlib import Path

# Süreç içi yazma kilidi: Flask thread'lerinden eşzamanlı iki kaydetme
# birbirinin değişikliğini ezemesin (load-değiştir-save atomik olsun).
_WRITE_LOCK = threading.Lock()

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
COMPANIES_FILE = OUTPUT_DIR / "companies.json"
BACKUP_KEEP = 10

# Arayüzden düzenlenebilecek alanlar — bunun dışındakiler reddedilir.
EDITABLE_FIELDS = {
    "status", "note", "phone", "email", "website", "linkedin", "priority",
    "takip_tarihi",  # YYYY-MM-DD — "bu firmayı şu gün tekrar ara" hatırlatması
}

STATUS_CHOICES = [
    "⬜ Aranmadı", "📞 Arandı", "✉️ Mail Atıldı", "🤝 Görüşülüyor",
    "✅ Anlaşıldı", "❌ İlgilenmiyor",
]


def make_id(company: dict) -> str:
    key = (company.get("name", "") + "|" + company.get("fair", "")).encode("utf-8")
    return hashlib.md5(key).hexdigest()[:12]


# "Türkiye"/"Turkey" gibi yazımların hepsini yerli sayan alias seti — maps_worker.py'deki
# aynı mantık (country boşsa da yerli varsayılır: çoğu eski kayıt Türk firması).
TR_ALIASES = {"turkiye", "türkiye", "turkey", "tr", ""}


def _fold_tr(s: str) -> str:
    table = str.maketrans({
        "İ": "i", "I": "i", "ı": "i", "Ş": "s", "ş": "s", "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u", "Ö": "o", "ö": "o", "Ç": "c", "ç": "c",
    })
    return s.translate(table)


def is_foreign(c: dict) -> bool:
    country = _fold_tr((c.get("country") or "").strip()).lower()
    return country not in TR_ALIASES


def compute_mensei(c: dict) -> str:
    return "🌍 Yabancı" if is_foreign(c) else "🇹🇷 Yerli"


def compute_guven(c: dict) -> str:
    site_ok = str(c.get("site_dogrulama", "")).startswith("✅")
    tel_ok = str(c.get("tel_dogrulama", "")).startswith("✅")
    if site_ok and tel_ok:
        return "🟢 Yüksek"
    if site_ok or tel_ok:
        return "🟡 Orta"
    return "⚪ Belirsiz"


# ── Satış sıcaklığı / skoru ──────────────────────────────────────────────────
# Reyart stand tasarımı satıyor → müşterinin stand kararı fuardan ~1-5 ay önce
# verilir. Fuarı 0-150 gün içinde olan firma "🔥 Sıcak" (satış penceresi açık),
# daha uzaktaki "🌡 Yaklaşıyor", geçmiş/tarihi bilinmeyen "❄ Soğuk".

FAIR_DATES_FILE = OUTPUT_DIR / "fair_dates.json"
_FAIR_DATES: dict = {}
_FAIR_DATES_MTIME: float = -1.0

HOT_WINDOW_DAYS = 150


def _fair_dates() -> dict:
    """fair_dates.json'ı mtime kontrolüyle önbellekler (her istekte diskten
    okumamak için). Dosya build_fair_dates.py ile üretilir/yenilenir."""
    global _FAIR_DATES, _FAIR_DATES_MTIME
    try:
        mtime = FAIR_DATES_FILE.stat().st_mtime
    except OSError:
        return {}
    if mtime != _FAIR_DATES_MTIME:
        try:
            _FAIR_DATES = json.loads(FAIR_DATES_FILE.read_text(encoding="utf-8"))
            _FAIR_DATES_MTIME = mtime
        except (json.JSONDecodeError, OSError):
            return _FAIR_DATES
    return _FAIR_DATES


def compute_sicaklik(fair_start_iso: str, today_iso: str) -> tuple[str, int]:
    """(etiket, taban_puan) döndürür."""
    if not fair_start_iso:
        return "❄ Soğuk", 0
    if fair_start_iso < today_iso:
        return "❄ Soğuk", 0
    days = (date.fromisoformat(fair_start_iso) - date.fromisoformat(today_iso)).days
    if days <= HOT_WINDOW_DAYS:
        return "🔥 Sıcak", 100
    return "🌡 Yaklaşıyor", 50


def compute_satis_skoru(c: dict, base: int) -> int:
    score = base
    prio = c.get("priority", "")
    if prio == "⭐⭐⭐":
        score += 30
    elif prio == "⭐⭐":
        score += 15
    if is_foreign(c):
        score += 20  # yabancı katılımcının yerel stand partneri yok — en değerli segment
    if c.get("phone"):
        score += 5
    if c.get("email"):
        score += 5
    return score


# load() önbelleği: 9600+ kayıtlık JSON'u her API isteğinde yeniden parse edip
# alanları yeniden hesaplamamak için. Anahtar dosya mtime+boyut (değişince
# otomatik tazelenir) + bugünün tarihi (sıcaklık/skor güne bağlı) + fair_dates
# mtime'ı. update_company kaydettiğinde mtime değişir → sonraki load tazelenir.
_LOAD_CACHE: dict = {"key": None, "data": None}


def load() -> list[dict]:
    if not COMPANIES_FILE.exists():
        return []
    st = COMPANIES_FILE.stat()
    try:
        fd_mtime = FAIR_DATES_FILE.stat().st_mtime_ns
    except OSError:
        fd_mtime = 0
    today_iso = date.today().isoformat()
    key = (st.st_mtime_ns, st.st_size, today_iso, fd_mtime)
    if _LOAD_CACHE["key"] == key:
        return _LOAD_CACHE["data"]

    data = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    fair_dates = _fair_dates()
    changed = False
    for c in data:
        if "id" not in c:
            c["id"] = make_id(c)
            changed = True
        c["guven_seviyesi"] = compute_guven(c)
        c["mensei"] = compute_mensei(c)
        fair_date = fair_dates.get(c.get("fair", ""), "")
        c["fuar_tarihi"] = fair_date
        sicaklik, base = compute_sicaklik(fair_date, today_iso)
        c["sicaklik"] = sicaklik
        c["satis_skoru"] = compute_satis_skoru(c, base)
    if changed:
        save(data, backup=False)  # sadece id eklendi, yedek şart değil
    _LOAD_CACHE["key"] = key
    _LOAD_CACHE["data"] = data
    return data


def _rotate_backups():
    backups = sorted(
        OUTPUT_DIR.glob("companies_backup_*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for old in backups[BACKUP_KEEP:]:
        try:
            old.unlink()
        except OSError:
            pass


def save(data: list[dict], backup: bool = True):
    OUTPUT_DIR.mkdir(exist_ok=True)
    if backup and COMPANIES_FILE.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = OUTPUT_DIR / f"companies_backup_{ts}.json"
        backup_path.write_bytes(COMPANIES_FILE.read_bytes())
        _rotate_backups()

    payload = json.dumps(data, ensure_ascii=False, indent=2)
    # Atomik yazma: aynı dizinde geçici dosyaya yaz, sonra rename et
    fd, tmp_path = tempfile.mkstemp(
        dir=str(OUTPUT_DIR), prefix=".companies_", suffix=".tmp"
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(payload)
        os.replace(tmp_path, COMPANIES_FILE)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def update_company(company_id: str, fields: dict) -> dict | None:
    """Tek firmayı günceller; başarılıysa güncel kaydı döndürür, yoksa None.
    Yazma kilidi altında çalışır: iki satışçı aynı anda farklı firmaları
    kaydederse ikisi de korunur (son okuyan diğerinin yazdığını ezmez)."""
    bad = set(fields) - EDITABLE_FIELDS
    if bad:
        raise ValueError(f"Düzenlenemez alan(lar): {', '.join(sorted(bad))}")
    with _WRITE_LOCK:
        data = load()
        target = next((c for c in data if c.get("id") == company_id), None)
        if target is None:
            return None
        for k, v in fields.items():
            target[k] = str(v).strip()
        target["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save(data)
    target["guven_seviyesi"] = compute_guven(target)
    return target


def append_companies(new: list[dict]) -> int:
    """Yeni firmaları (ör. yeni kazınan fuar) mevcut veriye kilit altında ekler.
    Aynı isim+fuar zaten varsa atlanır. Eklenen sayısını döndürür."""
    with _WRITE_LOCK:
        data = load()
        seen = {(c.get("name", "").lower().strip(), c.get("fair", "")) for c in data}
        added = 0
        for c in new:
            key = (c.get("name", "").lower().strip(), c.get("fair", ""))
            if key in seen or not key[0]:
                continue
            c["id"] = make_id(c)
            data.append(c)
            seen.add(key)
            added += 1
        if added:
            save(data)
    return added


def data_version() -> str:
    """Verinin sürüm damgası — istemciler bunu poll edip değişince yenilenir."""
    try:
        st = COMPANIES_FILE.stat()
        return f"{st.st_mtime_ns}-{st.st_size}"
    except OSError:
        return "0"


def stats(data: list[dict] | None = None) -> dict:
    if data is None:
        data = load()
    fairs: dict[str, dict] = {}
    for c in data:
        f = c.get("fair", "") or "(Fuar Belirtilmemiş)"
        s = fairs.setdefault(f, {
            "total": 0, "web": 0, "tel": 0, "mail": 0, "linkedin": 0,
            "yuksek": 0, "orta": 0, "belirsiz": 0, "yerli": 0, "yabanci": 0,
            "sicak": 0,
        })
        s["total"] += 1
        if c.get("website"):
            s["web"] += 1
        if c.get("phone"):
            s["tel"] += 1
        if c.get("email"):
            s["mail"] += 1
        if c.get("linkedin"):
            s["linkedin"] += 1
        g = c.get("guven_seviyesi", "")
        if "🟢" in g:
            s["yuksek"] += 1
        elif "🟡" in g:
            s["orta"] += 1
        else:
            s["belirsiz"] += 1
        if is_foreign(c):
            s["yabanci"] += 1
        else:
            s["yerli"] += 1
        if "🔥" in c.get("sicaklik", ""):
            s["sicak"] += 1
    total = {
        k: sum(s[k] for s in fairs.values())
        for k in ("total", "web", "tel", "mail", "linkedin", "yuksek", "orta",
                  "belirsiz", "yerli", "yabanci", "sicak")
    }
    return {"fairs": fairs, "total": total}
