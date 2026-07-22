"""
Önümüzdeki N gün içinde başlayan fuarları TOBB takviminden seçip YENİDEN tarar.

bulk_scrape_calendar.py'den farkı: zaten taranmış fuarları ATLAMAZ — fuar
yaklaştıkça katılımcı listesi büyüdüğü için aynı fuarı tekrar taramak yeni
firmaları yakalar. datastore.append_companies (isim+fuar) dedupe'u sayesinde
mevcut kayıtlar ellenmez, sadece yeni firmalar eklenir.

Mükerrer bucket önleme: TOBB'un uzun resmi adı, companies.json'daki mevcut
fuar adıyla token bazında eşleştirilir (ör. TOBB "Maktek Avrasya 2026 - 9.
Uluslararası..." → mevcut "Maktek Avrasya 2026" bucket'ına yazılır).

Kullanım:
    python rescrape_upcoming.py                  # önümüzdeki 150 gün (sıcak pencere)
    python rescrape_upcoming.py --days 90
    python rescrape_upcoming.py --dry-run        # sadece hedefleri listele
    python rescrape_upcoming.py --refresh        # önce TOBB takvimini güncelle
    python rescrape_upcoming.py --fair maktek    # sadece adı eşleşen fuar(lar)

İlerleme output/rescrape_progress.json'a canlı yazılır. Bitince
build_fair_dates çalıştırılır (yeni fuar bucket'ları tarih alsın diye).
"""

import argparse
import json
import logging
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import scraper
import datastore as ds
import tobb_takvim
import build_fair_dates

log = logging.getLogger(__name__)

OUT_DIR = Path(__file__).parent / "output"
PROGRESS_FILE = OUT_DIR / "rescrape_progress.json"

# Katılımcı listesi yerine haber başlıkları/marketing metni döndüren organizatör
# siteleri — buradan gelen fuarlar TARANMAZ (Tem 2026: İZFAŞ 4 fuara 656 sahte
# "firma" bulaştırdı, elle temizlendi).
SKIP_DOMAINS = ["fuarizmir.com", "izfas"]

# Eşleştirmede anlam taşımayan kelimeler (fuar adlarının hepsinde geçenler)
STOPWORDS = {
    "fuari", "fuar", "uluslararasi", "expo", "ve", "the", "and",
    "eurasia", "avrasya",
}

# Şehir adları: subset eşleşmesine katılmaz ama iki ad da şehir içeriyorsa
# şehirler kesişmeli (Ankara Kitap Fuarı ≠ İstanbul Kitap Fuarı).
CITIES = {
    "istanbul", "izmir", "ankara", "bursa", "antalya", "adana", "konya",
    "gaziantep", "denizli", "diyarbakir", "samsun", "trabzon", "kayseri",
    "eskisehir", "mersin", "malatya", "sivas", "erzurum", "van",
}

_YEAR_RE = re.compile(r"^(19|20)\d{2}$")

# Token eşleşmesinin çalışmadığı bucket'lar (adı tamamen genel kelimelerden
# oluşanlar): normalize TOBB adında substring → mevcut bucket adı.
MANUAL_ALIASES = {
    "f istanbul": "F İstanbul 2026",
}


def _tokens(name: str) -> tuple[set, set, set]:
    """(anlamlı kelimeler, yıllar, şehirler) — Türkçe katlanmış, noktalama temiz."""
    n = re.sub(r"[^\w\s]", " ", scraper._norm(name))
    words, years, cities = set(), set(), set()
    for t in n.split():
        if _YEAR_RE.match(t):
            years.add(t)
        elif t in CITIES:
            cities.add(t)
        elif t not in STOPWORDS and not t.isdigit() and len(t) > 1:
            words.add(t)
    return words, years, cities


def _first_word(name: str) -> str:
    """Adın ilk anlamlı kelimesi (marka adı fuar adının başında olur)."""
    n = re.sub(r"[^\w\s]", " ", scraper._norm(name))
    for t in n.split():
        if (not _YEAR_RE.match(t) and t not in STOPWORDS and t not in CITIES
                and not t.isdigit() and len(t) > 1):
            return t
    return ""


def find_bucket(tobb_name: str, existing: list[str]) -> str | None:
    """TOBB fuar adını companies.json'daki mevcut fuar adıyla eşleştir.
    Kural: mevcut adın tüm anlamlı kelimeleri TOBB adında geçmeli; iki ad da
    yıl içeriyorsa yıllar da kesişmeli (Hometex 2025 ≠ Hometex 2026). Tek
    kelimelik adlar zayıf — o kelime TOBB adının İLK anlamlı kelimesi olmalı
    (marka adları başta gelir; 'Kitap' gibi genel kelimelerin ortada
    geçmesi eşleşme sayılmaz)."""
    tn = scraper._norm(tobb_name)
    for sub, bucket in MANUAL_ALIASES.items():
        if re.search(rf"(?<!\w){re.escape(sub)}(?!\w)", tn) and bucket in existing:
            return bucket
    tw, ty, tc = _tokens(tobb_name)
    best, best_len = None, 0
    for ex in existing:
        ew, ey, ec = _tokens(ex)
        if not ew or not ew <= tw:
            continue
        if ey and ty and not (ey & ty):
            continue
        if ec and tc and not (ec & tc):
            continue
        if len(ew) == 1 and next(iter(ew)) != _first_word(tobb_name):
            continue
        if len(ew) > best_len:
            best, best_len = ex, len(ew)
    return best


def _write_progress(state: dict):
    PROGRESS_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Yaklaşan fuarları yeniden tara")
    parser.add_argument("--days", type=int, default=150, help="bugünden itibaren gün penceresi (varsayılan 150 = sıcak pencere)")
    parser.add_argument("--since", default=None, help="YYYY-MM-DD — pencere başlangıcı (--days yerine)")
    parser.add_argument("--until", default=None, help="YYYY-MM-DD — pencere bitişi (--days yerine)")
    parser.add_argument("--refresh", action="store_true", help="önce TOBB takvimini yeniden çek")
    parser.add_argument("--dry-run", action="store_true", help="taramadan sadece hedef listesini göster")
    parser.add_argument("--fair", default=None, help="sadece adında bu geçen fuar(lar)ı tara")
    args = parser.parse_args()

    if args.refresh:
        log.info("TOBB takvimi yenileniyor...")
        cal = tobb_takvim.scrape_and_save()
    else:
        cal = tobb_takvim.load_cached()
        if not cal["fairs"]:
            cal = tobb_takvim.scrape_and_save()

    today = date.today()
    since = args.since or today.isoformat()
    until = args.until or (today + timedelta(days=args.days)).isoformat()

    targets = [f for f in cal["fairs"] if since <= f["start_date_iso"] <= until]
    targets = [f for f in targets
               if not any(d in (f.get("web") or "").lower() for d in SKIP_DOMAINS)]
    if args.fair:
        needle = scraper._norm(args.fair)
        targets = [f for f in targets if needle in scraper._norm(f["name"])]
    targets.sort(key=lambda f: f["start_date_iso"])

    existing_fairs = sorted({c.get("fair", "") for c in ds.load() if c.get("fair")})

    plan = []
    for f in targets:
        bucket = find_bucket(f["name"], existing_fairs)
        plan.append((f, bucket))

    log.info("Pencere: %s → %s | %d hedef fuar (%d'i mevcut bucket'a eşleşti)",
             since, until, len(plan), sum(1 for _, b in plan if b))

    if args.dry_run:
        for f, bucket in plan:
            mark = f"→ mevcut: {bucket}" if bucket else "→ YENİ"
            print(f"{f['start_date_iso']}  {f['name'][:70]:70s} {mark}")
        return

    state = {
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": "",
        "window": f"{since} → {until}",
        "total": len(plan),
        "done": 0,
        "found_lists": 0,
        "no_list": 0,
        "errors": 0,
        "companies_added": 0,
        "current": "",
        "results": [],  # {name, bucket, start, status, added}
        "finished": False,
    }
    _write_progress(state)

    for i, (f, bucket) in enumerate(plan, 1):
        name, web = f["name"], f["web"]
        fair_name = bucket or name
        state["current"] = f"[{i}/{len(plan)}] {name}"
        _write_progress(state)
        log.info("(%d/%d) %s → bucket: %s", i, len(plan), name, fair_name)

        result = {"name": name, "bucket": fair_name, "start": f["start_date"], "status": "", "added": 0}
        try:
            if not web:
                result["status"] = "website yok"
                state["no_list"] += 1
            else:
                url = web if web.startswith("http") else f"https://{web}"
                companies = scraper.generic_scrape(url, fair_name=fair_name)
                used_url = url
                if not companies:
                    discovered = scraper.discover_participant_url(url)
                    if discovered and discovered != url:
                        companies = scraper.generic_scrape(discovered, fair_name=fair_name)
                        used_url = discovered
                if companies:
                    added = ds.append_companies(companies)
                    result["status"] = f"✅ {len(companies)} firma bulundu, {added} yeni ({used_url})"
                    result["added"] = added
                    state["found_lists"] += 1
                    state["companies_added"] += added
                else:
                    result["status"] = "❌ liste yok/erişilemedi"
                    state["no_list"] += 1
        except Exception as e:
            result["status"] = f"⚠️ hata: {e}"
            state["errors"] += 1
            log.error("Hata (%s): %s", name, e)

        state["done"] = i
        state["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        state["results"].append(result)
        _write_progress(state)
        time.sleep(1.5)  # fuar siteleri arası kibar bekleme

    state["finished"] = True
    state["current"] = ""
    _write_progress(state)
    log.info("BİTTİ: %d fuar denendi, %d liste bulundu, %d YENİ firma eklendi",
             state["total"], state["found_lists"], state["companies_added"])

    if state["companies_added"]:
        log.info("fair_dates.json güncelleniyor (yeni bucket'lar tarih alsın)...")
        build_fair_dates.main()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        datefmt="%H:%M:%S")
    main()
