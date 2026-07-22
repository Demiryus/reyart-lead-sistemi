"""
TOBB Fuar Takvimi'ndeki tüm gelecek fuarları (bugünden verilen bitiş tarihine
kadar) tek tek dener: generic_scrape + discover_participant_url ile katılımcı
listesi bulunabilenleri companies.json'a ekler. Bulunamayanları/login-duvarlı
olanları atlar, hiçbir sahte veri üretmez.

Uzun sürer (yüzlerce site) — arka planda çalıştırılmak üzere tasarlandı.
İlerleme output/bulk_scrape_progress.json'a canlı yazılır (durum raporu için).

Kullanım:
    python bulk_scrape_calendar.py --until 2026-12-31
"""

import argparse
import json
import logging
import time
from datetime import datetime
from pathlib import Path

import scraper
import datastore as ds
import tobb_takvim

log = logging.getLogger(__name__)

OUT_DIR = Path(__file__).parent / "output"
PROGRESS_FILE = OUT_DIR / "bulk_scrape_progress.json"

# Bu fuarların katılımcı listesi zaten farklı bir isimle companies.json'da var —
# TOBB'un uzun resmi adıyla tekrar taranıp mükerrer fuar bucket'ı oluşturmasın.
# (5'i web-scrape ile, 12'si kullanıcının Listeler/ klasörüne koyduğu hazır
# Excel'lerden import_listeler.py ile eklendi — bkz. CLAUDE.md.)
ALREADY_COVERED_SUBSTRINGS = [
    "f istanbul", "intermob", "maden turkiye", "komatek", "win eurasia",
    "automechanika", "aymod", "food expo", "hightex", "hometex",
    "pencere", "kapi", "maktek",
    # not: "mermer" bilerek eklenmedi — Listeler/izmir mermer fuar.xlsx TOBB
    # takviminde eşleşmiyor (İzmir Mermer, TOBB'da yalnızca "Afyonkarahisar Blok
    # Mermer Fuarı" olarak farklı bir fuar için geçiyor; genel "mermer" ile
    # hariç tutmak o farklı fuarı da yanlışlıkla atlatırdı).
]


def _norm(s: str) -> str:
    return scraper._norm(s)


def _already_covered(fair_name: str) -> bool:
    n = _norm(fair_name)
    return any(sub in n for sub in ALREADY_COVERED_SUBSTRINGS)


def _write_progress(state: dict):
    PROGRESS_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--until", default="2026-12-31", help="YYYY-MM-DD — bu tarihe kadar olan fuarlar")
    parser.add_argument("--since", default=None, help="YYYY-MM-DD — varsayılan bugün")
    args = parser.parse_args()

    since = args.since or datetime.now().strftime("%Y-%m-%d")
    until = args.until

    cal = tobb_takvim.load_cached()
    if not cal["fairs"]:
        cal = tobb_takvim.scrape_and_save()

    targets = [f for f in cal["fairs"] if since <= f["start_date_iso"] <= until]
    targets = [f for f in targets if not _already_covered(f["name"])]
    targets.sort(key=lambda f: f["start_date_iso"])

    state = {
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": "",
        "total": len(targets),
        "done": 0,
        "found_lists": 0,
        "no_list": 0,
        "errors": 0,
        "companies_added": 0,
        "current": "",
        "results": [],  # {name, sehir, status, count}
        "finished": False,
    }
    _write_progress(state)
    log.info("Toplam %d fuar denenecek (%s - %s)", len(targets), since, until)

    for i, f in enumerate(targets, 1):
        name, web, sehir = f["name"], f["web"], f["sehir"]
        state["current"] = f"[{i}/{len(targets)}] {name}"
        _write_progress(state)
        log.info("(%d/%d) %s — %s", i, len(targets), name, web)

        result = {"name": name, "sehir": sehir, "start": f["start_date"], "status": "", "count": 0}
        try:
            if not web:
                result["status"] = "website yok"
                state["no_list"] += 1
            else:
                url = web if web.startswith("http") else f"https://{web}"
                companies = scraper.generic_scrape(url, fair_name=name)
                used_url = url
                if not companies:
                    discovered = scraper.discover_participant_url(url)
                    if discovered and discovered != url:
                        companies = scraper.generic_scrape(discovered, fair_name=name)
                        used_url = discovered
                if companies:
                    added = ds.append_companies(companies)
                    result["status"] = f"✅ bulundu ({used_url})"
                    result["count"] = len(companies)
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
    log.info("BİTTİ: %d fuar denendi, %d liste bulundu, %d firma eklendi",
              state["total"], state["found_lists"], state["companies_added"])


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                         datefmt="%H:%M:%S")
    main()
