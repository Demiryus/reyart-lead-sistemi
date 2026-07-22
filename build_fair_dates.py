"""
companies.json'daki her fuar adına başlangıç tarihi eşler → output/fair_dates.json

Kaynaklar:
1. TOBB takvim cache'i (output/tobb_fuar_takvimi.json) — toplu taramayla eklenen
   fuarların adı TOBB'daki adla birebir aynı, direkt eşleşir. Aynı ada birden
   fazla kayıt varsa (ör. Aymod Mart + Ekim) GELECEK olan tercih edilir.
2. MANUAL_DATES — elle eklenen/Listeler'den import edilen fuarlar (TOBB adıyla
   birebir eşleşmeyenler).

Satış skoru bu tarihlerden hesaplanır (datastore.py) — fuarı yaklaşan firmanın
stand ihtiyacı ŞİMDİ olduğu için satış önceliği fuar tarihine bağlı.

Çalıştırma: python build_fair_dates.py   (TOBB takvimi her güncellendiğinde tekrar)
"""

import json
import logging
from datetime import date
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

OUT_DIR = Path(__file__).parent / "output"
COMPANIES_FILE = OUT_DIR / "companies.json"
TOBB_FILE = OUT_DIR / "tobb_fuar_takvimi.json"
FAIR_DATES_FILE = OUT_DIR / "fair_dates.json"

# Elle eklenen fuarlar (companies.json'daki ad → başlangıç tarihi ISO).
# TOBB'daki resmi adları farklı olduğu için otomatik eşleşmiyorlar.
MANUAL_DATES = {
    "KOMATEK 2026": "2026-06-03",            # geçti
    "Maden Türkiye 2026": "2026-04-08",      # geçti
    "Intermob 2026": "2026-09-17",
    "F İstanbul 2026": "2026-08-26",
    "Maktek Avrasya 2026": "2026-09-28",
    "Aymod İstanbul 2026": "2026-10-07",     # TOBB'da Mart + Ekim var; elimizdeki liste Ekim edisyonu için
    "Hometex 2025": "2025-05-20",            # geçti (2025 listesi — 2026 edisyonu 2026-05-19 ama liste eski)
    "Hightex 2026": "2026-03-26",            # geçti (dosya adından)
    "Avrasya Kapı Pencere 2026": "2026-11-21",
    "Automechanika İstanbul 2026": "",       # tarih doğrulanamadı — bilinmiyor bırak
    "İzmir Mermer Fuarı 2026": "2026-04-01", # tahmini geçmiş (Mermer İzmir genelde Mart/Nisan); satışta soğuk say
    "Food Expo İstanbul 2026": "",           # tarih doğrulanamadı
}


def main():
    companies = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    fair_names = sorted({c.get("fair", "") for c in companies if c.get("fair")})

    tobb = {}
    if TOBB_FILE.exists():
        cal = json.loads(TOBB_FILE.read_text(encoding="utf-8"))
        today = date.today().isoformat()
        for f in cal.get("fairs", []):
            name = f.get("name", "")
            iso = f.get("start_date_iso", "")
            if not name or not iso:
                continue
            prev = tobb.get(name)
            if prev is None:
                tobb[name] = iso
            else:
                # Aynı ad birden fazla: gelecekteki en yakını tercih et
                cands = [d for d in (prev, iso) if d >= today]
                tobb[name] = min(cands) if cands else max(prev, iso)

    result, matched, manual, unknown = {}, 0, 0, 0
    for name in fair_names:
        if name in MANUAL_DATES:
            result[name] = MANUAL_DATES[name]
            manual += 1
        elif name in tobb:
            result[name] = tobb[name]
            matched += 1
        else:
            result[name] = ""
            unknown += 1
            log.warning("Tarih bulunamadı: %r", name)

    FAIR_DATES_FILE.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("fair_dates.json yazıldı: %d fuar (%d TOBB, %d manuel, %d bilinmiyor)",
             len(result), matched, manual, unknown)


if __name__ == "__main__":
    main()
