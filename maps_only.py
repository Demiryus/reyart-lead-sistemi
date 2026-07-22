"""
Sadece Google Maps fallback'ı çalıştır.
Mevcut output/companies.json'daki telefonu eksik firmalar için
MapsScraperProv1 worker'ını çağırır, sonucu leads.xlsx'e işler.
"""

import json
from pathlib import Path

from enricher import run_maps_fallback, save_excel, COMPANIES_FILE, log


def main():
    if not COMPANIES_FILE.exists():
        log.error("companies.json bulunamadı.")
        return

    companies = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    log.info("Toplam %d firma yüklendi", len(companies))

    eksik_once = sum(1 for c in companies if not c.get("phone"))
    log.info("Telefonu eksik: %d firma", eksik_once)

    companies = run_maps_fallback(companies)

    eksik_sonra = sum(1 for c in companies if not c.get("phone"))
    log.info("Maps sonrası eksik: %d (önce: %d)", eksik_sonra, eksik_once)

    COMPANIES_FILE.write_text(
        json.dumps(companies, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    save_excel(companies)
    log.info("Tamamlandı.")


if __name__ == "__main__":
    main()
