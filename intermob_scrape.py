"""
Intermob 2026 katılımcı listesi scraper.
Kaynak: https://intermobistanbul.com/katilimci-listesi (düz HTML, ?page=N sayfalama)
Çıktı: output/intermob_2026.json + output/intermob_2026.xlsx
--append ile output/companies.json'a da eklenir (fair="Intermob 2026").
"""

import json
import logging
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl

from scraper import make_company, deduplicate, HEADERS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

BASE_URL = "https://intermobistanbul.com/katilimci-listesi"
TOTAL_PAGES = 15
OUT_DIR = Path(__file__).parent / "output"
OUT_DIR.mkdir(exist_ok=True)
COMPANIES_FILE = OUT_DIR / "companies.json"


def parse_page(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    companies = []
    for item in soup.select("div.brand-item"):
        name_el = item.select_one("h2.brand-name")
        if not name_el:
            continue
        name = name_el.get_text(strip=True)
        if not name:
            continue

        country_el = item.select_one("p.brand-country")
        country = country_el.get_text(strip=True) if country_el else ""

        hall, stand = "", ""
        for loc in item.select("div.location-item span"):
            text = loc.get_text(strip=True)
            if text.lower().startswith("salon"):
                hall = text.split(":", 1)[-1].strip()
            elif text.lower().startswith("stant"):
                stand = text.split(":", 1)[-1].strip()

        link_el = item.select_one("a.brand-link")
        detail_path = link_el["href"] if link_el and link_el.has_attr("href") else ""

        company = make_company(name, country=country, fair="Intermob 2026")
        company["hall"] = hall
        company["stand"] = stand
        company["detail_url"] = (
            f"https://intermobistanbul.com/{detail_path}" if detail_path else ""
        )
        companies.append(company)
    return companies


def scrape() -> list[dict]:
    all_companies = []
    for page in range(1, TOTAL_PAGES + 1):
        url = BASE_URL if page == 1 else f"{BASE_URL}?page={page}"
        log.info("Sayfa %d/%d çekiliyor: %s", page, TOTAL_PAGES, url)
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            log.error("Sayfa %d başarısız: %s", page, e)
            continue

        page_companies = parse_page(resp.text)
        log.info("  %d firma bulundu", len(page_companies))
        all_companies.extend(page_companies)
        time.sleep(1.0)

    return deduplicate(all_companies)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Intermob 2026 katılımcı scraper")
    parser.add_argument("--append", action="store_true",
                        help="output/companies.json'a da ekle")
    args = parser.parse_args()

    companies = scrape()
    if not companies:
        log.error("HATA: hiç firma bulunamadı, HTML yapısı değişmiş olabilir.")
        return

    log.info("Toplam %d firma bulundu.", len(companies))

    json_path = OUT_DIR / "intermob_2026.json"
    json_path.write_text(json.dumps(companies, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Intermob 2026"
    ws.append(["Firma Adı", "Ülke", "Salon", "Stant", "Detay URL", "Öncelik"])
    for c in companies:
        ws.append([c["name"], c["country"], c["hall"], c["stand"], c["detail_url"], c["priority"]])
    xlsx_path = OUT_DIR / "intermob_2026.xlsx"
    wb.save(xlsx_path)

    log.info("JSON : %s", json_path)
    log.info("XLSX : %s", xlsx_path)

    if args.append:
        existing = json.loads(COMPANIES_FILE.read_text(encoding="utf-8")) if COMPANIES_FILE.exists() else []
        merged = deduplicate(existing + companies)
        COMPANIES_FILE.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
        log.info("companies.json güncellendi → toplam %d firma", len(merged))

    log.info("İlk 10 firma:")
    for c in companies[:10]:
        log.info("  [%s] %s (%s) — Salon %s / Stant %s", c["priority"], c["name"], c["country"], c["hall"], c["stand"])


if __name__ == "__main__":
    main()
