"""
Genel Tüyap platformu katılımcı listesi scraper.
intermobistanbul.com ve madenturkiyefuari.com gibi aynı CMS'i kullanan
Tüyap fuar sitelerinin hepsinde katılımcı listesi sayfası aynı yapıda:
div.brand-item > h2.brand-name / p.brand-country / div.location-item span
(Salon/Stant) / a.brand-link (detay linki), ?page=N sayfalama, düz HTML.

Kullanım:
    python tuyap_platform_scrape.py --base-url https://X.com/katilimci-listesi \
        --pages 65 --name "Maden Türkiye 2026" --slug maden_turkiye_2026 --append
"""

import argparse
import json
import logging
import time
import unicodedata
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl

from scraper import make_company, deduplicate, HEADERS
import datastore as ds

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

OUT_DIR = Path(__file__).parent / "output"
OUT_DIR.mkdir(exist_ok=True)
COMPANIES_FILE = OUT_DIR / "companies.json"


def parse_page(html: str, base_domain: str, fair_name: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    companies = []
    for item in soup.select("div.brand-item"):
        name_el = item.select_one("h2.brand-name")
        if not name_el:
            continue
        name = name_el.get_text(strip=True)
        if not name:
            continue

        # Tüyap CMS'inin kendi çıktısında "Türkı̇ye" gibi combining-dot (U+0307)
        # bozuk karakterler geliyor (kaynak sitenin kendi Türkçe İ/lower() hatası
        # — bkz. enricher.py'deki aynı bug). Kazıma anında düzeltiyoruz ki her
        # seferinde elle companies.json normalize etmek gerekmesin.
        country_el = item.select_one("p.brand-country")
        country = country_el.get_text(strip=True) if country_el else ""
        if country:
            country = unicodedata.normalize("NFC", country.replace("ı̇", "i").replace("i̇", "i"))

        hall, stand = "", ""
        for loc in item.select("div.location-item span"):
            text = loc.get_text(strip=True)
            if text.lower().startswith("salon"):
                hall = text.split(":", 1)[-1].strip()
            elif text.lower().startswith("stant"):
                stand = text.split(":", 1)[-1].strip()

        link_el = item.select_one("a.brand-link")
        detail_path = link_el["href"] if link_el and link_el.has_attr("href") else ""

        company = make_company(name, country=country, fair=fair_name)
        company["hall"] = hall
        company["stand"] = stand
        company["detail_url"] = f"{base_domain}/{detail_path}" if detail_path else ""
        companies.append(company)
    return companies


def scrape(base_url: str, total_pages: int, fair_name: str) -> list[dict]:
    base_domain = base_url.split("/katilimci", 1)[0].rstrip("/")
    all_companies = []
    for page in range(1, total_pages + 1):
        url = base_url if page == 1 else f"{base_url}?page={page}"
        log.info("Sayfa %d/%d çekiliyor: %s", page, total_pages, url)
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            log.error("Sayfa %d başarısız: %s", page, e)
            continue

        page_companies = parse_page(resp.text, base_domain, fair_name)
        log.info("  %d firma bulundu", len(page_companies))
        all_companies.extend(page_companies)
        time.sleep(1.0)

    return deduplicate(all_companies)


def main():
    parser = argparse.ArgumentParser(description="Tüyap platformu katılımcı scraper (genel)")
    parser.add_argument("--base-url", required=True, help="örn: https://X.com/katilimci-listesi")
    parser.add_argument("--pages", type=int, required=True, help="toplam sayfa sayısı")
    parser.add_argument("--name", required=True, help="fuar adı (companies.json'daki 'fair' alanı)")
    parser.add_argument("--slug", required=True, help="çıktı dosya adı öneki (örn. maden_turkiye_2026)")
    parser.add_argument("--append", action="store_true", help="output/companies.json'a da ekle")
    args = parser.parse_args()

    companies = scrape(args.base_url, args.pages, args.name)
    if not companies:
        log.error("HATA: hiç firma bulunamadı, HTML yapısı değişmiş olabilir.")
        return

    log.info("Toplam %d firma bulundu.", len(companies))

    json_path = OUT_DIR / f"{args.slug}.json"
    json_path.write_text(json.dumps(companies, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = args.name[:31]
    ws.append(["Firma Adı", "Ülke", "Salon", "Stant", "Detay URL", "Öncelik"])
    for c in companies:
        ws.append([c["name"], c["country"], c["hall"], c["stand"], c["detail_url"], c["priority"]])
    xlsx_path = OUT_DIR / f"{args.slug}.xlsx"
    wb.save(xlsx_path)

    log.info("JSON : %s", json_path)
    log.info("XLSX : %s", xlsx_path)

    if args.append:
        # datastore.append_companies: webapp çalışırken bile güvenli — atomik
        # yazma + yedek rotasyonu, dedupe zaten var (name+fair anahtarı).
        added = ds.append_companies(companies)
        log.info("companies.json güncellendi → %d yeni firma eklendi", added)

    log.info("İlk 10 firma:")
    for c in companies[:10]:
        log.info("  [%s] %s (%s) — Salon %s / Stant %s", c["priority"], c["name"], c["country"], c["hall"], c["stand"])


if __name__ == "__main__":
    main()
