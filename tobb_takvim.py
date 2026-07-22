"""
TOBB Fuar Takvimi entegrasyonu — https://fuarlar.tobb.org.tr/FuarTakvimi

Sayfa Blazor Server ile render ediliyor ama tablo satırları ilk HTML
yanıtında zaten dolu geliyor (server-side prerender) — Playwright/JS
GEREKMİYOR, düz requests+BeautifulSoup yeterli.

Tablo kolon sırası (th başlıklarından doğrulandı):
  # | Başlangıç Tar. | Bitiş Tar. | Fuarın Adı | Konusu | Başlıca Ürün
  Grupları | Türü | Fuar Yeri | Şehir | Düzenleyici | Konu 1/2/3 (kod) |
  Web | E-Mail | (boş)
"""

import json
import logging
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

TOBB_URL = "https://fuarlar.tobb.org.tr/FuarTakvimi"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}

OUT_DIR = Path(__file__).parent / "output"
CACHE_FILE = OUT_DIR / "tobb_fuar_takvimi.json"


def _parse_date(s: str) -> str:
    """DD.MM.YYYY -> YYYY-MM-DD (sıralama/karşılaştırma için); parse edilemezse boş."""
    try:
        return datetime.strptime(s.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def fetch_calendar() -> list[dict]:
    resp = requests.get(TOBB_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    fairs = []
    for tr in soup.select("tbody tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 15:
            continue
        cell = [td.get_text(" ", strip=True) for td in tds]

        start_raw, end_raw = cell[1], cell[2]
        start_iso, end_iso = _parse_date(start_raw), _parse_date(end_raw)
        if not start_iso:
            continue  # başlık/boş satır vb.

        fairs.append({
            "start_date": start_raw,
            "end_date": end_raw,
            "start_date_iso": start_iso,
            "end_date_iso": end_iso,
            "name": cell[3],
            "konu": cell[4],
            "urun_gruplari": cell[5],
            "tur": cell[6],
            "yer": cell[7],
            "sehir": cell[8],
            "duzenleyici": cell[9],
            "web": cell[13] if len(cell) > 13 else "",
            "email": cell[14] if len(cell) > 14 else "",
        })

    fairs.sort(key=lambda f: f["start_date_iso"])
    return fairs


def scrape_and_save() -> dict:
    fairs = fetch_calendar()
    OUT_DIR.mkdir(exist_ok=True)
    payload = {
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count": len(fairs),
        "fairs": fairs,
    }
    CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("TOBB fuar takvimi kaydedildi: %d fuar", len(fairs))
    return payload


def load_cached() -> dict:
    if not CACHE_FILE.exists():
        return {"fetched_at": "", "count": 0, "fairs": []}
    return json.loads(CACHE_FILE.read_text(encoding="utf-8"))


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    result = scrape_and_save()
    log.info("Toplam %d fuar bulundu.", result["count"])
