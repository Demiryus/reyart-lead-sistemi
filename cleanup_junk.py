"""
companies.json'daki junk kayıtları temizler (yedek alarak).

Toplu takvim taramasında (bulk_scrape_calendar.py) bazı organizatör siteleri
katılımcı listesi yerine kendi fuar menüsünü/pazarlama metnini döndürdü:
- Üniversite tercih fuarları: "firma" diye organizatörün diğer fuarlarının
  adları kaydedilmiş → bu bucket'lar KOMPLE silinir.
- Nav/pazarlama metinleri ("Upcoming Events", "Why Should You Choose..."),
  istatistik satırları ("27.6%", "293.1"), alıntılar → kayıt bazında silinir.

Çalıştırma: python cleanup_junk.py           (önce ne sileceğini raporlar)
            python cleanup_junk.py --apply   (gerçekten siler)
"""

import argparse
import json
import re
from pathlib import Path

import datastore as ds
from scraper import _norm

OUT_DIR = Path(__file__).parent / "output"
TOBB_FILE = OUT_DIR / "tobb_fuar_takvimi.json"

# Bucket'ı komple sil: fuar adında bu geçenler (organizatör sitesi katılımcı
# listesi değil, kendi fuar takvimini/marketing sayfasını vermiş)
JUNK_FAIR_SUBSTRINGS = ["tercih", "umre ve turizm"]

NAV_KEYWORDS = {
    "about", "home", "events", "contact", "exhibitions", "corporate",
    "upcoming", "mission", "media", "sponsors", "participants", "vision",
    "english", "deutsch", "turkce", "gallery", "downloads", "register",
}

PURE_NUMBER_RE = re.compile(r"^[\d.,]+\s*%?$")


def is_junk_record(c: dict, tobb_names: set[str]) -> str:
    """Junk ise sebep string'i, değilse '' döndürür."""
    name = c.get("name", "").strip()
    n = _norm(name)

    if PURE_NUMBER_RE.match(name):
        return "sayı/istatistik"
    if name.startswith(('"', "“", "'", "‘")):
        return "alıntı/testimonial"
    if "tercih fuari" in n or "tercih gunleri" in n:
        return "fuar adı (tercih)"
    if n in tobb_names:
        return "TOBB fuar adıyla birebir aynı"
    words = set(re.findall(r"[a-z]+", n))
    nav_hits = words & NAV_KEYWORDS
    if len(nav_hits) >= 2:
        return f"nav menüsü ({','.join(sorted(nav_hits))})"
    return ""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="gerçekten sil (yoksa sadece rapor)")
    args = parser.parse_args()

    tobb_names = set()
    if TOBB_FILE.exists():
        cal = json.loads(TOBB_FILE.read_text(encoding="utf-8"))
        tobb_names = {_norm(f.get("name", "")) for f in cal.get("fairs", [])}

    data = ds.load()
    keep, removed = [], []
    for c in data:
        fair_n = _norm(c.get("fair", ""))
        if any(sub in fair_n for sub in JUNK_FAIR_SUBSTRINGS):
            removed.append((c["name"], c["fair"], "junk fuar bucket'ı"))
            continue
        reason = is_junk_record(c, tobb_names)
        if reason:
            removed.append((c["name"], c["fair"], reason))
            continue
        keep.append(c)

    report = OUT_DIR / "_cleanup_report.txt"
    with open(report, "w", encoding="utf-8") as f:
        f.write(f"Toplam: {len(data)} | Silinecek: {len(removed)} | Kalacak: {len(keep)}\n\n")
        for name, fair, reason in removed:
            f.write(f"[{reason}] {name[:70]}  <<{fair[:40]}>>\n")
    print(f"Rapor: {report}")
    print(f"Toplam {len(data)} -> silinecek {len(removed)}, kalacak {len(keep)}")

    if args.apply:
        ds.save(keep)  # otomatik yedek alır
        print("UYGULANDI (yedek otomatik alındı).")
    else:
        print("Kuru çalıştırma — silmek için: python cleanup_junk.py --apply")


if __name__ == "__main__":
    main()
