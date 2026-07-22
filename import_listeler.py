"""
Obsidian vault'un Listeler/ klasörüne kullanıcının elle koyduğu hazır katılımcı
Excel'lerini companies.json'a aktarır.

Üç dosya (Intermob, KOMATEK, F İstanbul) zaten sistemde olan fuarlarla eşleşiyor
— bunlar için MERGE modu: isim eşleşirse eksik telefon/e-posta/website alanları
dolduruluyor (üzerine yazmıyor), eşleşmeyen yeni firmalar ekleniyor. Diğerleri
tamamen yeni fuar olarak ekleniyor.
"""

import json
import logging
import unicodedata
from pathlib import Path

import openpyxl

import datastore as ds
from scraper import make_company, _fold_tr, _norm

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

VAULT_DIR = Path(__file__).parent.parent.parent  # .../obsidian
LISTELER_DIR = VAULT_DIR / "Listeler"

HEADER_ALIASES = {
    "sirket adi": "name", "firma adi": "name", "firma": "name", "isim": "name",
    "company_name": "name",
    "telefon": "phone",
    "e-posta": "email", "eposta": "email", "mail": "email",
    "website": "website",
    "ulke": "country", "country": "country",
    "hall": "hall",
    "stand": "stand", "stand_no": "stand",
    "aciklama": "note",
}

# (dosya adı, fuar adı, mevcut fuarla birleştir mi)
IMPORT_JOBS = [
    ("Automekanika.xlsx", "Automechanika İstanbul 2026", False),
    ("aymod.xlsx", "Aymod İstanbul 2026", False),
    ("F istanbul tam liste.xlsx", "F İstanbul 2026", True),
    ("F İSTANBUL Katılımcı listesi.xlsx", "F İstanbul 2026", True),
    ("food expo.xlsx", "Food Expo İstanbul 2026", False),
    ("hightex2026_exhibitors_2026-03-26_131037_doldurulmus.xlsx", "Hightex 2026", False),
    ("Hometex_2025_Katilimcilar.xlsx", "Hometex 2025", False),
    ("intermob istanbul.xlsx", "Intermob 2026", True),
    ("izmir mermer fuar.xlsx", "İzmir Mermer Fuarı 2026", False),
    ("Kapı pencere fuarı.xlsx", "Avrasya Kapı Pencere 2026", False),
    ("komatek 2026-2.xlsx", "KOMATEK 2026", True),
    ("maktek 2026.xlsx", "Maktek Avrasya 2026", False),
]


def _clean_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s = s.lstrip("").strip()  # bazı dosyalarda adres hücrelerinin başında ikon glyph'i var
    return s


def _clean_country(v: str) -> str:
    if not v:
        return ""
    return unicodedata.normalize("NFC", v.replace("ı̇", "i").replace("i̇", "i"))


def load_rows(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter)
    col_map = {}  # col_index -> canonical field
    for i, h in enumerate(header):
        key = _norm(str(h or ""))
        if key in HEADER_ALIASES:
            col_map[i] = HEADER_ALIASES[key]

    out = []
    for row in rows_iter:
        rec = {}
        for i, field in col_map.items():
            if i < len(row):
                rec[field] = _clean_text(row[i])
        name = rec.get("name", "")
        if not name or _norm(name) == "sirket adi" or _norm(name) == "firma adi":
            continue  # boş satır veya başlık tekrarı (ör. komatek 2026-2.xlsx'teki hatalı satır)
        out.append(rec)
    wb.close()
    return out


# F İSTANBUL Katılımcı listesi.xlsx'teki AÇIKLAMA sütunu aslında gerçek arama
# geçmişi (ör. "ULAŞAMADIM", "ANLAŞMIŞLAR") — bu bilgiyi durum alanına da
# işliyoruz ki kaybolmasın, satışçı aynı firmayı sıfırdan aramasın.
NOTE_TO_STATUS = {
    "ulasamadim": "📞 Arandı",
    "anlasmislar": "✅ Anlaşıldı",
}


def _status_from_note(note: str) -> str:
    n = _norm(note)
    for key, status in NOTE_TO_STATUS.items():
        if key in n:
            return status
    if "mail atildi" in n:
        return "✉️ Mail Atıldı"
    return ""


def build_company(rec: dict, fair_name: str) -> dict:
    c = make_company(rec["name"], country=_clean_country(rec.get("country", "")), fair=fair_name)
    c["phone"] = rec.get("phone", "")
    c["email"] = rec.get("email", "")
    c["website"] = rec.get("website", "")
    c["note"] = rec.get("note", "")
    if rec.get("hall"):
        c["hall"] = rec["hall"]
    if rec.get("stand"):
        c["stand"] = rec["stand"]
    mapped_status = _status_from_note(rec.get("note", ""))
    if mapped_status:
        c["status"] = mapped_status
    return c


def main():
    data = ds.load()
    by_key = {(_norm(c.get("name", "")), c.get("fair", "")): c for c in data}

    summary = []
    for filename, fair_name, merge in IMPORT_JOBS:
        path = LISTELER_DIR / filename
        if not path.exists():
            summary.append((filename, "DOSYA BULUNAMADI", 0, 0, 0))
            continue

        rows = load_rows(path)
        added, backfilled, skipped_dupe = 0, 0, 0

        if merge:
            for rec in rows:
                key = (_norm(rec["name"]), fair_name)
                existing = by_key.get(key)
                if existing:
                    changed = False
                    for field in ("phone", "email", "website"):
                        new_val = rec.get(field, "")
                        if new_val and not existing.get(field):
                            existing[field] = new_val
                            changed = True
                    note_val = rec.get("note", "")
                    if note_val and not existing.get("note"):
                        existing["note"] = note_val
                        changed = True
                    mapped_status = _status_from_note(note_val)
                    if mapped_status and existing.get("status", "").startswith("⬜"):
                        existing["status"] = mapped_status
                        changed = True
                    if changed:
                        backfilled += 1
                    else:
                        skipped_dupe += 1
                else:
                    new_c = build_company(rec, fair_name)
                    new_c["id"] = ds.make_id(new_c)
                    data.append(new_c)
                    by_key[key] = new_c
                    added += 1
            ds.save(data)
        else:
            new_companies = [build_company(rec, fair_name) for rec in rows]
            added = ds.append_companies(new_companies)
            skipped_dupe = len(rows) - added
            data = ds.load()  # append_companies kendi save'ini yaptı, tazele
            by_key = {(_norm(c.get("name", "")), c.get("fair", "")): c for c in data}

        summary.append((filename, fair_name, added, backfilled, skipped_dupe))
        log.info("%s -> %s | +%d yeni, %d alan tamamlandı, %d zaten vardı",
                  filename, fair_name, added, backfilled, skipped_dupe)

    print("\n=== ÖZET ===")
    for filename, fair_name, added, backfilled, skipped_dupe in summary:
        print(f"{filename:55s} -> {fair_name:35s} +{added:4d} yeni  {backfilled:4d} tamamlandı  {skipped_dupe:4d} zaten var")


if __name__ == "__main__":
    main()
