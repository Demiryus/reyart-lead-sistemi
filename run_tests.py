"""
Reyart Lead Manager — hızlı regresyon testi.

Kapsam: datastore birimleri, enricher.save_progress merge kuralları,
scraper junk filtresi, webapp API uçları (canlı sunucu gerekir),
Excel export. Windows cp1252 konsolunda çökmemesi için rapor dosyaya yazılır.

Çalıştırma: python run_tests.py   → output/_test_report.txt
Çıkış kodu: 0 = hepsi geçti, 1 = en az bir hata.
"""

import json
import sys
import traceback
from pathlib import Path

OUT = Path(__file__).parent / "output"
REPORT = OUT / "_test_report.txt"

results = []


def check(name, fn):
    try:
        fn()
        results.append(("PASS", name, ""))
    except Exception as e:
        results.append(("FAIL", name, f"{e.__class__.__name__}: {e}\n{traceback.format_exc(limit=3)}"))


# ── 1. scraper junk filtresi ─────────────────────────────────────────────────
def t_junk_filter():
    from scraper import is_valid_company
    junk = ["2024", "27.6%", "İstanbul Üniversite Tercih Fuarı",
            "Upcoming Events Exhibitions", '"Bir alıntı metni"',
            "Home About Contact Media", "Türkiye", "Germany", "cookie_x"]
    real = ["SKF", "NTN", "Anthon GmbH", "HYUNDAI L&C", "Media Markt",
            "3K PLASTİK VE KALIP SAN. VE TİC. LTD. ŞTİ.",
            "Tüyap Tüm Fuarcılık Yapım A.Ş"]
    for s in junk:
        assert not is_valid_company(s), f"junk geçti: {s!r}"
    for s in real:
        assert is_valid_company(s), f"gerçek firma elendi: {s!r}"


# ── 2. datastore birimleri ───────────────────────────────────────────────────
def t_datastore_load():
    import datastore as ds
    data = ds.load()
    assert len(data) > 9000, f"beklenmedik kayıt sayısı: {len(data)}"
    c = data[0]
    for field in ("id", "guven_seviyesi", "mensei", "sicaklik", "satis_skoru", "fuar_tarihi"):
        assert field in c, f"hesaplanan alan yok: {field}"
    # önbellek: ikinci çağrı aynı objeyi döndürmeli (mtime değişmedi)
    assert ds.load() is data, "load() önbelleği çalışmıyor"


def t_datastore_sicaklik():
    import datastore as ds
    assert ds.compute_sicaklik("", "2026-07-09")[0].startswith("❄")
    assert ds.compute_sicaklik("2026-06-01", "2026-07-09")[0].startswith("❄")  # geçmiş
    assert ds.compute_sicaklik("2026-08-26", "2026-07-09")[0].startswith("🔥")  # 48 gün
    assert ds.compute_sicaklik("2026-12-31", "2026-07-09")[0].startswith("🌡")  # 175 gün


def t_datastore_update_roundtrip():
    import datastore as ds
    data = ds.load()
    cid = data[0]["id"]
    orig_note = data[0].get("note", "")
    upd = ds.update_company(cid, {"note": "TEST_NOTU_XYZ"})
    assert upd["note"] == "TEST_NOTU_XYZ"
    assert ds.update_company(cid, {"note": orig_note})["note"] == orig_note
    try:
        ds.update_company(cid, {"salakalan": "x"})
        raise AssertionError("geçersiz alan kabul edildi")
    except ValueError:
        pass


# ── 3. enricher.save_progress merge kuralları ────────────────────────────────
def t_save_progress_merge():
    import datastore as ds
    from enricher import save_progress
    data = ds.load()
    n0 = len(data)
    sample = dict(data[0])
    sid = sample["id"]
    orig_linkedin = data[0].get("linkedin", "")
    orig_phone = data[0].get("phone", "")

    ghost = {"id": "deadbeef0000", "name": "HAYALET", "fair": "X", "phone": "123"}
    sample["linkedin"] = "https://linkedin.com/company/merge-test"
    sample["phone"] = "+900000000000"

    after = save_progress([ghost, sample])
    assert len(after) == n0, "kayıt sayısı değişti (hayalet geri geldi?)"
    t = next(c for c in after if c["id"] == sid)
    if orig_phone:
        assert t["phone"] == orig_phone, "dolu telefon ezildi"
    if not orig_linkedin:
        assert t["linkedin"] == "https://linkedin.com/company/merge-test", "boş alan dolmadı"
        # geri al
        fresh = ds.load()
        next(c for c in fresh if c["id"] == sid)["linkedin"] = ""
        ds.save(fresh, backup=False)


# ── 4. webapp API (canlı sunucu) ────────────────────────────────────────────
BASE = "http://127.0.0.1:5000"


def t_api_stats():
    import requests
    st = requests.get(f"{BASE}/api/stats", timeout=15).json()
    for k in ("total", "sicak", "yabanci", "yerli"):
        assert k in st["total"], f"stats.total.{k} yok"
    assert st["total"]["yerli"] + st["total"]["yabanci"] == st["total"]["total"]


def t_api_filters():
    import requests
    hot = requests.get(f"{BASE}/api/companies", params={"sicaklik": "🔥", "per_page": 10}, timeout=30).json()
    assert hot["total"] > 0
    assert all("🔥" in c["sicaklik"] for c in hot["items"])
    skor = requests.get(f"{BASE}/api/companies", params={"sort": "skor", "per_page": 10}, timeout=30).json()
    scores = [c["satis_skoru"] for c in skor["items"]]
    assert scores == sorted(scores, reverse=True), "skor sıralaması bozuk"
    yab = requests.get(f"{BASE}/api/companies", params={"mensei": "yabanci", "per_page": 5}, timeout=30).json()
    assert all("🌍" in c["mensei"] for c in yab["items"])
    # kombine filtre
    combo = requests.get(f"{BASE}/api/companies",
                         params={"sicaklik": "🔥", "mensei": "yabanci", "missing": "email"}, timeout=30).json()
    for c in combo["items"]:
        assert "🔥" in c["sicaklik"] and "🌍" in c["mensei"] and not c.get("email")


def t_api_meta_calendar():
    import requests
    meta = requests.get(f"{BASE}/api/meta", timeout=15).json()
    for k in ("fairs", "statuses", "sicaklik_levels"):
        assert k in meta
    cal = requests.get(f"{BASE}/api/calendar", timeout=30).json()
    assert cal["count"] > 400


def t_api_update_invalid():
    import requests
    r = requests.post(f"{BASE}/api/companies/yok_boyle_id", json={"note": "x"}, timeout=15)
    assert r.status_code == 404
    r2 = requests.post(f"{BASE}/api/companies/yok_boyle_id", json={}, timeout=15)
    assert r2.status_code == 400


def t_api_mail_preview():
    import requests
    data = requests.get(f"{BASE}/api/companies", params={"missing": "", "per_page": 3}, timeout=30).json()
    ids = [c["id"] for c in data["items"]]
    r = requests.post(f"{BASE}/api/mail/preview", json={"ids": ids, "template": "tanitim_tr"}, timeout=15)
    assert r.status_code == 200
    j = r.json()
    assert j["selected"] == len(ids)


# ── 5. Excel export ──────────────────────────────────────────────────────────
def t_excel_export():
    import datastore as ds
    from enricher import save_excel
    test_path = OUT / "_test_export.xlsx"
    save_excel(ds.load(), path=test_path)
    import openpyxl
    wb = openpyxl.load_workbook(test_path, read_only=True)
    names = wb.sheetnames
    assert "Özet" in names
    assert len(names) < 30, f"çok fazla sheet: {len(names)}"
    assert "Diğer Fuarlar" in names
    wb.close()
    test_path.unlink()


ALL = [
    ("scraper junk filtresi", t_junk_filter),
    ("datastore load + önbellek", t_datastore_load),
    ("datastore sıcaklık hesabı", t_datastore_sicaklik),
    ("datastore update roundtrip", t_datastore_update_roundtrip),
    ("enricher save_progress merge", t_save_progress_merge),
    ("API /stats", t_api_stats),
    ("API filtreler + skor sıralama", t_api_filters),
    ("API meta + takvim", t_api_meta_calendar),
    ("API geçersiz update", t_api_update_invalid),
    ("API mail preview", t_api_mail_preview),
    ("Excel export (sheet sınırı)", t_excel_export),
]

if __name__ == "__main__":
    for name, fn in ALL:
        check(name, fn)
    passed = sum(1 for r in results if r[0] == "PASS")
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write(f"SONUÇ: {passed}/{len(results)} geçti\n\n")
        for status, name, err in results:
            f.write(f"[{status}] {name}\n")
            if err:
                f.write(err + "\n")
    print(f"{passed}/{len(results)} passed -> {REPORT}")
    sys.exit(0 if passed == len(results) else 1)
