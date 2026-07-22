"""
Günlük 50 kişiselleştirilmiş tanıtım maili — Windows Görev Zamanlayıcı ile
her gün otomatik koşar (Claude'a gerek yok, token harcamaz).

Seçim: e-postası olan + fuarı GELECEKTE olan + bu şablonu daha önce ALMAMIŞ
firmalar, fuar tarihi en yakın olandan başlayarak 50 tanesi. Yerliye
tanitim_tr, yabancıya tanitim_en gider. Katalog PDF'leri otomatik eklenir
(mailer.CATALOG_ATTACHMENTS). Çift gönderim mail_log.jsonl ile engelli.

Sonra Masaüstü/Listeler/Mail-Raporu.xlsx yeniden üretilir (tüm gönderim
geçmişi) ve ekranda özet bildirimi gösterilir.

Kullanım:
    python send_daily_50.py            # gerçek gönderim
    python send_daily_50.py --dry-run  # sadece kimlere gideceğini göster
Görev Zamanlayıcı kaydı: send_daily_50.bat (aynı klasör) her gün 09:30.
"""

import argparse
import ctypes
import json
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import datastore as ds
import mailer

DAILY_TARGET = 50
# Mailler arası 60-180 sn rastgele bekleme: 50 mail ~09:30'dan öğlene yayılır,
# patlama deseni yerine doğal tempo (spam filtresi riski düşer).
DELAY_RANGE = (60, 180)
REPORT_XLSX = Path.home() / "OneDrive" / "Desktop" / "Listeler" / "Mail-Raporu.xlsx"


def pick_candidates(n: int) -> list[dict]:
    already = mailer._load_sent_keys()
    today = date.today().isoformat()
    cands = []
    for c in ds.load():
        email = (c.get("email") or "").strip()
        if not email or not mailer.EMAIL_RE.match(email):
            continue
        if not c.get("fuar_tarihi") or c["fuar_tarihi"] < today:
            continue
        template = "tanitim_en" if "Yabancı" in str(c.get("mensei", "")) else "tanitim_tr"
        if c.get("id", "") + "|" + template in already:
            continue
        c["_template"] = template
        cands.append(c)
    cands.sort(key=lambda c: c["fuar_tarihi"])
    return cands[:n]


def rebuild_report():
    from openpyxl import Workbook
    data = {c["id"]: c for c in ds.load()}
    wb = Workbook()
    ws = wb.active
    ws.title = "Gönderilen Mailler"
    ws.append(["Gönderim Zamanı", "Firma", "E-posta", "Fuar", "Fuar Tarihi", "Şablon", "Telefon"])
    if mailer.LOG_FILE.exists():
        with open(mailer.LOG_FILE, encoding="utf-8") as f:
            for line in f:
                try:
                    o = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if o.get("ok"):
                    c = data.get(o.get("company_id", ""), {})
                    ws.append([o.get("ts", ""), o.get("name", ""), o.get("to", ""),
                               c.get("fair", ""), c.get("fuar_tarihi", ""),
                               o.get("template", ""), c.get("phone", "")])
    for col, w in zip("ABCDEFG", [19, 38, 34, 22, 12, 12, 18]):
        ws.column_dimensions[col].width = w
    REPORT_XLSX.parent.mkdir(exist_ok=True)
    wb.save(REPORT_XLSX)


def notify(title: str, text: str):
    # 0x40=bilgi ikonu, 0x1000=en üstte (system modal) — görev zamanlayıcıdan
    # koşarken de görünsün diye
    ctypes.windll.user32.MessageBoxW(0, text, title, 0x40 | 0x1000)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-popup", action="store_true")
    args = parser.parse_args()

    picked = pick_candidates(DAILY_TARGET)
    if not picked:
        if not args.no_popup:
            notify("Reyart Günlük Mail", "Gönderilecek yeni aday kalmadı (e-postalı ve fuarı yaklaşan tüm firmalara gönderildi).")
        print("aday yok")
        return

    groups = {}
    for c in picked:
        groups.setdefault(c["_template"], []).append(c)

    sent = skipped = errors = 0
    lines = []
    try:
        for template, comps in groups.items():
            r = mailer.send_bulk(comps, template, dry_run=args.dry_run,
                                 delay_range=None if args.dry_run else DELAY_RANGE)
            sent += len(r["sent"])
            skipped += len(r["skipped"])
            errors += len(r["errors"])
            for e in r["errors"][:3]:
                lines.append(f"HATA {e['name']}: {e['error'][:60]}")
    except RuntimeError as e:  # Outlook kapalı/hesap yok
        if not args.no_popup:
            notify("Reyart Günlük Mail — GÖNDERİLEMEDİ", str(e))
        print("HATA:", e)
        sys.exit(1)

    if args.dry_run:
        print(f"[dry-run] gönderilecek: {sent}, atlanacak: {skipped}")
        for c in picked[:5]:
            print(" ", c["fuar_tarihi"], c["fair"][:30], "|", c["name"][:35], "->", c["email"], f"[{c['_template']}]")
        return

    rebuild_report()
    fairs = sorted({c["fair"] for c in picked})
    msg = (f"Gönderilen: {sent}  |  Atlanan: {skipped}  |  Hata: {errors}\n"
           f"Fuarlar: {', '.join(fairs)[:200]}\n"
           f"Rapor: Masaüstü\\Listeler\\Mail-Raporu.xlsx")
    if lines:
        msg += "\n" + "\n".join(lines)
    print(msg)
    if not args.no_popup:
        notify("Reyart Günlük Mail — Tamamlandı ✅", msg)


if __name__ == "__main__":
    main()
