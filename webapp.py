"""
Reyart Lead Manager — web arayüz (localhost).

Çalıştırma:
    python webapp.py
    → http://127.0.0.1:5000

API:
    GET  /api/companies   ?fair=&search=&priority=&guven=&missing=&page=&per_page=
    GET  /api/stats
    GET  /api/meta        (fuar listesi, durum seçenekleri)
    POST /api/companies/<id>    body: {status?, note?, phone?, email?, website?, linkedin?, priority?}
    GET  /api/export      → leads.xlsx indir
"""

import subprocess
import sys
import threading
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file

import datastore as ds
import mailer
import tobb_takvim

app = Flask(__name__)
app.json.ensure_ascii = False


def _fold_tr(s: str) -> str:
    table = str.maketrans({
        "İ": "i", "I": "i", "ı": "i", "Ş": "s", "ş": "s", "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u", "Ö": "o", "ö": "o", "Ç": "c", "ç": "c",
    })
    return s.translate(table)


def _norm(s: str) -> str:
    return _fold_tr(s or "").lower()


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/companies")
def api_companies():
    data = ds.load()

    fair = request.args.get("fair", "").strip()
    search = _norm(request.args.get("search", "").strip())
    priority = request.args.get("priority", "").strip()
    guven = request.args.get("guven", "").strip()
    missing = request.args.get("missing", "").strip()  # phone|email|website|linkedin
    status = request.args.get("status", "").strip()
    mensei = request.args.get("mensei", "").strip()  # yerli|yabanci
    sicaklik = request.args.get("sicaklik", "").strip()  # 🔥|🌡|❄
    takip = request.args.get("takip", "").strip()  # "bugun" → takip tarihi gelmiş olanlar
    sort = request.args.get("sort", "").strip()  # "skor" → satış skoruna göre azalan

    from datetime import date as _date
    today_iso = _date.today().isoformat()

    def keep(c):
        if fair and c.get("fair") != fair:
            return False
        if priority and c.get("priority") != priority:
            return False
        if guven and guven not in c.get("guven_seviyesi", ""):
            return False
        if status and c.get("status") != status:
            return False
        if mensei == "yerli" and ds.is_foreign(c):
            return False
        if mensei == "yabanci" and not ds.is_foreign(c):
            return False
        if sicaklik and sicaklik not in c.get("sicaklik", ""):
            return False
        if takip == "bugun":
            t = c.get("takip_tarihi", "")
            if not t or t > today_iso:
                return False
        if missing and c.get(missing):
            return False
        if search:
            hay = _norm(" ".join(str(c.get(k, "")) for k in
                                 ("name", "country", "email", "website", "phone", "note")))
            if search not in hay:
                return False
        return True

    filtered = [c for c in data if keep(c)]

    if sort == "skor":
        filtered.sort(key=lambda c: (-c.get("satis_skoru", 0), c.get("fuar_tarihi") or "9999",
                                     c.get("name", "")))

    try:
        page = max(1, int(request.args.get("page", 1)))
        per_page = min(200, max(10, int(request.args.get("per_page", 50))))
    except ValueError:
        page, per_page = 1, 50

    start = (page - 1) * per_page
    return jsonify({
        "total": len(filtered),
        "page": page,
        "per_page": per_page,
        "items": filtered[start:start + per_page],
    })


@app.get("/api/stats")
def api_stats():
    return jsonify(ds.stats())


@app.get("/api/version")
def api_version():
    # Hafif poll endpoint'i: istemci 8 sn'de bir çağırır, damga değiştiyse
    # tabloyu yeniler → bir satışçının işaretlediğini diğerleri saniyeler
    # içinde görür.
    return jsonify({"version": ds.data_version()})


@app.get("/api/meta")
def api_meta():
    data = ds.load()
    fairs = sorted({c.get("fair", "") for c in data if c.get("fair")})
    priorities = ["⭐⭐⭐", "⭐⭐", "⭐"]
    return jsonify({
        "fairs": fairs,
        "priorities": priorities,
        "statuses": ds.STATUS_CHOICES,
        "guven_levels": ["🟢", "🟡", "⚪"],
        "sicaklik_levels": ["🔥", "🌡", "❄"],
    })


@app.post("/api/companies/<company_id>")
def api_update(company_id):
    body = request.get_json(silent=True) or {}
    if not body:
        return jsonify({"error": "Boş istek gövdesi"}), 400
    try:
        updated = ds.update_company(company_id, body)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    if updated is None:
        return jsonify({"error": "Firma bulunamadı"}), 404
    return jsonify(updated)


# ── Mail sistemi ──────────────────────────────────────────────────────────────

@app.get("/api/mail/templates")
def api_mail_templates():
    t = mailer.load_templates()
    return jsonify([
        {"id": k, "name": v["name"], "language": v["language"], "subject": v["subject"]}
        for k, v in t.items()
    ])


@app.get("/api/mail/status")
def api_mail_status():
    st = mailer.config_status()
    st["sent_today"] = mailer.sent_today_count()
    return jsonify(st)


def _companies_by_ids(ids: list[str]) -> list[dict]:
    data = ds.load()
    wanted = set(ids)
    return [c for c in data if c.get("id") in wanted]


@app.post("/api/mail/preview")
def api_mail_preview():
    body = request.get_json(silent=True) or {}
    ids, template_id = body.get("ids", []), body.get("template", "")
    templates = mailer.load_templates()
    if template_id not in templates:
        return jsonify({"error": "Şablon bulunamadı"}), 400
    companies = _companies_by_ids(ids)
    eligible = [c for c in companies if c.get("email")]
    previews = [
        {"name": c["name"], **mailer.render(templates[template_id], c)}
        for c in eligible[:3]
    ]
    return jsonify({
        "selected": len(companies),
        "eligible": len(eligible),
        "no_email": len(companies) - len(eligible),
        "previews": previews,
    })


@app.post("/api/mail/send")
def api_mail_send():
    body = request.get_json(silent=True) or {}
    ids = body.get("ids", [])
    template_id = body.get("template", "")
    dry_run = bool(body.get("dry_run", True))
    if not ids:
        return jsonify({"error": "Firma seçilmedi"}), 400
    companies = _companies_by_ids(ids)
    try:
        results = mailer.send_bulk(companies, template_id, dry_run=dry_run)
    except (ValueError, RuntimeError) as e:
        return jsonify({"error": str(e)}), 400

    # Gerçek gönderimde durumu güncelle (Aranmadı → Mail Atıldı)
    if not dry_run and results["sent"]:
        sent_names = {s["name"] for s in results["sent"]}
        for c in companies:
            if c["name"] in sent_names and c.get("status", "").startswith("⬜"):
                ds.update_company(c["id"], {"status": "✉️ Mail Atıldı"})
    return jsonify(results)


@app.post("/api/mail/compose")
def api_mail_compose():
    """Tek firma için classic Outlook'ta hazır taslak açar (göndermez).
    mailto: yerine bunu kullan — Windows'un varsayılan mail uygulamasına
    (yanlış hesaplı yeni Outlook olabiliyor) değil, doğrudan COM ile classic
    Outlook'a, doğru hesapla gider. Bu, webapp.py'nin çalıştığı PC'de açılır —
    LAN'daki başka bir bilgisayardan tıklanırsa pencere SUNUCU ekranında açılır."""
    body = request.get_json(silent=True) or {}
    ids = body.get("ids", [])
    template_id = body.get("template", "tanitim_tr")
    if not ids:
        return jsonify({"error": "Firma seçilmedi"}), 400
    companies = _companies_by_ids(ids)
    if not companies:
        return jsonify({"error": "Firma bulunamadı"}), 404
    try:
        result = mailer.compose_draft(companies[0], template_id)
    except (ValueError, RuntimeError) as e:
        return jsonify({"error": str(e)}), 400
    return jsonify(result)


# ── Fuar ekleme (URL'den kazı + otomatik zenginleştir) ───────────────────────

SCRAPE_JOBS: dict[str, dict] = {}


def _run_scrape_job(job_id: str, url: str, fair_name: str):
    job = SCRAPE_JOBS[job_id]
    base = Path(__file__).parent
    try:
        from scraper import generic_scrape, discover_participant_url
        job["status"] = "🔍 Katılımcı listesi kazınıyor..."
        companies = generic_scrape(url, fair_name=fair_name)
        used_url = url

        if not companies:
            # Verilen URL doğrudan liste değilse (ör. organizatörün ana sayfası),
            # sayfa içindeki linklerden katılımcı/exhibitor sayfasını otomatik ara.
            job["status"] = "🔎 Doğrudan bulunamadı, katılımcı sayfası otomatik aranıyor..."
            discovered = discover_participant_url(url)
            if discovered and discovered != url:
                job["status"] = "🔎 Katılımcı sayfası bulundu, deneniyor..."
                companies = generic_scrape(discovered, fair_name=fair_name)
                used_url = discovered

        job["found"] = len(companies)
        job["scraped_url"] = used_url
        if not companies:
            job["status"] = "hata"
            job["error"] = ("Bu sayfada (ve otomatik bulunan alt sayfalarda) firma listesi "
                            "bulunamadı. URL gerçekten katılımcı listesi sayfası mı? Login "
                            "gerektiren platformlar kazınamaz — bu durumda uyarı loglanır.")
            return
        job["added"] = ds.append_companies(companies)
        job["status"] = "🌐 İletişim bilgileri toplanıyor (arka planda sürer)..."
        # Zenginleştirme ayrı süreçte — canlı senkron sayesinde sonuçlar
        # arayüze parça parça düşer (enricher her 20 firmada bir kaydediyor).
        log_path = base / "output" / f"_fairjob_{job_id}.log"
        with open(log_path, "w", encoding="utf-8") as lf:
            proc = subprocess.Popen(
                [sys.executable, "enricher.py", "--fair", fair_name,
                 "--only-empty", "--no-maps"],
                cwd=str(base), stdout=lf, stderr=subprocess.STDOUT,
                env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"},
            )
            proc.wait(timeout=7200)
        job["status"] = "✅ Tamamlandı"
    except Exception as e:
        job["status"] = "hata"
        job["error"] = str(e)


@app.post("/api/fairs/scrape")
def api_fair_scrape():
    body = request.get_json(silent=True) or {}
    url = (body.get("url") or "").strip()
    name = (body.get("name") or "").strip()
    if not url.startswith("http") or not name:
        return jsonify({"error": "Geçerli bir URL ve fuar adı gerekli"}), 400
    job_id = uuid.uuid4().hex[:8]
    SCRAPE_JOBS[job_id] = {
        "status": "başlatılıyor", "url": url, "name": name,
        "found": 0, "added": 0, "error": "",
    }
    threading.Thread(target=_run_scrape_job, args=(job_id, url, name), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.get("/api/fairs/scrape/<job_id>")
def api_fair_scrape_status(job_id):
    job = SCRAPE_JOBS.get(job_id)
    if job is None:
        return jsonify({"error": "İş bulunamadı"}), 404
    return jsonify(job)


@app.get("/api/calendar")
def api_calendar():
    # Cache dosyası yoksa ilk seferde otomatik çeker (kullanıcı boş ekran görmesin).
    data = tobb_takvim.load_cached()
    if not data["fairs"]:
        try:
            data = tobb_takvim.scrape_and_save()
        except Exception as e:
            return jsonify({"error": str(e), "fetched_at": "", "count": 0, "fairs": []}), 502
    return jsonify(data)


@app.post("/api/calendar/refresh")
def api_calendar_refresh():
    try:
        data = tobb_takvim.scrape_and_save()
    except Exception as e:
        return jsonify({"error": f"TOBB sitesine ulaşılamadı: {e}"}), 502
    return jsonify(data)


@app.get("/api/export")
def api_export():
    # enricher'ın save_excel'i fuar-bazlı sheetler + Özet üretir.
    # leads.xlsx kullanıcının Excel'inde açık (kilitli) olabilir — o yüzden
    # zaman damgalı ayrı bir dosyaya yazıp onu gönderiyoruz.
    from datetime import datetime
    from enricher import save_excel, OUTPUT_DIR

    data = ds.load()
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    export_path = Path(OUTPUT_DIR) / f"leads_export_{ts}.xlsx"
    save_excel(data, path=export_path)
    # Eski export dosyalarını temizle (son 3 kalsın)
    old = sorted(Path(OUTPUT_DIR).glob("leads_export_*.xlsx"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    for p in old[3:]:
        try:
            p.unlink()
        except OSError:
            pass
    return send_file(
        export_path.resolve(),
        as_attachment=True,
        download_name=f"reyart_leads_{ts}.xlsx",
    )


@app.get("/api/export/csv")
def api_export_csv():
    """Tüm veriyi tek CSV olarak indirir. Türkçe Excel uyumu için UTF-8 BOM
    (utf-8-sig) + noktalı virgül ayracı kullanılır (TR locale Excel'i virgülü
    ayraç saymaz, ç/ğ/ş bozulmasın diye BOM şart)."""
    import csv
    import io
    from datetime import datetime

    cols = [
        ("name", "Firma"), ("fair", "Fuar"), ("fuar_tarihi", "Fuar Tarihi"),
        ("mensei", "Menşei"), ("country", "Ülke"), ("sector", "Sektör"),
        ("website", "Website"), ("phone", "Telefon"), ("email", "E-posta"),
        ("linkedin", "LinkedIn"), ("priority", "Öncelik"),
        ("sicaklik", "Sıcaklık"), ("satis_skoru", "Satış Skoru"),
        ("guven_seviyesi", "Güven"), ("status", "Durum"),
        ("takip_tarihi", "Takip Tarihi"), ("note", "Not"),
    ]
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\r\n")
    w.writerow([h for _, h in cols])
    for c in ds.load():
        w.writerow([c.get(k, "") for k, _ in cols])

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return Response(
        "﻿" + buf.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename=reyart_leads_{ts}.csv"},
    )


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Reyart Lead Manager web arayüzü")
    parser.add_argument(
        "--lan", action="store_true",
        help="Ofis ağına aç (0.0.0.0) — ekip http://<bu-pc-ip>:5000 ile bağlanır. "
             "DİKKAT: auth yok, sadece güvenilen yerel ağda kullan.",
    )
    parser.add_argument("--port", type=int, default=5000)
    args = parser.parse_args()

    host = "0.0.0.0" if args.lan else "127.0.0.1"
    if args.lan:
        import socket
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            lan_ip = s.getsockname()[0]
            s.close()
            # Windows konsolu cp1252 olabilir — Türkçe karakter basma, çökme riski
            print(f"\n  Ekip baglanti adresi: http://{lan_ip}:{args.port}\n", flush=True)
        except OSError:
            print("\n  LAN IP tespit edilemedi; 'ipconfig' ile IPv4 adresine bak.\n", flush=True)
    app.run(host=host, port=args.port, debug=False, threaded=True)
