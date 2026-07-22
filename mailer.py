"""
Mail motoru — şablonlu, kişiselleştirilmiş toplu e-posta.

Gönderim, kurulu **klasik (classic) Outlook** masaüstü uygulaması üzerinden,
COM otomasyonu (pywin32) ile yapılır — SMTP şifresi/uygulama şifresi GEREKMEZ.
Outlook'un açık ve demir@reyartfuar.com hesabının o Outlook'a ekli olması yeterli.
Bu betik Outlook'un çalıştığı AYNI Windows makinesinde koşmalı (webapp.py de
zaten aynı PC'de çalışıyor).

- Şablonlar: mail_templates.json ({firma}, {fuar}, {ulke} yer tutucuları).
- Ayarlar: mail_config.json (opsiyonel — yoksa varsayılanlar kullanılır):
    {
      "outlook_account": "demir@reyartfuar.com",
      "sender_name": "Reyart Fuar Stand Tasarım",
      "daily_limit": 100
    }
- Her gönderim output/mail_log.jsonl'a loglanır; aynı firmaya aynı şablon
  ikinci kez gönderilmez (çift mail koruması).

Not: Outlook, dışarıdan (COM ile) e-posta gönderildiğinde bazı güvenlik
yapılandırmalarında "Bir program sizin adınıza e-posta göndermeye çalışıyor"
uyarısı gösterebilir — çıkarsa "İzin Ver" demen yeterli, ekranı başında
olduğun için (bu betik senin bilgisayarında çalışıyor) görüp onaylayabilirsin.
"""

import json
import random
import re
import time
from datetime import date, datetime
from pathlib import Path

import pythoncom
import win32com.client

BASE_DIR = Path(__file__).parent
TEMPLATES_FILE = BASE_DIR / "mail_templates.json"
CONFIG_FILE = BASE_DIR / "mail_config.json"
LOG_FILE = BASE_DIR / "output" / "mail_log.jsonl"

DEFAULT_OUTLOOK_ACCOUNT = "demir@reyartfuar.com"
SEND_DELAY_SECONDS = 2  # ard arda gönderimde Outlook'u/alıcı sunucuları boğmamak için

# Katalog dosyaları (masaüstü, OneDrive altında) — SADECE compose_draft'ta
# (tekli, elle gözden geçirilen taslak) eklenir; toplu cold gönderimde eklenmez.
DESKTOP_DIR = Path.home() / "OneDrive" / "Desktop"
CATALOG_ATTACHMENTS = [
    DESKTOP_DIR / "Reyart Catalog TR.pdf",
    DESKTOP_DIR / "Reyart Catalog ENG.pdf",
]


def _existing_attachments() -> list[Path]:
    """Var olan katalog dosyalarını döndürür; taşınmış/silinmişse sessizce atlar
    (log'a düşer, gönderim durmaz)."""
    found = [p for p in CATALOG_ATTACHMENTS if p.exists()]
    missing = [p for p in CATALOG_ATTACHMENTS if not p.exists()]
    for p in missing:
        print(f"[mailer] UYARI: katalog dosyası bulunamadı, eklenemedi: {p}")
    return found


def load_templates() -> dict:
    return json.loads(TEMPLATES_FILE.read_text(encoding="utf-8"))


def load_config() -> dict:
    cfg = {}
    if CONFIG_FILE.exists():
        try:
            cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            cfg = {}
    return {
        "outlook_account": cfg.get("outlook_account", DEFAULT_OUTLOOK_ACCOUNT),
        "sender_name": cfg.get("sender_name", "Reyart Fuar Stand Tasarım"),
        "daily_limit": cfg.get("daily_limit", 100),
    }


def _find_account(ns, email: str):
    email = email.lower().strip()
    for acc in ns.Accounts:
        if acc.SmtpAddress.lower() == email:
            return acc
    return None


def _connect_outlook(account_email: str):
    """Classic Outlook'a COM ile bağlanır, istenen hesabı bulur.
    Başarısızsa RuntimeError fırlatır (Türkçe, arayüzde doğrudan gösterilebilir)."""
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        account = _find_account(ns, account_email)
        if account is None:
            names = ", ".join(a.SmtpAddress for a in ns.Accounts) or "(hiç hesap yok)"
            raise RuntimeError(
                f"'{account_email}' hesabı classic Outlook'ta bulunamadı. "
                f"Outlook'a ekli hesaplar: {names}"
            )
        return outlook, account
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(
            f"Classic Outlook'a bağlanılamadı — Outlook'un açık olduğundan emin ol. ({e})"
        )


def config_status() -> dict:
    cfg = load_config()
    try:
        _outlook, account = _connect_outlook(cfg["outlook_account"])
        return {"configured": True, "email": account.SmtpAddress, "hint": ""}
    except RuntimeError as e:
        return {"configured": False, "email": cfg["outlook_account"], "hint": str(e)}
    finally:
        pythoncom.CoUninitialize()


def render(template: dict, company: dict) -> dict:
    """Şablondaki yer tutucuları firma verisiyle doldurur."""
    fields = {
        "firma": company.get("name", ""),
        "fuar": company.get("fair", ""),
        "ulke": company.get("country", ""),
    }

    def fill(text: str) -> str:
        for k, v in fields.items():
            text = text.replace("{" + k + "}", v)
        return text

    return {
        "to": company.get("email", ""),
        "subject": fill(template["subject"]),
        "body": fill(template["body"]),
    }


# ── Gönderim log'u / çift mail koruması ──────────────────────────────────────

def _load_sent_keys() -> set[str]:
    if not LOG_FILE.exists():
        return set()
    keys = set()
    with open(LOG_FILE, encoding="utf-8") as f:
        for line in f:
            try:
                o = json.loads(line)
                if o.get("ok"):
                    keys.add(o["company_id"] + "|" + o["template"])
            except (json.JSONDecodeError, KeyError):
                continue
    return keys


def _log(entry: dict):
    LOG_FILE.parent.mkdir(exist_ok=True)
    entry["ts"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def sent_today_count() -> int:
    if not LOG_FILE.exists():
        return 0
    today = date.today().isoformat()
    n = 0
    with open(LOG_FILE, encoding="utf-8") as f:
        for line in f:
            try:
                o = json.loads(line)
                if o.get("ok") and o.get("ts", "").startswith(today):
                    n += 1
            except json.JSONDecodeError:
                continue
    return n


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")


def compose_draft(company: dict, template_id: str) -> dict:
    """Tek firma için classic Outlook'ta HAZIR bir taslak açar (Display),
    GÖNDERMEZ. mailto: linkinin Windows'un varsayılan mail uygulamasına (ör.
    yanlış hesaplı yeni Outlook) gitmesi sorununu bypass eder — COM ile
    doğrudan classic Outlook'u açar, hesap zaten doğru (SendUsingAccount)
    seçili gelir; kullanıcı gözden geçirip kendi eliyle gönderir."""
    templates = load_templates()
    if template_id not in templates:
        raise ValueError(f"Şablon yok: {template_id}")
    email = (company.get("email") or "").strip()
    if not email or not EMAIL_RE.match(email):
        raise ValueError("Bu firmanın geçerli bir e-posta adresi yok")

    cfg = load_config()
    msg_data = render(templates[template_id], company)

    pythoncom.CoInitialize()
    try:
        outlook, account = _connect_outlook(cfg["outlook_account"])
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = msg_data["to"]
        mail.Subject = msg_data["subject"]
        mail.Body = msg_data["body"]
        mail.SendUsingAccount = account
        for path in _existing_attachments():
            mail.Attachments.Add(str(path))
        mail.Display(False)  # modal olmayan pencere — Flask isteğini kilitlemez
        return {"opened": True, "account": account.SmtpAddress, "to": email}
    finally:
        pythoncom.CoUninitialize()


def send_bulk(companies: list[dict], template_id: str, dry_run: bool = True,
              delay_range: tuple[float, float] | None = None) -> dict:
    """Seçili firmalara şablonlu mail gönderir (classic Outlook üzerinden).

    dry_run=True → Outlook'a hiç dokunmaz, ne gönderileceğini raporlar.
    delay_range=(min_sn, max_sn) → mailler arası RASTGELE bekleme; verilmezse
    sabit SEND_DELAY_SECONDS. Toplu cold gönderimde patlama deseni spam
    filtrelerine yakalanmasın diye güne yaymak için kullanılır.
    Dönen: {sent: [...], skipped: [...], errors: [...], dry_run: bool}
    """
    templates = load_templates()
    if template_id not in templates:
        raise ValueError(f"Şablon yok: {template_id}")
    template = templates[template_id]
    cfg = load_config()

    already_sent = _load_sent_keys()
    daily_limit = cfg["daily_limit"]
    sent_today = sent_today_count()

    results = {"sent": [], "skipped": [], "errors": [], "dry_run": dry_run}

    # Toplu/cold gönderimde katalog PDF'i EKLENMEZ — tanımayan alıcıya ~2MB
    # ek klasik spam sinyali (Demir'in kararı, Tem 2026). Katalog, cevap veren
    # firmaya compose_draft (tekli taslak) akışıyla gönderilir.
    outlook = account = None
    if not dry_run:
        outlook, account = _connect_outlook(cfg["outlook_account"])  # hata varsa burada patlar

    try:
        for c in companies:
            cid = c.get("id", "")
            name = c.get("name", "?")
            email = (c.get("email") or "").strip()

            if not email or not EMAIL_RE.match(email):
                results["skipped"].append({"name": name, "reason": "geçerli e-posta yok"})
                continue
            if cid + "|" + template_id in already_sent:
                results["skipped"].append({"name": name, "reason": "bu şablon daha önce gönderilmiş"})
                continue
            if not dry_run and sent_today >= daily_limit:
                results["skipped"].append({"name": name, "reason": f"günlük limit ({daily_limit}) doldu"})
                continue

            msg_data = render(template, c)
            if dry_run:
                results["sent"].append({"name": name, "to": email, "subject": msg_data["subject"]})
                continue

            try:
                mail = outlook.CreateItem(0)  # olMailItem
                mail.To = email
                mail.Subject = msg_data["subject"]
                mail.Body = msg_data["body"]
                mail.SendUsingAccount = account
                mail.Send()

                _log({"ok": True, "company_id": cid, "name": name, "to": email,
                      "template": template_id})
                results["sent"].append({"name": name, "to": email, "subject": msg_data["subject"]})
                sent_today += 1
                if delay_range:
                    time.sleep(random.uniform(*delay_range))
                else:
                    time.sleep(SEND_DELAY_SECONDS)
            except Exception as e:
                _log({"ok": False, "company_id": cid, "name": name, "to": email,
                      "template": template_id, "error": str(e)})
                results["errors"].append({"name": name, "error": str(e)})
    finally:
        if not dry_run:
            pythoncom.CoUninitialize()

    return results
