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
import os
import random
import re
import smtplib
import time
from datetime import date, datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

BASE_DIR = Path(__file__).parent
TEMPLATES_FILE = BASE_DIR / "mail_templates.json"
CONFIG_FILE = BASE_DIR / "mail_config.json"
LOG_FILE = BASE_DIR / "output" / "mail_log.jsonl"

# .env dosyasından SMTP ayarları (Render'da kullanılır)
def _load_env_smtp():
    """Render'da .env veya ortam değişkenlerinden SMTP ayarlarını okur."""
    return {
        "smtp_server": os.environ.get("SMTP_SERVER", "smtp.office365.com"),
        "smtp_port": int(os.environ.get("SMTP_PORT", "587")),
        "smtp_username": os.environ.get("SMTP_USERNAME", ""),
        "smtp_password": os.environ.get("SMTP_PASSWORD", ""),
    }

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
    if pythoncom is None:
        raise RuntimeError("pywin32 (pythoncom) kurulu değil — bu fonksiyon yalnızca Windows'ta çalışır.")
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
    if pythoncom is None:
        # Linux/Render: SMTP durumunu kontrol et
        smtp = _load_env_smtp()
        if smtp["smtp_username"] and smtp["smtp_password"]:
            return {"configured": True, "email": smtp["smtp_username"], "hint": "SMTP (Render/Linux)", "method": "smtp"}
        return {"configured": False, "email": cfg["outlook_account"], "hint": "pywin32 kurulu değil (Linux/Render). SMTP ayarları .env'de yok.", "method": "smtp"}
    try:
        _outlook, account = _connect_outlook(cfg["outlook_account"])
        return {"configured": True, "email": account.SmtpAddress, "hint": "", "method": "outlook"}
    except RuntimeError as e:
        return {"configured": False, "email": cfg["outlook_account"], "hint": str(e), "method": "outlook"}
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
    """Tek firma için mail hazırlar.
    - Windows'da: classic Outlook'ta HAZIR taslak açar (Display), GÖNDERMEZ.
    - Linux/Render'da: SMTP ile doğrudan gönderir.
    """
    templates = load_templates()
    if template_id not in templates:
        raise ValueError(f"Şablon yok: {template_id}")
    email = (company.get("email") or "").strip()
    if not email or not EMAIL_RE.match(email):
        raise ValueError("Bu firmanın geçerli bir e-posta adresi yok")

    cfg = load_config()
    msg_data = render(templates[template_id], company)

    smtp_cfg = _load_env_smtp()
    use_smtp = pythoncom is None and smtp_cfg["smtp_username"] and smtp_cfg["smtp_password"]

    if use_smtp:
        # SMTP ile gönder (Render/Linux)
        smtp_server = smtplib.SMTP(smtp_cfg["smtp_server"], smtp_cfg["smtp_port"])
        smtp_server.starttls()
        smtp_server.login(smtp_cfg["smtp_username"], smtp_cfg["smtp_password"])
        try:
            msg = MIMEMultipart()
            msg["From"] = smtp_cfg["smtp_username"]
            msg["To"] = email
            msg["Subject"] = msg_data["subject"]
            msg.attach(MIMEText(msg_data["body"], "plain", "utf-8"))
            smtp_server.send_message(msg)
            return {"sent": True, "account": smtp_cfg["smtp_username"], "to": email, "method": "smtp"}
        finally:
            smtp_server.quit()

    if pythoncom is None:
        raise RuntimeError("pywin32 (pythoncom) kurulu değil ve SMTP ayarları yok. Mail gönderilemez.")
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
        return {"opened": True, "account": account.SmtpAddress, "to": email, "method": "outlook"}
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
    smtp_cfg = _load_env_smtp()
    use_smtp = pythoncom is None and smtp_cfg["smtp_username"] and smtp_cfg["smtp_password"]
    if not dry_run:
        if use_smtp:
            smtp_server = smtplib.SMTP(smtp_cfg["smtp_server"], smtp_cfg["smtp_port"])
            smtp_server.starttls()
            smtp_server.login(smtp_cfg["smtp_username"], smtp_cfg["smtp_password"])
        else:
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
                if use_smtp:
                    # SMTP ile doğrudan gönder (Render/Linux)
                    msg = MIMEMultipart()
                    msg["From"] = smtp_cfg["smtp_username"]
                    msg["To"] = email
                    msg["Subject"] = msg_data["subject"]
                    msg.attach(MIMEText(msg_data["body"], "plain", "utf-8"))
                    smtp_server.send_message(msg)
                else:
                    # Outlook COM ile gönder (Windows)
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
        if not dry_run and pythoncom is not None:
            pythoncom.CoUninitialize()
        if not dry_run and use_smtp:
            smtp_server.quit()

    return results
