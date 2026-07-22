#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Contact Finder Worker — sadece Google Maps üzerinden arama yapar.
Argüman 1: JSON config dosya yolu
Argüman 2: Çıktı JSONL dosya yolu
"""
import json, re, sys, time
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import quote_plus

PHONE_RE = re.compile(
    r"(?:\+90[\s\-]?(?:\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2})"
    r"|0[\s\-]?(?:\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2})"
    r"|\(\d{3,4}\)[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2})"
)
SPAM_PHONES = {"05000000000", "05550000000", "02120000000"}
LEGAL_RE = re.compile(
    r"\b(A\.S\.|A\.Ş\.|LTD\.STI\.|LTD\.ŞTİ\.|SAN\.|TİC\.|TIC\.|"
    r"VE TİC\.|DIŞ TİC\.|DIS TIC\.|İTH\.|İHR\.|PAZ\.)\b", re.IGNORECASE)

# Maps'te tıklanan ilk sonuç sorgudakiyle alakasız bir yer olabiliyor
# (özellikle kısa/özel isimli firmalarda). Kabul etmeden önce yerin adını
# sorgudaki firma adıyla kelime örtüşmesiyle doğruluyoruz.
LEGAL_STOPWORDS = {
    "ve", "a.s.", "as", "a.ş.", "ltd", "sti", "şti", "san", "tic", "dis",
    "dış", "ith", "ihr", "ithalat", "ihracat", "pazarlama", "paz", "co",
    "inc", "gmbh", "kg", "srl", "spa", "ag", "sa", "sirketi", "şirketi",
    "anonim", "limited", "sanayi", "ticaret", "endustri", "endüstri",
    "turkiye", "türkiye", "turkey",
}
TITLE_SUFFIX_RE = re.compile(r"\s*[-·|]\s*Google\s*(Haritalar|Maps).*$", re.IGNORECASE)

# Token-overlap eşiği tutmayan ama yazım farkı/şube-marka varyasyonu olan
# eşleşmeler için ek fuzzy sinyal (difflib.SequenceMatcher.ratio).
FUZZY_RATIO_THRESHOLD = 0.6

# Sorguya körlemesine " Türkiye" eklememek için: companies.json'daki country
# alanı zaten Türkiye/Turkey/TR ise davranış aynı kalır, yabancı ülkelerde
# sorguya o ülke adı eklenir.
TR_ALIASES = {"turkiye", "türkiye", "turkey", "tr"}

_out_path = None


def _alnum_norm(s):
    """Fuzzy karşılaştırma için boşluk/noktalama atılmış, TR karakterleri sadeleştirilmiş hali."""
    return re.sub(r"[^a-z0-9]", "", _fold_tr(s or "").lower())


def query_suffix(country):
    """Maps sorgusuna eklenecek ülke eki. Ülke boş/Türkiye ise 'Türkiye', değilse ülkenin kendisi."""
    if not country or _fold_tr(country).lower().strip() in TR_ALIASES:
        return "Türkiye"
    return country


def clean_phone(raw):
    d = re.sub(r"\D", "", raw)
    if d.startswith("90") and len(d) == 12:
        d = "0" + d[2:]
    if len(d) == 10 and not d.startswith("0"):
        d = "0" + d
    return d if 10 <= len(d) <= 12 else ""


def short_name(name):
    s = LEGAL_RE.sub("", name).strip(" .,")
    return s if s else name


def _fold_tr(s):
    table = str.maketrans({
        "İ": "i", "I": "i", "ı": "i", "Ş": "s", "ş": "s", "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u", "Ö": "o", "ö": "o", "Ç": "c", "ç": "c",
    })
    return s.translate(table)


def sig_words(name):
    # ÖNEMLİ: fold_tr önce, .lower() sonra — "İ".lower() Python'da combining-dot
    # karakteri (U+0307) ekliyor ve kelimeleri stopword listesiyle eşleşmez yapıyor.
    s = _fold_tr(name).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    return {w for w in s.split() if w not in LEGAL_STOPWORDS and len(w) > 1}


def match_diagnostics(expected_name, place_title):
    """Token-overlap + fuzzy benzerlik detaylarını döndürür (eşleşme kararı + log/debug için).
    Token-overlap eşiği tutmazsa isim/başlık arasındaki fuzzy oran da ek sinyal olarak
    denenir (typo, kısaltma, şube/marka varyasyonlarını yakalamak için)."""
    if not place_title:
        return {"matched": False, "hits": 0, "wanted": 0, "ratio": 0.0,
                "reason": "yer adı boş (sayfa başlığı okunamadı)"}
    wanted = sig_words(expected_name)
    if not wanted:
        return {"matched": True, "hits": 0, "wanted": 0, "ratio": 1.0, "reason": ""}
    found = _fold_tr(place_title).lower()
    hits = sum(1 for w in wanted if w in found)
    threshold = max(1, len(wanted) // 2)
    token_ok = hits >= threshold
    ratio = SequenceMatcher(None, _alnum_norm(expected_name), _alnum_norm(place_title)).ratio()
    fuzzy_ok = ratio >= FUZZY_RATIO_THRESHOLD
    matched = token_ok or fuzzy_ok
    reason = "" if matched else (
        f"kelime örtüşmesi {hits}/{len(wanted)} (eşik {threshold}), fuzzy benzerlik {ratio:.2f}"
    )
    return {"matched": matched, "hits": hits, "wanted": len(wanted), "ratio": ratio, "reason": reason}


def place_matches(expected_name, place_title):
    """Google Maps'te açılan yerin adı, aranan firma adıyla örtüşüyor mu?"""
    return match_diagnostics(expected_name, place_title)["matched"]


def emit(obj):
    with open(_out_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")
        f.flush()


def emit_error(i, total, name, msg):
    emit({"type": "error", "i": i, "total": total, "isim": name, "msg": msg})


def search_maps(page, query, expected_name):
    """Google Maps'te ara, ilk sonucun telefon ve adresini döndür.
    İlk sonucun adı sorgudaki firma adıyla örtüşmüyorsa hiçbir şey döndürmez
    (alakasız yerin telefonunu firmaya yapıştırmamak için).
    Dönen 5. eleman: reddedilme sebebi (eşleşince "")."""
    phones = set()
    address = ""
    matched_title = ""
    try:
        url = f"https://www.google.com/maps/search/{quote_plus(query)}?hl=tr"
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        time.sleep(2)

        # İlk sonuca tıkla (liste görünümündeyse)
        first = page.query_selector('a[href*="/maps/place/"]')
        if first:
            first.click()
            time.sleep(2)

        # Açılan yerin adını al (sekme başlığı: "Yer Adı - Google Haritalar")
        raw_title = page.title() or ""
        matched_title = TITLE_SUFFIX_RE.sub("", raw_title).strip()
        diag = match_diagnostics(expected_name, matched_title)
        if not diag["matched"]:
            return phones, address, matched_title, False, diag["reason"]

        # Telefon: data-item-id="phone:tel:..."
        phone_el = page.query_selector('[data-item-id^="phone:tel:"]')
        if phone_el:
            raw = (phone_el.get_attribute("data-item-id") or "").replace("phone:tel:", "")
            p = clean_phone(raw)
            if p and p not in SPAM_PHONES:
                phones.add(p)

        # Fallback: aria-label içinde telefon
        if not phones:
            for el in page.query_selector_all('[aria-label]'):
                lbl = el.get_attribute("aria-label") or ""
                if "telefon" in lbl.lower() or "phone" in lbl.lower():
                    for m in PHONE_RE.finditer(lbl):
                        p = clean_phone(m.group())
                        if p and p not in SPAM_PHONES:
                            phones.add(p)

        # Fallback: tüm sayfa metni
        if not phones:
            content = page.inner_text("body")
            for m in PHONE_RE.finditer(content):
                p = clean_phone(m.group())
                if p and p not in SPAM_PHONES:
                    phones.add(p)

        # Adres
        addr_el = page.query_selector('[data-item-id="address"]')
        if addr_el:
            address = addr_el.inner_text().strip()

    except Exception as ex:
        raise ex

    return phones, address, matched_title, True, ""


def main():
    global _out_path
    cfg_path  = sys.argv[1]
    _out_path = sys.argv[2]
    cfg       = json.loads(Path(cfg_path).read_text(encoding="utf-8"))
    # "companies" varsa (isim + ülke) onu kullan; yoksa eski "names" listesiyle
    # geriye dönük uyumlu çalış (ülke bilinmiyorsa Türkiye varsayılır).
    if cfg.get("companies"):
        entries = [(c.get("name", ""), c.get("country", "")) for c in cfg["companies"]]
    else:
        entries = [(n, "") for n in cfg.get("names", [])]
    headless  = cfg.get("headless", False)
    delay     = cfg.get("delay", 1.5)
    proxy_str = cfg.get("proxy", None)

    proxy_cfg = None
    if proxy_str:
        # format: http://user:pass@host:port  or  http://host:port
        proxy_cfg = {"server": proxy_str}

    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        launch_args = {
            "headless": headless,
            "args": ["--disable-blink-features=AutomationControlled"],
        }
        if proxy_cfg:
            launch_args["proxy"] = proxy_cfg

        browser = pw.chromium.launch(**launch_args)
        ctx_args = {
            "user_agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            "locale": "tr-TR",
        }
        if proxy_cfg:
            ctx_args["proxy"] = proxy_cfg

        ctx  = browser.new_context(**ctx_args)
        page = ctx.new_page()
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        total = len(entries)
        for i, (name, country) in enumerate(entries, 1):
            sn           = short_name(name)
            suffix       = query_suffix(country)
            all_phones   = set()
            address      = ""
            matched_name = ""
            matched      = False
            reddedilenler = []

            try:
                phones, addr, title, ok, reason = search_maps(page, name + " " + suffix, name)
                if ok:
                    all_phones |= phones
                    matched_name = title
                    matched = True
                    if addr:
                        address = addr
                elif title:
                    reddedilenler.append({
                        "deneme": "tam isim", "sorgu": name + " " + suffix,
                        "bulunan_yer": title, "sebep": reason,
                    })
            except Exception as ex:
                emit_error(i, total, name, str(ex))

            # Kısa isimle ikinci deneme (eşleşme bulunamadıysa ve isim farklıysa)
            if not matched and sn != name:
                try:
                    phones, addr, title, ok, reason = search_maps(page, sn + " " + suffix, name)
                    if ok:
                        all_phones |= phones
                        matched_name = title
                        matched = True
                        if addr and not address:
                            address = addr
                    elif title:
                        reddedilenler.append({
                            "deneme": "kisa isim", "sorgu": sn + " " + suffix,
                            "bulunan_yer": title, "sebep": reason,
                        })
                except Exception as ex:
                    emit_error(i, total, name, f"[kısa isim] {ex}")

            emit({
                "type":       "result",
                "i":          i,
                "total":      total,
                "isim":       name,
                "ulke":       country,
                "sorgu_eki":  suffix,
                "eslesen_yer": matched_name,
                "eslesti":    matched,
                "emailler":   [],
                "telefonlar": sorted(all_phones),
                "kaynaklar":  [address] if address else [],
                # Reddedilen adaylar: eşleşmeyip atlanan yer(ler) + sebep — sessizce
                # atlamak yerine debug/manuel denetim için iz bırakır.
                "reddedilen_yerler": reddedilenler,
            })

            time.sleep(delay)

        browser.close()

    emit({"type": "done"})


if __name__ == "__main__":
    main()
