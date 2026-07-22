"""
Reyart Lead Scraper — Adım 1
KOMATEK 2026 ve WIN EURASIA 2026 katılımcı listelerini çeker,
output/companies.json dosyasına kaydeder.
"""

import json
import re as _re
import time
import random
import logging
import urllib.parse as _urlparse
import urllib3
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# SSL uyarılarını bastır (kurumsal proxy/Windows CA sorunu)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)
COMPANIES_FILE = OUTPUT_DIR / "companies.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
}

# Küresel markalar — öncelik ⭐⭐⭐
PRIORITY_3_KEYWORDS = {
    "caterpillar", "cat", "liebherr", "komatsu", "hyundai", "bobcat",
    "hidromek", "sandvik", "metso", "develon", "new holland", "case",
    "manitou", "kobelco", "jungheinrich", "trumpf", "volvo", "doosan",
    "hitachi", "jcb", "terex", "atlas copco", "epiroc", "wirtgen",
    "dynapac", "hamm", "vogele", "kleemann", "powerscreen", "sany",
    "xcmg", "zoomlion", "liugong", "shantui", "sumitomo", "kubota",
    "takeuchi", "mitsubishi", "furukawa", "putzmeister", "schwing",
    "bomag", "vibromax", "ammann", "finlay", "powerscreen", "linde",
    "crown", "toyota", "yale", "hyster", "combilift", "merlo",
    "manitou", "faresin", "lgmg", "sinoboom",
}

def _fold_tr(s: str) -> str:
    """Türkçe karakterleri ASCII'ye katlar. ÖNEMLİ: .lower()'dan ÖNCE çağrılmalı —
    Python'da 'İ'.lower() combining-dot (U+0307) ekleyip karşılaştırmaları bozuyor,
    'I'.lower() da 'ı' değil 'i' verdiği için 'KATILIMCI' junk filtresini kaçırıyordu."""
    table = str.maketrans({
        "İ": "i", "I": "i", "ı": "i", "Ş": "s", "ş": "s", "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u", "Ö": "o", "ö": "o", "Ç": "c", "ç": "c",
    })
    return s.translate(table)


def _norm(s: str) -> str:
    return _fold_tr(s).lower().strip()


# Menü, cookie, navigasyon metinleri — bunları filtrele
JUNK_NAMES = {
    "home", "about", "contact", "events", "access", "layout", "plans",
    "product groups", "sales offices", "sponsorship", "participation",
    "technical sessions", "kvkk", "privacy", "personal data",
    "cookie", "gerekli", "bcookie", "ci_session", "iabv2",
    "informing", "bilgiler", "hakkında", "hakkimizda",
    "list of komatek 2026 participants", "list of komatek 2024 participants",
    "live construction machinery demonstrations",
    "machine operator competitions",
    "megatrends 2030 tepav report",
    "tobb megatrends 2030 launch meeting",
    "megatrends 2030",
    "megatrends eventsmegatrends",
    "komatek eventslive",
    "exhibitor's manual", "exhibitor portal",
    "information & pricing", "terms of contract",
    "katılımcı 2026", "rıza",
    # WIN EURASIA fallback'inden sızan gerçek junk örnekleri (2026-07 temizliği)
    "turkish", "english", "exhibition calendar", "organizers", "venues",
    "add my event", "contact us", "istanbul", "art", "general", "books",
    "education", "chemistry", "optics", "travel",
}

# Ülke isimleri — fuar sitelerinde çok yaygın bir yanlış-pozitif kaynağı:
# "ziyaretçi geldiği ülkeler" bayrak galerisi / "ihracat yaptığımız ülkeler"
# listesi, tekrar eden kart yapısı yüzünden firma listesi sanılabiliyor
# (örn. F İstanbul'daki div.flag > div.label grid'i — 18 ülke adı firma
# diye companies.json'a girmişti, bu yüzden eklendi).
COUNTRY_NAMES = {
    "turkiye", "turkey", "almanya", "germany", "fransa", "france", "italya", "italy",
    "ispanya", "spain", "portekiz", "portugal", "hollanda", "netherlands", "belcika",
    "belgium", "avusturya", "austria", "isvicre", "switzerland", "polonya", "poland",
    "cek cumhuriyeti", "czech republic", "slovakya", "slovakia", "macaristan", "hungary",
    "romanya", "romania", "bulgaristan", "bulgaria", "yunanistan", "greece", "sirbistan",
    "serbia", "hirvatistan", "croatia", "ukrayna", "ukraine", "rusya", "russia",
    "beyaz rusya", "belarus", "moldova", "ingiltere", "united kingdom", "uk", "irlanda",
    "ireland", "danimarka", "denmark", "isvec", "sweden", "norvec", "norway", "finlandiya",
    "finland", "abd", "usa", "united states", "kanada", "canada", "meksika", "mexico",
    "brezilya", "brazil", "arjantin", "argentina", "sili", "chile", "peru", "kolombiya",
    "colombia", "misir", "egypt", "cezayir", "algeria", "fas", "morocco", "tunus", "tunisia",
    "libya", "guney afrika", "south africa", "nijerya", "nigeria", "kenya", "etiyopya",
    "ethiopia", "cin", "china", "japonya", "japan", "guney kore", "south korea", "hindistan",
    "india", "pakistan", "bangladeş", "bangladesh", "vietnam", "tayland", "thailand",
    "malezya", "malaysia", "endonezya", "indonesia", "filipinler", "philippines",
    "singapur", "singapore", "iran", "irak", "iraq", "urdun", "jordan", "suriye", "syria",
    "lubnan", "lebanon", "israil", "israel", "suudi arabistan", "saudi arabia", "bae",
    "uae", "united arab emirates", "katar", "qatar", "kuveyt", "kuwait", "bahreyn",
    "bahrain", "umman", "oman", "yemen", "afganistan", "afghanistan", "azerbaycan",
    "azerbaijan", "gurcistan", "georgia", "ermenistan", "armenia", "kazakistan",
    "kazakhstan", "ozbekistan", "uzbekistan", "turkmenistan", "kirgizistan", "kyrgyzstan",
    "tacikistan", "tajikistan", "avustralya", "australia", "yeni zelanda", "new zealand",
    "el salvador", "guatemala", "honduras", "nikaragua", "nicaragua", "panama", "kuba",
    "cuba", "dominik cumhuriyeti", "dominican republic",
}
JUNK_NAMES |= COUNTRY_NAMES

# JUNK_NAMES'i Türkçe-katlanmış normal formda tut — is_valid_company da
# aynı normalizasyonla karşılaştırır (case-folding kaçaklarını önler).
JUNK_NAMES = {_norm(n) for n in JUNK_NAMES}

# Kategori/etiket kalıpları: fallback sayfalarından sızan fuar istatistikleri
# (örn. "700+ EXHIBITORS", "40,243 VISITORS", "55.000 sqm EXHIBITION AREA")
# ve virgülle ayrılmış sektör taksonomileri (örn. "Transport, Logistics, Maritime").
JUNK_REGEXES = [
    _re.compile(r"^[\d.,+]+\s", _re.IGNORECASE),          # sayıyla başlayan istatistik satırı
    _re.compile(r"^[A-Za-z &/\-]+,\s*[A-Za-z &/\-]+(,\s*[A-Za-z &/\-]+)+$"),  # "X, Y, Z" taksonomi
    _re.compile(r"^[\d.,]+\s*%?$"),                        # salt sayı/yüzde ("2024", "27.6%")
]

# Nav menüsü / pazarlama sayfası kelimeleri — bir isimde 2+ tanesi geçiyorsa
# o bir firma değil, site menüsüdür (2026-07 toplu tarama temizliğinden).
NAV_KEYWORDS = {
    "about", "home", "events", "contact", "exhibitions", "corporate",
    "upcoming", "mission", "media", "sponsors", "participants", "vision",
    "english", "deutsch", "turkce", "gallery", "downloads", "register",
}

# Gerçek firma adlarının özellikleri — karşılaştırmalar _norm() (TR-katlanmış
# lowercase) üzerinden yapılır, "KATILIMCI" vs "katılımcı" kaçağı olmaz.
_EXTRA_JUNK = {_norm(n) for n in {
    "lift", "forklift", "access", "home", "events",
    "contact", "participation", "rıza", "bilgiler", "hakkında",
}}

JUNK_PATTERNS = [
    lambda s, n: len(s) < 3,
    lambda s, n: len(s) > 120,
    lambda s, n: n in JUNK_NAMES,
    lambda s, n: n in _EXTRA_JUNK,
    lambda s, n: n.startswith("cookie"),
    lambda s, n: n.startswith("[#"),
    lambda s, n: s.lstrip().startswith(('"', "“", "'", "‘")),  # alıntı/testimonial
    lambda s, n: "bu saglayici tarafindan" in n,
    lambda s, n: "toplanan verilerin" in n,
    lambda s, n: "tercih fuari" in n or "tercih gunleri" in n,  # organizatörün kendi fuar menüsü
    lambda s, n: len(set(_re.findall(r"[a-z]+", n)) & NAV_KEYWORDS) >= 2,
    lambda s, n: any(rx.match(s) for rx in JUNK_REGEXES),
]


def is_valid_company(name: str) -> bool:
    name = name.strip()
    n = _norm(name)
    return not any(p(name, n) for p in JUNK_PATTERNS)


def assign_priority(name: str) -> str:
    lower = name.lower()
    if any(k in lower for k in PRIORITY_3_KEYWORDS):
        return "⭐⭐⭐"
    intl_signals = ["gmbh", " ag ", " sa ", " spa", " bv ", "ltd", "inc",
                    "corp", "group", "international", "global", "co.,ltd",
                    "co., ltd", "machinery", "equipment", "industries",
                    "systems", "solutions", "technology", "holding"]
    if any(k in lower for k in intl_signals):
        return "⭐⭐"
    return "⭐"


def make_company(name, country="", sector="", fair="") -> dict:
    return {
        "name": name.strip(),
        "country": country,
        "sector": sector,
        "fair": fair,
        "website": "",
        "phone": "",
        "email": "",
        "linkedin": "",
        "priority": assign_priority(name),
        "status": "⬜ Aranmadı",
        "note": "",
    }


def scrape_komatek() -> list[dict]:
    url = "https://komatekfuar.com/en/list-of-komatek-2026-participants/"
    log.info("KOMATEK listesi çekiliyor: %s", url)

    try:
        resp = requests.get(url, headers=HEADERS, timeout=30, verify=False)
        resp.raise_for_status()
    except Exception as e:
        log.error("KOMATEK isteği başarısız: %s", e)
        return []

    soup = BeautifulSoup(resp.text, "lxml")
    companies = []
    seen = set()

    # Sayfa yapısı: menü UL'leri küçük (2-10 li), firma UL'leri büyük (>20 li)
    # UL#8 ve sonrası (HOME, PARTICIPATION, KVKK) atlama eşiği: en fazla 10 li
    all_uls = soup.find_all("ul")
    MENU_CLASSES = {"sub-menu-custom", "sub-menu-inner", "sub-menu", "menu"}

    for ul in all_uls:
        ul_classes = set(ul.get("class", []))
        # Menü sınıfları olan UL'leri atla
        if ul_classes & MENU_CLASSES:
            continue
        lis = ul.find_all("li", recursive=False)
        # Çok küçük listeler menü/navigasyon olabilir, atla
        if len(lis) < 10:
            continue
        for li in lis:
            text = li.get_text(separator=" ", strip=True)
            if is_valid_company(text) and text not in seen:
                seen.add(text)
                companies.append(make_company(text, fair="KOMATEK 2026"))

    log.info("KOMATEK: %d firma bulundu", len(companies))
    return companies


def scrape_win_eurasia() -> list[dict]:
    url = "https://platform.win-eurasia.com/participants?new"
    log.info("WIN EURASIA listesi çekiliyor (Playwright): %s", url)
    companies = []

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="tr-TR",
                ignore_https_errors=True,
            )
            page = ctx.new_page()
            page.goto(url, timeout=60_000, wait_until="networkidle")

            for _ in range(15):
                page.keyboard.press("End")
                time.sleep(1.2)

            html = page.content()
            browser.close()

        soup = BeautifulSoup(html, "lxml")
        seen = set()

        for sel in ["[class*='exhibitor']", "[class*='participant']",
                    "[class*='company']", "[class*='card']", "article"]:
            for el in soup.select(sel):
                name_el = el.find(["h2", "h3", "h4", "strong"])
                name = (name_el.get_text(strip=True) if name_el
                        else el.get_text(strip=True)[:80])
                if is_valid_company(name) and name not in seen:
                    seen.add(name)
                    country_el = el.find(class_=lambda c: c and "country" in c.lower())
                    country = country_el.get_text(strip=True) if country_el else ""
                    companies.append(make_company(name, country=country, fair="WIN EURASIA 2026"))

        if len(companies) < 5:
            log.warning("WIN EURASIA platformdan az veri, yedekler deneniyor...")
            companies.extend(_win_eurasia_fallback())

    except Exception as e:
        log.error("WIN EURASIA Playwright hatası: %s", e)
        companies = _win_eurasia_fallback()

    log.info("WIN EURASIA: %d firma bulundu", len(companies))
    return companies


def _win_eurasia_fallback() -> list[dict]:
    # ESKİ fallback expointurkey.org / tradefairdates.com landing sayfalarındaki
    # her <li>'yi firma sanıyordu — 118 kayıtlık junk üretti (menü, şehir listesi,
    # sektör taksonomisi, fuar istatistikleri). O sayfalarda katılımcı listesi YOK;
    # sahte veri üretmektense boş dönüp durumu açıkça loglamak daha doğru.
    log.warning(
        "WIN EURASIA katılımcı listesi platformdan çekilemedi. Yedek kaynak YOK "
        "(landing sayfaları gerçek liste içermiyor). Elle kontrol: "
        "https://platform.win-eurasia.com/participants"
    )
    return []


# ── Genel (herhangi bir fuar sitesi) tarama altyapısı ────────────────────────
# generic_scrape() artık: (1) çeşitli ağ hatalarına karşı retry+backoff yapar,
# (2) yanlış/eksik charset yüzünden bozuk Türkçe metne karşı encoding'i düzeltir,
# (3) login duvarlarını erkenden tespit edip sahte/boş veri üretmez,
# (4) UL/class-selector/tablo stratejilerine ek olarak tekrar eden div/article
#     "kart" bloklarını da tanır (WordPress/React grid düzenleri için),
# (5) bulduğu aday listelerini "unique-ratio" kalite kapısından geçirir,
# (6) ?page=N tarzı sayfalamayı otomatik takip eder (tek sayfa sanıp yarım
#     veri toplamayı önler),
# (7) Tüyap-CMS'i (div.brand-item) tanırsa tuyap_platform_scrape.py'nin
#     doğrulanmış parser'ını kullanır — bu platformda onlarca fuar sitesi var.

LOGIN_WALL_SIGNS = [
    "giris yap", "uye girisi", "please log in", "please login", "sign in to",
    "401 unauthorized", "access denied", "for members only", "b2b girisi",
    "kullanici adi ve sifre", "log in to continue", "you must be logged in",
]

PARTICIPANT_LINK_KEYWORDS = [
    "katilimci listesi", "katilimcilar", "exhibitor list", "exhibitors",
    "participant list", "participants", "kayitli firmalar", "uye firmalar",
    "firma listesi", "sergileyen firmalar", "exhibitor directory",
    "list of participants", "katilimci firmalar",
]

PAGE_PARAM_RE = _re.compile(r"([?&]page=)(\d+)")
MENU_CLASSES = {"sub-menu-custom", "sub-menu-inner", "sub-menu", "menu",
                "nav", "navigation", "footer-menu", "header-menu"}


def _fetch(url: str, timeout: int = 30, retries: int = 2) -> str:
    """Retry + backoff'lu GET; charset'i (apparent_encoding) düzeltir ki
    server charset bildirmediğinde Türkçe karakterler bozulmasın."""
    last_err = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
            if not r.encoding or r.encoding.lower() in ("iso-8859-1", "windows-1252"):
                r.encoding = r.apparent_encoding
            if r.status_code == 200:
                return r.text
            last_err = f"HTTP {r.status_code}"
        except Exception as e:
            last_err = str(e)
        if attempt < retries:
            time.sleep(1.5 * (attempt + 1))
    log.warning("İstek başarısız (%d deneme): %s — %s", retries + 1, url, last_err)
    return ""


def _fetch_playwright(url: str, scroll_rounds: int = 10) -> str:
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=HEADERS["User-Agent"], locale="tr-TR", ignore_https_errors=True,
            )
            page = ctx.new_page()
            page.goto(url, timeout=60_000, wait_until="networkidle")
            for _ in range(scroll_rounds):
                page.keyboard.press("End")
                time.sleep(1.0)
            html = page.content()
            browser.close()
            return html
    except Exception as e:
        log.error("Playwright hatası: %s", e)
        return ""


def _looks_login_walled(html: str) -> bool:
    text = _norm(BeautifulSoup(html, "lxml").get_text(" ", strip=True)[:4000])
    return any(_norm(sign) in text for sign in LOGIN_WALL_SIGNS)


def discover_participant_url(homepage_url: str) -> str | None:
    """Bir fuar organizatörünün ana sayfasında katılımcı/exhibitor listesine
    giden linki otomatik arar (metin + href anahtar kelime eşleşmesiyle).
    Kullanıcı 'Fuar Ekle'ye direkt liste URL'si yerine ana sayfa yapıştırırsa
    devreye girer. Bulamazsa None döner."""
    html = _fetch(homepage_url)
    if not html:
        return None
    soup = BeautifulSoup(html, "lxml")
    best_href, best_score = None, 0
    for a in soup.find_all("a", href=True):
        text = _norm(a.get_text(" ", strip=True))
        href_n = _norm(a["href"])
        score = 0
        for kw in PARTICIPANT_LINK_KEYWORDS:
            kwn = _norm(kw)
            if kwn in text:
                score = max(score, 100 - min(len(text), 90))
            if kwn.replace(" ", "") in href_n.replace("-", "").replace("_", ""):
                score = max(score, 60)
        if score > best_score:
            best_score, best_href = score, a["href"]
    if not best_href:
        return None
    return _urlparse.urljoin(homepage_url, best_href)


def _quality_filter(names: list[str]) -> list[str]:
    """Geçerli firma adlarını süzer; sonuç çok küçükse ya da aşırı tekrarlıysa
    (yanlış seçici → aynı metin defalarca) tüm adayı reddeder."""
    names = [n for n in names if is_valid_company(n)]
    if len(names) < 10:
        return []
    unique_ratio = len({_norm(n) for n in names}) / len(names)
    if unique_ratio < 0.5:
        return []
    return names


def _extract_candidates(soup: BeautifulSoup) -> list[tuple[str, list[str]]]:
    candidates: list[tuple[str, list[str]]] = []

    # Strateji 1: menü class'ı OLMAYAN, >10 li içeren UL
    for ul in soup.find_all("ul"):
        if set(ul.get("class", [])) & MENU_CLASSES:
            continue
        lis = ul.find_all("li", recursive=False)
        if len(lis) < 10:
            continue
        names = _quality_filter([li.get_text(" ", strip=True) for li in lis])
        if names:
            candidates.append((f"ul[{len(lis)} li]", names))

    # Strateji 2: exhibitor / participant / company / katilimci / firm / brand class'ları
    for sel in ["[class*='exhibitor']", "[class*='participant']",
                "[class*='company']", "[class*='katilimci']", "[class*='firm']",
                "[class*='brand']"]:
        names = []
        for el in soup.select(sel):
            name_el = el.find(["h1", "h2", "h3", "h4", "strong", "b"]) or el
            names.append(name_el.get_text(" ", strip=True)[:120])
        names = _quality_filter(names)
        if names:
            candidates.append((f"selector {sel}", names))

    # Strateji 3: tablolar
    for table in soup.find_all("table"):
        names = []
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if tds:
                names.append(tds[0].get_text(" ", strip=True))
        names = _quality_filter(names)
        if names:
            candidates.append((f"table[{len(table.find_all('tr'))} rows]", names))

    # Strateji 4: tekrar eden div/article/li "kart" blokları (grid/kart düzeni
    # — WordPress/React sitelerinde ul/li veya table olmadan da liste olabilir)
    BLOCK_CLASS_EXCLUDE = {"flag", "flags", "country", "countries", "ulke", "ulkeler"}
    groups: dict[tuple, list] = {}
    for el in soup.find_all(["div", "article", "li"]):
        classes = el.get("class")
        if not classes:
            continue
        if {_norm(c) for c in classes} & BLOCK_CLASS_EXCLUDE:
            continue  # ör. "ziyaretçi ülkeleri" bayrak grid'i — firma listesi değil
        sig = (el.name, tuple(sorted(classes)))
        group = groups.setdefault(sig, [])
        if len(group) <= 2000:
            group.append(el)
    for sig, elements in groups.items():
        if len(elements) < 10 or len(elements) > 2000:
            continue
        names = []
        for el in elements:
            name_el = el.find(["h1", "h2", "h3", "h4", "h5", "strong", "b"])
            text = name_el.get_text(" ", strip=True) if name_el else el.get_text(" ", strip=True)[:120]
            names.append(text)
        names = _quality_filter(names)
        if names:
            candidates.append((f"blocks {sig[0]}.{'.'.join(sig[1])[:30]}", names))

    # Strateji 5: "öne çıkan katılımcılar" pazarlama metni formatı — her firma
    # kendi <p>'sinde, adı paragrafın en başında kalın (<strong>/<b>) yazıyor,
    # ardından düz yazı tanıtım metni geliyor (ör. "<p><strong>Tamek</strong>,
    # gıda ve içecek sektöründe...</p>"). Bunu ul/table/kart stratejileri
    # yakalayamıyor çünkü tekrar eden bir class/tag yapısı yok, sadece paragraf
    # başındaki kalın metin deseni tutarlı.
    names = []
    for p in soup.find_all("p"):
        first_tag = p.find(True)
        if first_tag and first_tag.name in ("strong", "b"):
            names.append(first_tag.get_text(" ", strip=True))
    names = _quality_filter(names)
    if names:
        candidates.append((f"p>strong lead [{len(names)}]", names))

    return candidates


def _pagination_candidates(url: str, max_pages: int):
    if PAGE_PARAM_RE.search(url):
        for n in range(2, max_pages + 1):
            yield PAGE_PARAM_RE.sub(rf"\g<1>{n}", url, count=1)
    else:
        sep = "&" if "?" in url else "?"
        for n in range(2, max_pages + 1):
            yield f"{url}{sep}page={n}"


def _scrape_tuyap_style(url: str, fair_name: str, max_pages: int) -> list[dict]:
    """Tüyap-platformu sitelerinin (intermobistanbul.com, madenturkiyefuari.com
    vb.) hepsi aynı yapıyı kullanıyor — doğrulanmış tuyap_platform_scrape.parse_page
    parser'ını burada da kullan, tekerleği yeniden icat etme."""
    from tuyap_platform_scrape import parse_page  # döngüsel import'tan kaçınmak için gecikmeli

    base_domain = url.split("/katilimci", 1)[0].rstrip("/") if "/katilimci" in url else url.rstrip("/")
    all_companies: list[dict] = []
    seen: set[str] = set()
    empty_streak = 0
    page = 1
    while page <= max_pages:
        page_url = url if page == 1 else (
            PAGE_PARAM_RE.sub(rf"\g<1>{page}", url, count=1) if PAGE_PARAM_RE.search(url)
            else f"{url}{'&' if '?' in url else '?'}page={page}"
        )
        html = _fetch(page_url)
        if not html:
            break
        page_companies = parse_page(html, base_domain, fair_name)
        new = [c for c in page_companies if _norm(c["name"]) not in seen]
        if not new:
            empty_streak += 1
            if empty_streak >= 2:
                break
        else:
            empty_streak = 0
            for c in new:
                seen.add(_norm(c["name"]))
                all_companies.append(c)
        page += 1
        time.sleep(0.8)
    log.info("Tüyap-CMS stili tarama: %d sayfa, %d firma", page - 1, len(all_companies))
    return all_companies


def generic_scrape(url: str, fair_name: str = "Custom Fair", max_pages: int = 60) -> list[dict]:
    """
    Verilen URL'den katılımcı/firma listesi çıkarır. Önce requests dener,
    aday bulamazsa Playwright'a geçer; çoklu strateji + otomatik sayfalama +
    login-duvarı tespiti içerir (bkz. modül başlığındaki özet).
    """
    log.info("Generic scrape: %s", url)

    html = _fetch(url)
    used_playwright = False
    if len(html) < 5000 or "<html" not in html.lower():
        log.info("İçerik az (%d byte), Playwright deneniyor", len(html))
        html2 = _fetch_playwright(url)
        if html2:
            html, used_playwright = html2, True

    if not html:
        log.error("Sayfa hiç yüklenemedi: %s", url)
        return []

    if _looks_login_walled(html):
        log.warning("UYARI: sayfa login/üyelik duvarı arkasında görünüyor, gerçek liste "
                     "görülemiyor: %s", url)
        return []

    soup = BeautifulSoup(html, "lxml")

    # Tüyap-CMS hızlı yol: bu platformdaki onlarca fuar sitesi aynı şablonu kullanıyor
    if soup.select_one("div.brand-item h2.brand-name"):
        return _scrape_tuyap_style(url, fair_name, max_pages)

    candidates = _extract_candidates(soup)

    if not candidates and not used_playwright:
        log.info("Statik HTML'de aday bulunamadı, Playwright ile tekrar deneniyor")
        html2 = _fetch_playwright(url)
        if html2 and not _looks_login_walled(html2):
            soup2 = BeautifulSoup(html2, "lxml")
            cands2 = _extract_candidates(soup2)
            if cands2:
                candidates = cands2
        elif html2:
            log.warning("UYARI: sayfa login/üyelik duvarı arkasında görünüyor: %s", url)
            return []

    if not candidates:
        log.warning("Hiçbir strateji 10+ geçerli kayıt bulamadı: %s", url)
        return []

    best_label, best_names = max(candidates, key=lambda x: len(x[1]))
    log.info("Seçilen strateji: %s → %d firma (sayfa 1)", best_label, len(best_names))

    all_names = list(best_names)
    seen = {_norm(n) for n in all_names}

    empty_streak = 0
    for page_url in _pagination_candidates(url, max_pages):
        html_p = _fetch(page_url)
        if not html_p:
            break
        soup_p = BeautifulSoup(html_p, "lxml")
        cands_p = _extract_candidates(soup_p)
        new_names = []
        if cands_p:
            _, names_p = max(cands_p, key=lambda x: len(x[1]))
            new_names = [n for n in names_p if _norm(n) not in seen]
        if not new_names:
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0
        for n in new_names:
            seen.add(_norm(n))
            all_names.append(n)
        time.sleep(0.6)

    if len(all_names) > len(best_names):
        log.info("Sayfalama ile toplam %d → %d firmaya çıktı", len(best_names), len(all_names))

    companies = []
    seen2: set[str] = set()
    for name in all_names:
        key = _norm(name)
        if key and key not in seen2:
            seen2.add(key)
            companies.append(make_company(name, fair=fair_name))
    return companies


def deduplicate(companies: list[dict]) -> list[dict]:
    seen = set()
    result = []
    for c in companies:
        key = c["name"].lower().strip()
        if key not in seen:
            seen.add(key)
            result.append(c)
    return result


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Reyart fuar katılımcı scraper")
    parser.add_argument("--url", help="Tek bir fuar için katılımcı listesi URL'si")
    parser.add_argument("--name", default="Custom Fair", help="Fuar adı")
    parser.add_argument("--append", action="store_true",
                        help="Mevcut companies.json'a ekle (varsayılan: üzerine yaz)")
    args = parser.parse_args()

    if args.url:
        # Tek fuar modu — generic scraper
        companies = generic_scrape(args.url, fair_name=args.name)
        if args.append and COMPANIES_FILE.exists():
            existing = json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
            companies = deduplicate(existing + companies)
        else:
            companies = deduplicate(companies)
    else:
        # Varsayılan mod — KOMATEK + WIN EURASIA
        komatek = scrape_komatek()
        time.sleep(random.uniform(1, 2))
        win = scrape_win_eurasia()
        companies = deduplicate(komatek + win)

    COMPANIES_FILE.write_text(
        json.dumps(companies, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    log.info("Toplam %d firma kaydedildi → %s", len(companies), COMPANIES_FILE)

    log.info("İlk 10 firma:")
    for c in companies[:10]:
        log.info("  [%s] %s", c["priority"], c["name"])


if __name__ == "__main__":
    main()
