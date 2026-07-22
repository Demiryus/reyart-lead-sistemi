"""
Zuchex 2025 katılımcı isimleri scraper.
visit.zuchex.com widget "Load more" butonu ile tüm firmalar çekilir.
Çıktı: output/zuchex_2025_names.txt + output/zuchex_2025.xlsx
"""
import asyncio
import json
from pathlib import Path

from playwright.async_api import async_playwright
from bs4 import BeautifulSoup
import openpyxl

WIDGET_URL = (
    "https://visit.zuchex.com/widget/event/zuchex-2025/exhibitors/"
    "RXZlbnRWaWV3XzEwODExNzI=?paginationMode=infinite&source=script&showActions=true&lng=tr-TR"
)

OUT_DIR = Path(__file__).parent / "output"
OUT_DIR.mkdir(exist_ok=True)

SKIP_NAMES = {
    "ZUCHEX", "Informa Markets", "Katılımcı Listesi", "Katılımcılar", "Filtreler",
    "Ülke", "Ürün Grubu", "Load more",
}


def parse_companies(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    spans = soup.find_all("span")
    companies = []
    seen = set()
    for i, s in enumerate(spans):
        classes = " ".join(s.get("class", []))
        if "fpmvUQ" not in classes:
            continue
        name = s.get_text(strip=True)
        if not name or len(name) < 2 or name in SKIP_NAMES:
            continue
        if name.startswith("Hol ") or name == "Fuaye":
            continue
        stand = ""
        if i + 1 < len(spans):
            next_cls = " ".join(spans[i + 1].get("class", []))
            if "bGoSdc" in next_cls:
                stand = spans[i + 1].get_text(strip=True)
        if name not in seen:
            seen.add(name)
            companies.append({"name": name, "stand": stand, "fair": "ZUCHEX 2025"})
    return companies


async def scrape() -> str:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
            ),
            locale="tr-TR",
        )
        page = await ctx.new_page()

        print("[1/3] Widget yükleniyor...")
        await page.goto(WIDGET_URL, timeout=60000)
        await page.wait_for_timeout(5000)

        print("[2/3] 'Load more' ile tüm firmalar yükleniyor...")
        click_count = 0
        while True:
            load_more = page.locator("button", has_text="Load more").first
            try:
                await load_more.wait_for(state="visible", timeout=4000)
                await load_more.scroll_into_view_if_needed()
                await load_more.click()
                click_count += 1
                await page.wait_for_timeout(1500)

                html = await page.content()
                count = len(parse_companies(html))
                print(f"  Load more #{click_count}: {count} firma yüklendi")
            except Exception:
                print(f"  'Load more' bitti. Toplam {click_count} kez tıklandı.")
                break

        html = await page.content()
        await browser.close()
        return html


def main():
    html = asyncio.run(scrape())

    print("[3/3] Firma isimleri ayrıştırılıyor...")
    companies = parse_companies(html)

    (OUT_DIR / "zuchex_debug.html").write_text(html, encoding="utf-8")

    if not companies:
        print("HATA: Firma bulunamadı. output/zuchex_debug.html dosyasını kontrol et.")
        return

    print(f"\nToplam {len(companies)} firma bulundu.")
    print("İlk 10:")
    for c in companies[:10]:
        print(f"  {c['name']} — Stand: {c['stand']}")

    txt_path = OUT_DIR / "zuchex_2025_names.txt"
    txt_path.write_text("\n".join(c["name"] for c in companies), encoding="utf-8")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Zuchex 2025"
    ws.append(["Firma Adı", "Stand No", "Fuar"])
    for c in companies:
        ws.append([c["name"], c["stand"], c["fair"]])
    xlsx_path = OUT_DIR / "zuchex_2025.xlsx"
    wb.save(xlsx_path)

    print(f"\nTXT : {txt_path}")
    print(f"XLSX: {xlsx_path}")


if __name__ == "__main__":
    main()
