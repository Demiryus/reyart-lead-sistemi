"""Tüm network'ü logla — gerçek API endpoint'i bul."""
from playwright.sync_api import sync_playwright

URL = "https://visit.zuchex.com/widget/event/zuchex-2025/exhibitors/RXZlbnRWaWV3XzEwODExNzI=?paginationMode=infinite&source=script&showActions=true&lng=tr-TR"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_context(viewport={"width": 1280, "height": 900}).new_page()

    def on_req(r):
        if any(x in r.url for x in ("graphql", "api", "swapcard")):
            print(">>", r.method, r.url[:200])

    page.on("request", on_req)
    page.goto(URL, wait_until="domcontentloaded")
    page.wait_for_timeout(8000)
    # Scroll
    for _ in range(5):
        page.mouse.wheel(0, 3000)
        page.wait_for_timeout(1500)
    page.wait_for_timeout(3000)
    browser.close()
